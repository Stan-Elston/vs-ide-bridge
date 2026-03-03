using System;
using System.Threading.Tasks;
using Microsoft.VisualStudio.OperationProgress;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Newtonsoft.Json.Linq;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed class ReadinessService
{
    private const int PollIntervalMilliseconds = 500;
    private const int StableStatusBarSampleCount = 2;

    public async Task<JObject> WaitForReadyAsync(IdeCommandContext context, int timeoutMilliseconds)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);

        if (context.Dte.Solution is null || !context.Dte.Solution.IsOpen)
        {
            throw new CommandErrorException("solution_not_open", "No solution is open.");
        }

        var startedAt = DateTimeOffset.UtcNow;
        var timeout = TimeSpan.FromMilliseconds(timeoutMilliseconds <= 0 ? 120000 : timeoutMilliseconds);
        var service = await context.Package.GetServiceAsync(typeof(SVsOperationProgressStatusService)).ConfigureAwait(true) as IVsOperationProgressStatusService;
        var stage = service?.GetStageStatusForSolutionLoad(CommonOperationProgressStageIds.Intellisense);
        var statusbar = await context.Package.GetServiceAsync(typeof(SVsStatusbar)).ConfigureAwait(true) as IVsStatusbar;
        var deadline = startedAt.Add(timeout);
        var readyStatusSamples = 0;
        var lastStatusBarText = string.Empty;
        var statusBarReady = false;
        var intellisenseCompleted = stage is not null && !stage.IsInProgress;
        var satisfiedBy = intellisenseCompleted ? "intellisense" : "pending";

        while (DateTimeOffset.UtcNow < deadline)
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);

            if (stage is not null)
            {
                intellisenseCompleted = !stage.IsInProgress;
                if (intellisenseCompleted)
                {
                    satisfiedBy = "intellisense";
                    break;
                }
            }

            lastStatusBarText = TryGetStatusBarText(statusbar);
            statusBarReady = IsReadyStatusText(lastStatusBarText);
            readyStatusSamples = statusBarReady ? readyStatusSamples + 1 : 0;
            if (readyStatusSamples >= StableStatusBarSampleCount)
            {
                satisfiedBy = "status-bar";
                break;
            }

            await Task.Delay(PollIntervalMilliseconds, context.CancellationToken).ConfigureAwait(false);
        }

        var timedOut = satisfiedBy == "pending";
        if (timedOut)
        {
            satisfiedBy = "timeout";
        }

        return new JObject
        {
            ["solutionPath"] = context.Dte.Solution.FullName,
            ["serviceAvailable"] = service is not null,
            ["intellisenseStageAvailable"] = stage is not null,
            ["intellisenseCompleted"] = intellisenseCompleted,
            ["timedOut"] = timedOut,
            ["elapsedMilliseconds"] = (DateTimeOffset.UtcNow - startedAt).TotalMilliseconds,
            ["isInProgress"] = stage?.IsInProgress ?? false,
            ["statusBarAvailable"] = statusbar is not null,
            ["statusBarText"] = lastStatusBarText,
            ["statusBarReady"] = statusBarReady,
            ["readyStatusSamples"] = readyStatusSamples,
            ["satisfiedBy"] = satisfiedBy,
        };
    }

    private static string TryGetStatusBarText(IVsStatusbar? statusbar)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (statusbar is null)
        {
            return string.Empty;
        }

        try
        {
            statusbar.GetText(out var text);
            return text?.Trim() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static bool IsReadyStatusText(string text)
    {
        return string.Equals(text?.Trim(), "Ready", StringComparison.OrdinalIgnoreCase);
    }
}
