using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Threading;
using Newtonsoft.Json.Linq;
using System;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using VsIdeBridge.Infrastructure;
using VsIdeBridge.Tooling.Build;

namespace VsIdeBridge.Services;

internal static class BuildServiceHelpers
{
    /// <summary>
    /// The Error List does not always receive linker/MSBuild failures (LNK2001/LNK1120 never
    /// populate it), so after a failed build parse the Build pane into structured rows the
    /// model can act on directly instead of trusting a zero-error Error List.
    /// </summary>
    internal static async Task AttachBuildOutputErrorsAsync(IdeCommandContext context, JObject buildResult)
    {
        if (buildResult["succeeded"]?.Value<bool>() != false)
        {
            return;
        }

        string? buildOutput = await context.Runtime.OutputWindowService.TryReadBuildPaneTextAsync(context.Dte).ConfigureAwait(true);
        IReadOnlyList<BuildOutputError> parsed = BuildOutputErrorParser.ParseErrors(buildOutput);
        JArray rows = [];
        foreach (BuildOutputError error in parsed)
        {
            JObject row = new()
            {
                ["severity"] = error.Severity,
                ["code"] = error.Code,
                ["message"] = error.Message,
                ["file"] = error.File,
                ["line"] = error.Line,
                ["column"] = error.Column,
                ["rawLine"] = error.RawLine,
            };

            if (error.HasProjectOrigin)
            {
                row["originFile"] = error.OriginFile;
                row["originLine"] = error.OriginLine;
                row["originColumn"] = error.OriginColumn;
                row["originRawLine"] = error.OriginRawLine;
            }

            rows.Add(row);
        }

        buildResult["buildOutputErrors"] = rows;
        buildResult["buildOutputErrorCount"] = rows.Count;
        buildResult["nextStep"] = rows.Count > 0
            ? "Build failed. buildOutputErrors contains the compiler/linker failures parsed from the Build output pane; the Error List may not show them. Fix those first, or call read_output {pane:\"Build\"} for full context."
            : "Build failed, but no error lines could be parsed from the Build output pane and the Error List may also be empty. Call read_output {pane:\"Build\"} to inspect the raw build log.";
    }
    internal static JObject GetBuildStateCore(DTE2 dte, string solutionNotOpenCode, string noSolutionOpen, string solutionPathKey, string activeConfigurationKey, string activePlatformKey)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (dte.Solution?.IsOpen != true)
        {
            throw new CommandErrorException(solutionNotOpenCode, noSolutionOpen);
        }

        SolutionBuild solutionBuild = dte.Solution.SolutionBuild;
        bool lastBuildInfoKnown = true;
        int lastBuildInfoValue = 0;
        string? lastBuildInfoReason = null;

        try
        {
            lastBuildInfoValue = solutionBuild.LastBuildInfo;
        }
        catch (COMException ex)
        {
            lastBuildInfoKnown = false;
            lastBuildInfoReason = ex.Message;
        }

        JObject buildStatus = new()
        {
            [solutionPathKey] = dte.Solution.FullName,
            [activeConfigurationKey] = solutionBuild.ActiveConfiguration?.Name ?? string.Empty,
            [activePlatformKey] = (solutionBuild.ActiveConfiguration as SolutionConfiguration2)?.PlatformName ?? string.Empty,
            ["buildState"] = solutionBuild.BuildState.ToString(),
            ["lastBuildInfoKnown"] = lastBuildInfoKnown,
            ["lastBuildInfo"] = lastBuildInfoKnown ? lastBuildInfoValue : JValue.CreateNull(),
        };

        if (!string.IsNullOrWhiteSpace(lastBuildInfoReason))
        {
            buildStatus["lastBuildInfoReason"] = lastBuildInfoReason;
        }

        return buildStatus;
    }

    internal static JObject ListConfigurationsCore(DTE2 dte, string solutionNotOpenCode, string noSolutionOpen, string solutionPathKey, string activeConfigurationKey, string activePlatformKey)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (dte.Solution?.IsOpen != true)
        {
            throw new CommandErrorException(solutionNotOpenCode, noSolutionOpen);
        }

        SolutionBuild solutionBuild = dte.Solution.SolutionBuild;
        string activeConfiguration = solutionBuild.ActiveConfiguration?.Name ?? string.Empty;
        string activePlatform = (solutionBuild.ActiveConfiguration as SolutionConfiguration2)?.PlatformName ?? string.Empty;
        JArray items = [];

        foreach (SolutionConfiguration2 item in solutionBuild.SolutionConfigurations)
        {
            items.Add(new JObject
            {
                ["name"] = item.Name ?? string.Empty,
                ["platform"] = item.PlatformName ?? string.Empty,
                ["isActive"] = string.Equals(item.Name, activeConfiguration, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(item.PlatformName ?? string.Empty, activePlatform, StringComparison.OrdinalIgnoreCase),
            });
        }

        return new JObject
        {
            [solutionPathKey] = dte.Solution.FullName,
            [activeConfigurationKey] = activeConfiguration,
            [activePlatformKey] = activePlatform,
            ["count"] = items.Count,
            ["items"] = items,
        };
    }

    internal static string? FindProjectUniqueName(DTE2 dte, string projectName)
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        foreach (Project project in dte.Solution.Projects)
        {
            if (string.Equals(project.Name, projectName, StringComparison.OrdinalIgnoreCase)
                || (project.UniqueName?.EndsWith(projectName, StringComparison.OrdinalIgnoreCase) == true))
            {
                return project.UniqueName;
            }
        }

        return null;
    }

    internal static JObject CreateBuildResult(
        DTE2 dte,
        SolutionBuild solutionBuild,
        DateTimeOffset startedAt,
        string solutionPathKey,
        string activeConfigurationKey,
        string activePlatformKey,
        string operation,
        string? projectName = null,
        string? uniqueName = null)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        JObject result = new()
        {
            ["operation"] = operation,
            [solutionPathKey] = dte.Solution.FullName,
            [activeConfigurationKey] = solutionBuild.ActiveConfiguration?.Name ?? string.Empty,
            [activePlatformKey] = (solutionBuild.ActiveConfiguration as SolutionConfiguration2)?.PlatformName ?? string.Empty,
            ["lastBuildInfo"] = solutionBuild.LastBuildInfo,
            ["succeeded"] = solutionBuild.LastBuildInfo == 0,
            ["elapsedMilliseconds"] = (DateTimeOffset.UtcNow - startedAt).TotalMilliseconds,
        };

        if (!string.IsNullOrWhiteSpace(projectName))
        {
            result["projectName"] = projectName;
        }

        if (!string.IsNullOrWhiteSpace(uniqueName))
        {
            result["projectUniqueName"] = uniqueName;
        }

        return result;
    }

    internal static bool TryActivateConfiguration(SolutionBuild solutionBuild, string? configuration, string? platform, bool requireMatch = false)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (string.IsNullOrWhiteSpace(configuration) && string.IsNullOrWhiteSpace(platform))
        {
            return true;
        }

        foreach (SolutionConfiguration2 item in solutionBuild.SolutionConfigurations)
        {
            if (!string.IsNullOrWhiteSpace(configuration) &&
                !string.Equals(item.Name, configuration, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!string.IsNullOrWhiteSpace(platform) &&
                !string.Equals(item.PlatformName, platform, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            item.Activate();
            return true;
        }

        return !requireMatch;
    }

    internal static async Task WaitForBuildEventAsync(Task buildTask, DateTimeOffset deadline, CancellationToken cancellationToken, string timeoutMessage)
    {
        while (!buildTask.IsCompleted)
        {
            TimeSpan remaining = deadline - DateTimeOffset.UtcNow;
            if (remaining <= TimeSpan.Zero)
            {
                throw new CommandErrorException("timeout", timeoutMessage);
            }

            TimeSpan pollDelay = remaining < TimeSpan.FromMilliseconds(250)
                ? remaining
                : TimeSpan.FromMilliseconds(250);

            await Task.Delay(pollDelay, cancellationToken).ConfigureAwait(false);
        }

        if (buildTask.IsCanceled)
        {
            throw new OperationCanceledException(cancellationToken);
        }

        if (buildTask.IsFaulted)
        {
            ExceptionDispatchInfo.Capture(buildTask.Exception?.GetBaseException() ?? buildTask.Exception!).Throw();
        }
    }

    internal static JObject CreateStartedOperationResult(DateTimeOffset startedAt, string operation)
    {
        return new JObject
        {
            ["operation"] = operation,
            ["status"] = "started",
            ["waitForCompletion"] = false,
            ["startedAtUtc"] = startedAt.ToString("O"),
            ["followUpHint"] = "Prompt the model again when the operation finishes, then read warnings, errors, messages, or diagnostics_snapshot.",
        };
    }

    internal static void ObserveBackgroundOperation(Task backgroundTask, IdeCommandContext context, string operation)
    {
        _ = backgroundTask.ContinueWith(
            antecedent =>
            {
                Exception? failure = antecedent.Exception?.GetBaseException();
                if (failure is null)
                {
                    return;
                }

                string logMessage = failure is CommandErrorException commandError
                    ? $"IDE Bridge: {operation} failed: {commandError.Message}"
                    : $"IDE Bridge: {operation} failed unexpectedly: {failure.Message}";

                ThreadHelper.JoinableTaskFactory.Run(async () =>
                {
                    await context.Logger.LogAsync(logMessage, CancellationToken.None).ConfigureAwait(true);
                });
            },
            CancellationToken.None,
            TaskContinuationOptions.OnlyOnFaulted,
            TaskScheduler.Default);
    }

    internal static async Task RunProjectBuildAsync(
        IdeCommandContext context,
        SolutionBuild solutionBuild,
        IVsSolutionBuildManager2? buildManager,
        string activeConfig,
        string uniqueName,
        DateTimeOffset deadline)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);
        const string timeoutMessage = "Timed out waiting for the build to finish.";
        if (buildManager is not null)
        {
            BuildCompletionWaiter waiter = new(buildManager);
            try
            {
                solutionBuild.BuildProject(activeConfig, uniqueName, false);
                await WaitForBuildEventAsync(waiter.CompletionTask, deadline, context.CancellationToken, timeoutMessage).ConfigureAwait(false);
            }
            finally
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(CancellationToken.None);
                waiter.Unsubscribe();
            }
        }
        else
        {
            solutionBuild.BuildProject(activeConfig, uniqueName, false);
            await BuildService.WaitForBuildCompletionAsync(context, solutionBuild, deadline, timeoutMessage).ConfigureAwait(true);
        }
    }

    internal static async Task RunSolutionBuildStepAsync(
        IdeCommandContext context,
        SolutionBuild solutionBuild,
        IVsSolutionBuildManager2? buildManager,
        bool clean,
        DateTimeOffset deadline,
        string timeoutMessage)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);
        if (buildManager is not null)
        {
            BuildCompletionWaiter waiter = new(buildManager);
            try
            {
                if (clean)
                    solutionBuild.Clean(false);
                else
                    solutionBuild.Build(false);

                await WaitForBuildEventAsync(waiter.CompletionTask, deadline, context.CancellationToken, timeoutMessage).ConfigureAwait(false);
            }
            finally
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(CancellationToken.None);
                waiter.Unsubscribe();
            }
        }
        else
        {
            if (clean)
                solutionBuild.Clean(true);
            else
                solutionBuild.Build(true);

            await BuildService.WaitForBuildCompletionAsync(context, solutionBuild, deadline, timeoutMessage).ConfigureAwait(true);
        }
    }

    internal sealed class BuildCompletionWaiter : IVsUpdateSolutionEvents
    {
        private readonly IVsSolutionBuildManager2 _buildManager;
        private readonly TaskCompletionSource<bool> _tcs = new();
        private uint _cookie;

        internal Task CompletionTask => _tcs.Task;
        internal bool IsCompleted => _tcs.Task.IsCompleted;

        internal int LastBuildInfo { get; private set; }

        internal BuildCompletionWaiter(IVsSolutionBuildManager2 buildManager)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _buildManager = buildManager;
            _buildManager.AdviseUpdateSolutionEvents(this, out _cookie);
        }

        internal void Unsubscribe()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_cookie != 0)
            {
                _buildManager.UnadviseUpdateSolutionEvents(_cookie);
                _cookie = 0;
            }

            _tcs.TrySetCanceled();
        }

        int IVsUpdateSolutionEvents.UpdateSolution_Begin(ref int pfCancelUpdate) => VSConstants.S_OK;

        int IVsUpdateSolutionEvents.UpdateSolution_Done(int fSucceeded, int fModified, int fCancelCommand)
        {
            LastBuildInfo = fSucceeded != 0 ? 0 : 1;
            _tcs.TrySetResult(true);
            return VSConstants.S_OK;
        }

        int IVsUpdateSolutionEvents.UpdateSolution_StartUpdate(ref int pfCancelUpdate) => VSConstants.S_OK;

        int IVsUpdateSolutionEvents.UpdateSolution_Cancel()
        {
            _tcs.TrySetResult(true);
            return VSConstants.S_OK;
        }

        int IVsUpdateSolutionEvents.OnActiveProjectCfgChange(IVsHierarchy pIVsHierarchy) => VSConstants.S_OK;
    }
}
