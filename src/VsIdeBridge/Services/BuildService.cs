using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;
using Newtonsoft.Json.Linq;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed class BuildService(ReadinessService readinessService)
{
    private const int DefaultBuildTimeoutMilliseconds = 600_000;
    private const int BuildPollIntervalMilliseconds = 500;

    private const string SolutionNotOpenCode = "solution_not_open";
    private const string NoSolutionOpen = "No solution is open. Call open_solution with a .sln or .slnx path to open one, or call bind_solution if the solution is already loaded in a different VS instance.";
    private const string SolutionPathKey = "solutionPath";
    private const string ActiveConfigurationKey = "activeConfiguration";
    private const string ActivePlatformKey = "activePlatform";

    private readonly ReadinessService _readinessService = readinessService;
    private readonly BuildBackgroundOperationTracker _backgroundOperations = new();

    public async Task<JObject> StartBuildSolutionAsync(IdeCommandContext context, int timeoutMilliseconds, string? configuration, string? platform)
    {
        return await QueueBackgroundBuildOperationAsync(
            context,
            "build",
            detached => BuildSolutionAsync(detached, timeoutMilliseconds, configuration, platform)).ConfigureAwait(true);
    }

    public async Task<JObject> StartRebuildSolutionAsync(IdeCommandContext context, int timeoutMilliseconds, string? configuration, string? platform)
    {
        return await QueueBackgroundBuildOperationAsync(
            context,
            "rebuild",
            detached => RebuildSolutionAsync(detached, timeoutMilliseconds, configuration, platform)).ConfigureAwait(true);
    }

    public async Task<JObject> StartCodeAnalysisAsync(IdeCommandContext context, int timeoutMilliseconds)
    {
        return await QueueBackgroundBuildOperationAsync(
            context,
            "code-analysis",
            detached => RunCodeAnalysisAsync(detached, timeoutMilliseconds)).ConfigureAwait(true);
    }

    public async Task<JObject> BuildSolutionAsync(IdeCommandContext context, int timeoutMilliseconds, string? configuration, string? platform)
    {
        (DTE2 dte, SolutionBuild solutionBuild) = await PrepareBuildAsync(context, configuration, platform).ConfigureAwait(true);
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);

        DateTimeOffset startedAt = DateTimeOffset.UtcNow;
        DateTimeOffset deadline = DateTimeOffset.UtcNow.AddMilliseconds(timeoutMilliseconds <= 0 ? DefaultBuildTimeoutMilliseconds : timeoutMilliseconds);
        string solutionName = dte.Solution.FullName;

        // GetServiceAsync does not require the main thread — release it here.
        IVsSolutionBuildManager2? buildManager = await context.Package.GetServiceAsync(typeof(SVsSolutionBuildManager)).ConfigureAwait(false) as IVsSolutionBuildManager2;

        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);
        await context.Logger.LogAsync($"IDE Bridge: build starting ({solutionName})", context.CancellationToken).ConfigureAwait(true);
        await BuildServiceHelpers.RunSolutionBuildStepAsync(context, solutionBuild, buildManager, clean: false, deadline, "Timed out waiting for the build to finish.").ConfigureAwait(true);

        double elapsed = (DateTimeOffset.UtcNow - startedAt).TotalMilliseconds;
        bool succeeded = solutionBuild.LastBuildInfo == 0;
        await context.Logger.LogAsync(
            $"IDE Bridge: build {(succeeded ? "succeeded" : "failed")} in {elapsed:0}ms",
            context.CancellationToken).ConfigureAwait(true);

        JObject buildResult = CreateBuildResult(dte, solutionBuild, startedAt, operation: "build");
        await BuildServiceHelpers.AttachBuildOutputErrorsAsync(context, buildResult).ConfigureAwait(true);
        return buildResult;
    }

    public async Task<JObject> RebuildSolutionAsync(IdeCommandContext context, int timeoutMilliseconds, string? configuration, string? platform)
    {
        (DTE2 dte, SolutionBuild solutionBuild) = await PrepareBuildAsync(context, configuration, platform).ConfigureAwait(true);
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);

        DateTimeOffset startedAt = DateTimeOffset.UtcNow;
        DateTimeOffset deadline = DateTimeOffset.UtcNow.AddMilliseconds(timeoutMilliseconds <= 0 ? DefaultBuildTimeoutMilliseconds : timeoutMilliseconds);
        string solutionName = dte.Solution.FullName;

        // GetServiceAsync does not require the main thread — release it here.
        IVsSolutionBuildManager2? buildManager = await context.Package.GetServiceAsync(typeof(SVsSolutionBuildManager)).ConfigureAwait(false) as IVsSolutionBuildManager2;

        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);
        await context.Logger.LogAsync($"IDE Bridge: rebuild starting ({solutionName})", context.CancellationToken).ConfigureAwait(true);
        await BuildServiceHelpers.RunSolutionBuildStepAsync(context, solutionBuild, buildManager, clean: true, deadline, "Timed out waiting for the clean step to finish.").ConfigureAwait(true);
        await BuildServiceHelpers.RunSolutionBuildStepAsync(context, solutionBuild, buildManager, clean: false, deadline, "Timed out waiting for the rebuild to finish.").ConfigureAwait(true);

        double elapsed = (DateTimeOffset.UtcNow - startedAt).TotalMilliseconds;
        bool succeeded = solutionBuild.LastBuildInfo == 0;
        await context.Logger.LogAsync(
            $"IDE Bridge: rebuild {(succeeded ? "succeeded" : "failed")} in {elapsed:0}ms",
            context.CancellationToken).ConfigureAwait(true);

        JObject rebuildResult = CreateBuildResult(dte, solutionBuild, startedAt, operation: "rebuild");
        await BuildServiceHelpers.AttachBuildOutputErrorsAsync(context, rebuildResult).ConfigureAwait(true);
        return rebuildResult;
    }

    public async Task<JObject> GetBuildStateAsync(DTE2 dte)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        JObject buildStatus = GetBuildStateCore(dte);
        if (_backgroundOperations.GetSnapshot() is { } backgroundOperation)
        {
            buildStatus["backgroundOperation"] = backgroundOperation;
        }

        return buildStatus;
    }

    private static JObject GetBuildStateCore(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (dte.Solution?.IsOpen != true)
        {
            throw new CommandErrorException(SolutionNotOpenCode, NoSolutionOpen);
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
            [SolutionPathKey] = dte.Solution.FullName,
            [ActiveConfigurationKey] = solutionBuild.ActiveConfiguration?.Name ?? string.Empty,
            [ActivePlatformKey] = (solutionBuild.ActiveConfiguration as SolutionConfiguration2)?.PlatformName ?? string.Empty,
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

    public async Task<JObject> ListConfigurationsAsync(DTE2 dte)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        return ListConfigurationsCore(dte);
    }

    private static JObject ListConfigurationsCore(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (dte.Solution?.IsOpen != true)
        {
            throw new CommandErrorException(SolutionNotOpenCode, NoSolutionOpen);
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
            [SolutionPathKey] = dte.Solution.FullName,
            [ActiveConfigurationKey] = activeConfiguration,
            [ActivePlatformKey] = activePlatform,
            ["count"] = items.Count,
            ["items"] = items,
        };
    }

    public async Task<JObject> SetConfigurationAsync(DTE2 dte, string configuration, string? platform)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        if (dte.Solution?.IsOpen != true)
        {
            throw new CommandErrorException(SolutionNotOpenCode, NoSolutionOpen);
        }

        SolutionBuild solutionBuild = dte.Solution.SolutionBuild;
        bool activated = TryActivateConfiguration(solutionBuild, configuration, platform, requireMatch: true);
        if (!activated)
        {
            throw new CommandErrorException(
                "build_configuration_not_found",
                $"Configuration '{configuration}' with platform '{platform ?? "<any>"}' was not found.");
        }

        return new JObject
        {
            [SolutionPathKey] = dte.Solution.FullName,
            [ActiveConfigurationKey] = solutionBuild.ActiveConfiguration?.Name ?? string.Empty,
            [ActivePlatformKey] = (solutionBuild.ActiveConfiguration as SolutionConfiguration2)?.PlatformName ?? string.Empty,
        };
    }

    public async Task<JObject> RunCodeAnalysisAsync(IdeCommandContext context, int timeoutMilliseconds)
    {
        // Background-thread work first — these don't need the VS main thread.
        DateTimeOffset startedAt = DateTimeOffset.UtcNow;
        DateTimeOffset deadline = DateTimeOffset.UtcNow.AddMilliseconds(timeoutMilliseconds <= 0 ? DefaultBuildTimeoutMilliseconds : timeoutMilliseconds);
        object? buildManagerObj = await context.Package.GetServiceAsync(typeof(SVsSolutionBuildManager)).ConfigureAwait(false);

        // Narrow main-thread scope to only the DTE/COM access + logging.
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);
        IVsSolutionBuildManager2? buildManager = buildManagerObj as IVsSolutionBuildManager2;
        DTE2 dte = context.Dte;
        if (dte.Solution?.IsOpen != true)
            throw new CommandErrorException(SolutionNotOpenCode, NoSolutionOpen);
        SolutionBuild solutionBuild = dte.Solution.SolutionBuild;
        if (solutionBuild.BuildState == vsBuildState.vsBuildStateInProgress)
            throw new CommandErrorException("build_in_progress", "A build or code analysis is already in progress. Wait for it to complete, then retry.");
        string solutionName = dte.Solution.FullName;
        string activeConfig = solutionBuild.ActiveConfiguration?.Name ?? string.Empty;
        string activePlatform = (solutionBuild.ActiveConfiguration as SolutionConfiguration2)?.PlatformName ?? string.Empty;
        await context.Logger.LogAsync($"IDE Bridge: code analysis starting ({solutionName})", context.CancellationToken).ConfigureAwait(true);

        // The JoinableTask defers ExecuteCommand to the next message-pump cycle via Task.Yield()
        // so ConfigureAwait(false) on WaitForBuildEventAsync can release the main thread first.
        // ExecuteCommand pumps the Windows message queue while blocked, so the Done event and the
        // finally's SwitchToMainThreadAsync both run inside it — no deadlock.
        const string timeoutMessage = "Timed out waiting for code analysis to finish. Increase timeout_ms and retry, or use wait_for_completion: false to start analysis without waiting.";
        int lastBuildInfo;
        if (buildManager is not null)
        {
            BuildCompletionWaiter waiter = new(buildManager);
            using CancellationTokenSource launchCts = new();
            JoinableTask executeTask = ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                await Task.Yield();
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(CancellationToken.None);
                try { dte.ExecuteCommand("Build.RunCodeAnalysisonSolution"); }
                catch (COMException ex)
                {
                    BridgeActivityLog.LogWarning(nameof(BuildService), "Failed to start solution code analysis command", ex);
                    launchCts.Cancel();
                }
            });
            try
            {
                using CancellationTokenSource linkedCts = CancellationTokenSource.CreateLinkedTokenSource(context.CancellationToken, launchCts.Token);
                await BuildServiceHelpers.WaitForBuildEventAsync(waiter.CompletionTask, deadline, linkedCts.Token, timeoutMessage).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (launchCts.IsCancellationRequested && !context.CancellationToken.IsCancellationRequested)
            {
                await executeTask;
                throw new CommandErrorException("analysis_launch_failed", "Code analysis could not be started. Check the build output for details.");
            }
            finally
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(CancellationToken.None);
                waiter.Unsubscribe();
            }
            await executeTask;
            lastBuildInfo = waiter.LastBuildInfo;
        }
        else
        {
            dte.ExecuteCommand("Build.RunCodeAnalysisonSolution");
            await WaitForBuildCompletionAsync(context, solutionBuild, deadline, timeoutMessage).ConfigureAwait(true);
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(CancellationToken.None);
            lastBuildInfo = solutionBuild.LastBuildInfo;
        }

        double elapsed = (DateTimeOffset.UtcNow - startedAt).TotalMilliseconds;
        bool succeeded = lastBuildInfo == 0;
        return new JObject
        {
            ["operation"] = "code-analysis",
            [SolutionPathKey] = solutionName,
            [ActiveConfigurationKey] = activeConfig,
            [ActivePlatformKey] = activePlatform,
            ["lastBuildInfo"] = lastBuildInfo,
            ["succeeded"] = succeeded,
            ["elapsedMilliseconds"] = elapsed,
        };
    }

    public async Task<JObject> BuildProjectAsync(IdeCommandContext context, int timeoutMilliseconds, string projectName, string? configuration, string? platform)
    {
        (DTE2 dte, SolutionBuild solutionBuild) = await PrepareBuildAsync(context, configuration, platform).ConfigureAwait(true);
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);

        string uniqueName = FindProjectUniqueName(dte, projectName)
            ?? throw new CommandErrorException("project_not_found", $"Project '{projectName}' was not found in the solution. Call list_projects to see all project names, then retry with the exact name.");
        string activeConfig = solutionBuild.ActiveConfiguration?.Name ?? "Debug";

        // GetServiceAsync does not require the main thread — release it here.
        IVsSolutionBuildManager2? buildManager = await context.Package.GetServiceAsync(typeof(SVsSolutionBuildManager)).ConfigureAwait(false) as IVsSolutionBuildManager2;

        DateTimeOffset startedAt = DateTimeOffset.UtcNow;
        DateTimeOffset deadline = DateTimeOffset.UtcNow.AddMilliseconds(timeoutMilliseconds <= 0 ? DefaultBuildTimeoutMilliseconds : timeoutMilliseconds);
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);
        await context.Logger.LogAsync($"IDE Bridge: building project '{uniqueName}'", context.CancellationToken).ConfigureAwait(true);
        await RunProjectBuildAsync(context, solutionBuild, buildManager, activeConfig, uniqueName, deadline).ConfigureAwait(true);

        double elapsed = (DateTimeOffset.UtcNow - startedAt).TotalMilliseconds;
        bool succeeded = solutionBuild.LastBuildInfo == 0;
        await context.Logger.LogAsync(
            $"IDE Bridge: project build {(succeeded ? "succeeded" : "failed")} in {elapsed:0}ms",
            context.CancellationToken).ConfigureAwait(true);

        return CreateBuildResult(dte, solutionBuild, startedAt, operation: "build", projectName, uniqueName);
    }

    private static async Task RunProjectBuildAsync(
        IdeCommandContext context,
        SolutionBuild solutionBuild,
        IVsSolutionBuildManager2? buildManager,
        string activeConfig,
        string uniqueName,
        DateTimeOffset deadline)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);
        const string timeoutMessage = "Timed out waiting for the build to finish. Increase timeout_ms and retry.";
        if (buildManager is not null)
        {
            BuildCompletionWaiter waiter = new(buildManager);
            try
            {
                solutionBuild.BuildProject(activeConfig, uniqueName, false);
                await BuildServiceHelpers.WaitForBuildEventAsync(waiter.CompletionTask, deadline, context.CancellationToken, timeoutMessage).ConfigureAwait(false);
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
            await WaitForBuildCompletionAsync(context, solutionBuild, deadline, timeoutMessage).ConfigureAwait(true);
        }
    }

    private static async Task RunSolutionBuildStepAsync(
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
                await BuildServiceHelpers.WaitForBuildEventAsync(waiter.CompletionTask, deadline, context.CancellationToken, timeoutMessage).ConfigureAwait(false);
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
            await WaitForBuildCompletionAsync(context, solutionBuild, deadline, timeoutMessage).ConfigureAwait(true);
        }
    }

    private static string? FindProjectUniqueName(DTE2 dte, string projectName)
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        foreach (EnvDTE.Project project in dte.Solution.Projects)
        {
            string? found = FindUniqueNameInProjectTree(project, projectName);
            if (found is not null)
            {
                return found;
            }
        }
        return null;
    }

    // Recursively walks solution folders, matching by display name, unique name, or full path.
    private static string? FindUniqueNameInProjectTree(EnvDTE.Project project, string projectName)
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        if (string.Equals(project.Kind, ProjectKinds.vsProjectKindSolutionFolder, StringComparison.OrdinalIgnoreCase))
        {
            foreach (EnvDTE.ProjectItem item in project.ProjectItems)
            {
                if (item.SubProject is { } sub)
                {
                    string? found = FindUniqueNameInProjectTree(sub, projectName);
                    if (found is not null)
                    {
                        return found;
                    }
                }
            }
            return null;
        }

        if (string.Equals(project.Name, projectName, StringComparison.OrdinalIgnoreCase)
            || string.Equals(project.UniqueName, projectName, StringComparison.OrdinalIgnoreCase)
            || (project.FullName is { Length: > 0 } && PathNormalization.AreEquivalent(project.FullName, projectName)))
        {
            return project.UniqueName;
        }

        return null;
    }

    public async Task<JObject> BuildAndCaptureErrorsAsync(IdeCommandContext context, int timeoutMilliseconds, bool waitForIntellisense)
    {
        JObject build = await BuildSolutionAsync(context, timeoutMilliseconds, null, null).ConfigureAwait(true);
        if (waitForIntellisense)
        {
            build["readiness"] = await _readinessService.WaitForReadyAsync(context, timeoutMilliseconds).ConfigureAwait(true);
        }

        return build;
    }

    private async Task<JObject> QueueBackgroundBuildOperationAsync(IdeCommandContext context, string operation, Func<IdeCommandContext, Task<JObject>> operationAsync)
    {
        IdeCommandContext detached = new(context.Package, context.Dte, context.Logger, context.Runtime, CancellationToken.None);
        DateTimeOffset startedAt = DateTimeOffset.UtcNow;
        string startedAtUtc = startedAt.ToString("O");

        _backgroundOperations.MarkQueued(operation, startedAt);

        JoinableTask<JObject> backgroundTask = context.Package.JoinableTaskFactory.RunAsync(async () =>
        {
            await Task.Yield();

            _backgroundOperations.MarkRunning(operation, startedAtUtc);
            return await operationAsync(detached).ConfigureAwait(true);
        });

        _ = backgroundTask.Task.ContinueWith(task =>
        {
            if (task.Status == TaskStatus.RanToCompletion)
            {
                _backgroundOperations.Complete(operation, startedAtUtc, task.Result);
                return;
            }

            Exception failure = task.Exception?.GetBaseException()
                ?? new OperationCanceledException($"The background '{operation}' operation did not complete.");
            _backgroundOperations.Fail(operation, startedAtUtc, failure);
        },
        CancellationToken.None,
        TaskContinuationOptions.ExecuteSynchronously,
        TaskScheduler.Default);

        BuildServiceHelpers.ObserveBackgroundOperation(backgroundTask.Task, detached, operation);
        JObject startedResult = BuildServiceHelpers.CreateStartedOperationResult(startedAt, operation);
        if (_backgroundOperations.GetSnapshot() is { } backgroundOperation)
        {
            startedResult["backgroundOperation"] = backgroundOperation;
        }

        return startedResult;
    }

    private static async Task<(DTE2 Dte, SolutionBuild SolutionBuild)> PrepareBuildAsync(IdeCommandContext context, string? configuration, string? platform)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);

        DTE2 dte = context.Dte;
        if (dte.Solution?.IsOpen != true)
        {
            throw new CommandErrorException(SolutionNotOpenCode, NoSolutionOpen);
        }

        SolutionBuild solutionBuild = dte.Solution.SolutionBuild;
        if (solutionBuild.BuildState == vsBuildState.vsBuildStateInProgress)
        {
            throw new CommandErrorException("build_in_progress", "A build is already in progress. Wait for it to complete, then retry.");
        }

        TryActivateConfiguration(solutionBuild, configuration, platform);
        return (dte, solutionBuild);
    }

    internal static async Task WaitForBuildCompletionAsync(
        IdeCommandContext context,
        SolutionBuild solutionBuild,
        DateTimeOffset deadline,
        string timeoutMessage)
    {
        while (true)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);
            if (solutionBuild.BuildState != vsBuildState.vsBuildStateInProgress)
            {
                return;
            }

            if (DateTimeOffset.UtcNow >= deadline)
            {
                throw new CommandErrorException("timeout", timeoutMessage);
            }

            await Task.Delay(BuildPollIntervalMilliseconds, context.CancellationToken).ConfigureAwait(false);
        }
    }

    private static JObject CreateBuildResult(
        DTE2 dte,
        SolutionBuild solutionBuild,
        DateTimeOffset startedAt,
        string operation,
        string? projectName = null,
        string? uniqueName = null)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        JObject result = new()
        {
            ["operation"] = operation,
            [SolutionPathKey] = dte.Solution.FullName,
            [ActiveConfigurationKey] = solutionBuild.ActiveConfiguration?.Name ?? string.Empty,
            [ActivePlatformKey] = (solutionBuild.ActiveConfiguration as SolutionConfiguration2)?.PlatformName ?? string.Empty,
            ["lastBuildInfo"] = solutionBuild.LastBuildInfo,
            ["succeeded"] = solutionBuild.LastBuildInfo == 0,
            ["elapsedMilliseconds"] = (DateTimeOffset.UtcNow - startedAt).TotalMilliseconds,
        };

        if (solutionBuild.LastBuildInfo != 0)
        {
            result["nextStep"] = "Build failed. Call errors to see the full Error List, or build_errors to see the raw build output with error details.";
        }

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

    private static bool TryActivateConfiguration(SolutionBuild solutionBuild, string? configuration, string? platform, bool requireMatch = false)
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

    private static async Task WaitForBuildEventAsync(
        BuildCompletionWaiter waiter,
        DateTimeOffset deadline,
        CancellationToken cancellationToken,
        string timeoutMessage)
    {
        TimeSpan remaining = deadline - DateTimeOffset.UtcNow;
        if (remaining <= TimeSpan.Zero)
        {
            throw new CommandErrorException("timeout", timeoutMessage);
        }

        Task buildTask = waiter.CompletionTask;
        Task delayTask = Task.Delay((int)Math.Min(remaining.TotalMilliseconds, int.MaxValue), cancellationToken);
        Task first = await Task.WhenAny(buildTask, delayTask).ConfigureAwait(false);

        if (first == buildTask)
        {
            return;
        }

        cancellationToken.ThrowIfCancellationRequested();
        throw new CommandErrorException("timeout", timeoutMessage);
    }

    private static JObject CreateStartedOperationResult(DateTimeOffset startedAt, string operation)
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

}
