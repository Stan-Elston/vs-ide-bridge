using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.Json.Nodes;
using VsIdeBridgeService.SystemTools;

namespace VsIdeBridgeService;

internal static partial class ToolCatalog
{
    private const string MicrosoftVisualStudio = "Microsoft Visual Studio";
    private const string DevenvExe = "devenv.exe";
    private const string DevenvPathKey = "devenv_path";
    private const string InstanceIdKey = "instanceId";
    private const string PipeNameKey = "pipeName";
    private const string ProcessIdKey = "processId";
    private const string SolutionKey = "solution";
    private const string SolutionPathKey = "solutionPath";
    private const string PendingLaunchFlagDirectoryName = "vs-ide-bridge";
    private const string ServiceName = "VsIdeBridgeService";
    private const int VsCloseWaitTimeoutMilliseconds = 10_000;
    private const int VsOpenRegistrationTimeoutMilliseconds = 30_000;
    private const int LauncherResultTimeoutMilliseconds = 15_000;
    private const string LauncherExe = "VsIdeBridgeLauncher.exe";
    private const string LaunchVisualStudioCommand = "launch-visual-studio";
    private const string VsOpenLaunchOptInEnvironmentVariable = "VS_IDE_BRIDGE_ENABLE_VS_OPEN_LAUNCH";
    private static readonly SemaphoreSlim VsOpenGate = new(1, 1);
    private static readonly object PendingVsLaunchGate = new();
    private static int PendingVsLaunchPid;
    private static string? PendingVsLaunchSolution;
    private static DateTimeOffset PendingVsLaunchStartedAtUtc;
    private const uint CreateUnicodeEnvironment = 0x00000400;
    private const int StartfUseShowWindow = 0x00000001;
    private const short SwShowNormal = 1;
    private const int ErrorPrivilegeNotHeld = 1314;
    private const int LogonWithProfile = 0x00000001;

    private static async Task<JsonNode> BridgeHealthAsync(JsonNode? id, BridgeConnection bridge)
    {
        IReadOnlyList<BridgeInstance> visibleInstances = await VsDiscovery.ListAsync(bridge.Mode).ConfigureAwait(false);
        BridgeInstance? staleInstance = bridge.DropCachedInstanceIfStale(visibleInstances);
        BridgeInstance? instance = bridge.CurrentInstance;
        JsonArray visibleInstanceItems = [];
        foreach (BridgeInstance visibleInstance in visibleInstances)
        {
            visibleInstanceItems.Add(BridgeInstanceToHealthJson(visibleInstance));
        }

        JsonObject health = new()
        {
            ["success"] = true,
            ["serviceVersion"] = ServiceVersion,
            ["discoveryMode"] = bridge.Mode.ToString(),
            ["currentSolutionPath"] = bridge.CurrentSolutionPath,
            ["bound"] = instance is not null,
            ["visibleInstanceCount"] = visibleInstances.Count,
            ["visibleInstances"] = visibleInstanceItems,
            ["toolDiscovery"] = BuildToolDiscoveryGuidance(),
        };

        if (staleInstance is not null)
        {
            health["staleBindingDropped"] = BridgeInstanceToHealthJson(staleInstance);
            health["BindingNotice"] = $"Dropped stale binding to {staleInstance.Label}; call bind_solution or bind_instance to choose a visible instance.";
        }

        if (visibleInstances.Count > 1)
        {
            string instanceList = string.Join(", ", visibleInstances.Select(i => i.Label));
            string boundNotice = instance is not null
                ? $"Bound to: {instance.Label}."
                : "No instance currently bound.";
            string others = instance is not null
                ? string.Join(", ", visibleInstances
                    .Where(i => !string.Equals(i.InstanceId, instance.InstanceId, StringComparison.OrdinalIgnoreCase))
                    .Select(i => i.Label))
                : instanceList;
            health["multipleInstancesWarning"] =
                $"⚠ {visibleInstances.Count} VS instances are open ({instanceList}). " +
                boundNotice +
                $" Other open: {others}. " +
                "Use bind_instance (by pid or instance_id) or bind_solution to target a specific instance.";
        }

        if (instance is not null)
        {
            health["modelGuidance"] = BoundSessionHint;
            health["recommendedTools"] = BuildBoundRecommendedTools();
            health["instance"] = BridgeInstanceToHealthJson(instance);
        }

        return BridgeResult(health);
    }

    private static JsonObject BridgeInstanceToHealthJson(BridgeInstance instance) => new()
    {
        [InstanceIdKey] = instance.InstanceId,
        ["label"] = instance.Label,
        [PipeNameKey] = instance.PipeName,
        [ProcessIdKey] = instance.ProcessId,
        [SolutionPathKey] = instance.SolutionPath ?? string.Empty,
        ["solutionName"] = instance.SolutionName ?? string.Empty,
        ["source"] = instance.Source,
        ["version"] = instance.Version ?? "pre-3.0.6",
    };

    // Version of this MCP service process, so bridge_health always states what is running.
    private static readonly string ServiceVersion = ResolveServiceVersion();

    private static string ResolveServiceVersion()
    {
        string informational = typeof(ToolCatalog).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? string.Empty;
        int metadataStart = informational.IndexOf('+');
        if (metadataStart > 0)
        {
            informational = informational[..metadataStart];
        }

        return string.IsNullOrWhiteSpace(informational)
            ? typeof(ToolCatalog).Assembly.GetName().Version?.ToString() ?? "unknown"
            : informational;
    }

    private static async Task<JsonNode> VsOpenAsync(JsonNode? id, JsonObject? args, BridgeConnection bridge)
    {
        // Ensure the Windows service is running before trying to launch and discover VS.
        EnsureServiceRunning();

        string? solution = args?[SolutionKey]?.GetValue<string>();
        string devenvPath = ResolveRequestedDevenvPath(id, args);
        await VsOpenGate.WaitAsync().ConfigureAwait(false);
        try
        {
            JsonNode? reusedResult = await TryReuseLaunchAsync(solution, devenvPath).ConfigureAwait(false);
            if (reusedResult is not null)
            {
                return reusedResult;
            }

            JsonNode? bridgedLaunchResult = await TryLaunchViaBridgeAsync(id, bridge, devenvPath, solution).ConfigureAwait(false);
            if (bridgedLaunchResult is not null)
            {
                return bridgedLaunchResult;
            }

            if (!IsVsOpenLaunchEnabled())
            {
                throw new McpRequestException(
                    id,
                    McpErrorCodes.BridgeError,
                    $"'vs_open' launch is disabled because it is not yet production-ready and can destabilize Visual Studio startup. Start " +
                    $"Visual Studio manually and bind to it, or set {VsOpenLaunchOptInEnvironmentVariable}=1 only when deliberately testing " +
                    $"the launch path.");
            }

            return await LaunchNewVsInstanceAsync(id, devenvPath, solution).ConfigureAwait(false);
        }
        finally
        {
            VsOpenGate.Release();
        }
    }

    private static string ResolveRequestedDevenvPath(JsonNode? id, JsonObject? args)
    {
        string? explicitDevenv = args?[DevenvPathKey]?.GetValue<string>();
        return string.IsNullOrWhiteSpace(explicitDevenv)
            ? ResolveDevenvPath(id)
            : explicitDevenv;
    }

    private static JsonNode CreateVsOpenResult(
        int pid,
        string devenvPath,
        string? solution,
        bool reused,
        string? instanceId = null,
        string? label = null,
        string? launchMode = null)
    {
        JsonObject result = new()
        {
            ["Success"] = true,
            ["pid"] = pid,
            [DevenvPathKey] = devenvPath,
            [SolutionKey] = solution ?? string.Empty,
            ["reused"] = reused,
        };

        if (!string.IsNullOrWhiteSpace(instanceId))
        {
            result[InstanceIdKey] = instanceId;
        }

        if (!string.IsNullOrWhiteSpace(label))
        {
            result["label"] = label;
        }

        if (!string.IsNullOrWhiteSpace(launchMode))
        {
            result["launchMode"] = launchMode;
        }

        return BridgeResult(result);
    }

    private static async Task<JsonNode?> TryReuseLaunchAsync(string? solution, string devenvPath)
    {
        int existingPid = await TryReuseExistingVsProcessAsync(solution).ConfigureAwait(false);
        if (existingPid > 0)
        {
            return CreateVsOpenResult(existingPid, devenvPath, solution, reused: true);
        }

        if (!TryGetPendingVsLaunch(solution, out int pendingPid))
        {
            return null;
        }

        BridgeInstance? pendingInstance = await WaitForRegisteredInstanceAsync(
            pendingPid,
            solution,
            VsOpenRegistrationTimeoutMilliseconds).ConfigureAwait(false);

        if (pendingInstance is null)
        {
            CleanupFailedVsLaunch(pendingPid);
            return null;
        }

        ClearPendingVsLaunch(pendingInstance.ProcessId);
        return CreateVsOpenResult(
            pendingInstance.ProcessId,
            devenvPath,
            pendingInstance.SolutionPath ?? solution,
            reused: true,
            instanceId: pendingInstance.InstanceId,
            label: pendingInstance.Label);
    }

    private static async Task<JsonNode?> TryLaunchViaBridgeAsync(JsonNode? id, BridgeConnection bridge, string devenvPath, string? solution)
    {
        int launchedByBridgePid = await TryLaunchViaBoundInstanceAsync(id, bridge, devenvPath, solution).ConfigureAwait(false);
        if (launchedByBridgePid <= 0)
        {
            return null;
        }

        RecordPendingVsLaunch(launchedByBridgePid, solution);
        BridgeInstance? bridgedLaunchInstance = await WaitForRegisteredInstanceAsync(
            launchedByBridgePid,
            solution,
            VsOpenRegistrationTimeoutMilliseconds).ConfigureAwait(false);

        if (bridgedLaunchInstance is null)
        {
            CleanupFailedVsLaunch(launchedByBridgePid);
            throw new McpRequestException(
                id,
                McpErrorCodes.BridgeError,
                $"Interactive VS launch created PID {launchedByBridgePid} but it never registered a live VS IDE Bridge instance within {VsOpenRegistrationTimeoutMilliseconds} ms.");
        }

        ClearPendingVsLaunch(bridgedLaunchInstance.ProcessId);
        return CreateVsOpenResult(
            bridgedLaunchInstance.ProcessId,
            devenvPath,
            bridgedLaunchInstance.SolutionPath ?? solution,
            reused: false,
            instanceId: bridgedLaunchInstance.InstanceId,
            label: bridgedLaunchInstance.Label,
            launchMode: "interactive-bridge");
    }

    private static async Task<JsonNode> LaunchNewVsInstanceAsync(JsonNode? id, string devenvPath, string? solution)
    {
        bool deferSolutionOpen = ShouldLaunchInInteractiveSession() && !string.IsNullOrWhiteSpace(solution);
        int pid = LaunchVisualStudio(devenvPath, deferSolutionOpen ? null : solution);
        RecordPendingVsLaunch(pid, solution);

        if (deferSolutionOpen)
        {
            WritePendingSolutionOpenFlag(pid, solution!);
        }

        if (string.IsNullOrWhiteSpace(solution))
        {
            WriteNoSolutionFlag(pid);
        }

        BridgeInstance? launchedInstance = await WaitForRegisteredInstanceAsync(
            pid,
            solution,
            VsOpenRegistrationTimeoutMilliseconds).ConfigureAwait(false);

        if (launchedInstance is null)
        {
            CleanupFailedVsLaunch(pid);
            throw new McpRequestException(
                id,
                McpErrorCodes.BridgeError,
                $"Visual Studio launched as PID {pid} but never registered a live VS IDE Bridge instance within {VsOpenRegistrationTimeoutMilliseconds} ms.");
        }

        ClearPendingVsLaunch(launchedInstance.ProcessId);
        return CreateVsOpenResult(
            launchedInstance.ProcessId,
            devenvPath,
            launchedInstance.SolutionPath ?? solution,
            reused: false,
            instanceId: launchedInstance.InstanceId,
            label: launchedInstance.Label);
    }

    private static async Task<int> TryLaunchViaBoundInstanceAsync(JsonNode? id, BridgeConnection bridge, string devenvPath, string? solution)
    {
        if (bridge.CurrentInstance is null)
        {
            return 0;
        }

        JsonObject payload = new()
        {
            [DevenvPathKey] = devenvPath,
            [SolutionKey] = solution ?? string.Empty,
        };

        JsonObject response = await bridge.SendIgnoringSolutionHintAsync(id, LaunchVisualStudioCommand, payload.ToJsonString()).ConfigureAwait(false);
        JsonNode? launchedPidNode = response["Data"]?["pid"];
        if (launchedPidNode is null)
        {
            return 0;
        }

        return launchedPidNode.GetValue<int>();
    }

    private static async Task<JsonNode> WaitForInstanceAsync(
        JsonNode? id, JsonObject? args, BridgeConnection bridge)
    {
        string? solutionHint = args?["solution"]?.GetValue<string>();
        int timeoutMs = args?["timeout_ms"]?.GetValue<int?>() ?? 60_000;

        IReadOnlyList<BridgeInstance> existing =
            await VsDiscovery.ListAsync(bridge.Mode).ConfigureAwait(false);
        HashSet<string> existingIds = existing
            .Select(static instance => instance.InstanceId)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        using CancellationTokenSource cts = new(timeoutMs);
        while (!cts.Token.IsCancellationRequested)
        {
            IReadOnlyList<BridgeInstance> current;
            try
            {
                // Each discovery poll gets its own 5 s deadline so a hung mutex
                // or file scan cannot block us past the outer timeout.
                Task<IReadOnlyList<BridgeInstance>> listTask =
                    VsDiscovery.ListAsync(bridge.Mode);
                if (await Task.WhenAny(listTask, Task.Delay(5_000, cts.Token))
                        .ConfigureAwait(false) != listTask)
                {
                    break; // outer timeout or poll timeout — give up
                }
                current = listTask.Result;
            }
            catch (OperationCanceledException)
            {
                break;
            }

            BridgeInstance? found = current.FirstOrDefault(instance =>
                !existingIds.Contains(instance.InstanceId) &&
                (solutionHint is null ||
                 (instance.SolutionPath?.Contains(solutionHint, StringComparison.OrdinalIgnoreCase) ?? false) ||
                 (instance.SolutionName?.Contains(solutionHint, StringComparison.OrdinalIgnoreCase) ?? false)));

            if (found is not null)
            {
                return BridgeResult(new JsonObject
                {
                    ["Success"] = true,
                    ["instanceId"] = found.InstanceId,
                    ["label"] = found.Label,
                    [PipeNameKey] = found.PipeName,
                    ["processId"] = found.ProcessId,
                    [SolutionPathKey] = found.SolutionPath ?? string.Empty,
                });
            }

            try
            {
                await Task.Delay(500, cts.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                break;
            }
        }

        throw new McpRequestException(id, McpErrorCodes.BridgeError,
            $"No new VS instance appeared within {timeoutMs} ms.");
    }

    private static Task<JsonNode> VsCloseAsync(
        JsonNode? id, JsonObject? args, BridgeConnection bridge)
    {
        int pid = args?["process_id"]?.GetValue<int?>() ??
                  bridge.CurrentInstance?.ProcessId ?? 0;
        bool force = args?["force"]?.GetValue<bool?>() ?? false;
        if (pid <= 0)
        {
            throw new McpRequestException(id, McpErrorCodes.InvalidParams,
                "No process_id specified and no VS instance bound.");
        }

        try
        {
            using Process process = Process.GetProcessById(pid);
            if (force)
            {
                process.Kill(entireProcessTree: true);
            }
            else
            {
                process.CloseMainWindow();
                if (!process.WaitForExit(VsCloseWaitTimeoutMilliseconds) && !process.HasExited)
                {
                    process.Kill(entireProcessTree: true);
                }
            }

            ClearPendingVsLaunch(pid);

            return Task.FromResult((JsonNode)BridgeResult(new JsonObject
            {
                ["Success"] = true,
                ["processId"] = pid,
                ["forced"] = force,
            }));
        }
        catch (Exception ex) when (ex is not null) // re-throw as MCP error; process kill can throw Win32Exception or InvalidOperationException
        {
            throw new McpRequestException(id, McpErrorCodes.BridgeError,
                $"Failed to close VS process {pid}: {ex.Message}");
        }
    }

    private static async Task<int> TryReuseExistingVsProcessAsync(string? solution)
    {
        IReadOnlyList<BridgeInstance> instances = await VsDiscovery.ListAsync().ConfigureAwait(false);
        BridgeInstance? existing = instances.FirstOrDefault(instance =>
            string.IsNullOrWhiteSpace(solution) ||
            string.Equals(instance.SolutionPath, solution, StringComparison.OrdinalIgnoreCase));
        if (existing is not null && existing.ProcessId > 0)
        {
            ClearPendingVsLaunch(existing.ProcessId);
            return existing.ProcessId;
        }

        return 0;
    }

    private static bool TryGetPendingVsLaunch(string? solution, out int pendingPid)
    {
        lock (PendingVsLaunchGate)
        {
            if (PendingVsLaunchPid <= 0)
            {
                pendingPid = 0;
                return false;
            }

            bool solutionMatches = string.IsNullOrWhiteSpace(solution) ||
                string.Equals(PendingVsLaunchSolution, solution, StringComparison.OrdinalIgnoreCase);
            bool launchIsFresh = DateTimeOffset.UtcNow - PendingVsLaunchStartedAtUtc < TimeSpan.FromMinutes(2);
            if (!solutionMatches || !launchIsFresh)
            {
                pendingPid = 0;
                return false;
            }

            try
            {
                using Process process = Process.GetProcessById(PendingVsLaunchPid);
                if (!process.HasExited)
                {
                    pendingPid = PendingVsLaunchPid;
                    return true;
                }
            }
        catch (Exception ex) when (ex is not null) // best-effort probe; race conditions in shared state
            {
                LogIgnoredException("Failed to probe pending Visual Studio launch state.", ex);
            }

            PendingVsLaunchPid = 0;
            PendingVsLaunchSolution = null;
            PendingVsLaunchStartedAtUtc = default;
            pendingPid = 0;
            return false;
        }
    }

    private static void RecordPendingVsLaunch(int pid, string? solution)
    {
        lock (PendingVsLaunchGate)
        {
            PendingVsLaunchPid = pid;
            PendingVsLaunchSolution = solution;
            PendingVsLaunchStartedAtUtc = DateTimeOffset.UtcNow;
        }
    }

    private static void ClearPendingVsLaunch(int pid)
    {
        lock (PendingVsLaunchGate)
        {
            if (PendingVsLaunchPid != pid)
            {
                return;
            }

            PendingVsLaunchPid = 0;
            PendingVsLaunchSolution = null;
            PendingVsLaunchStartedAtUtc = default;
        }
    }

    private static async Task<BridgeInstance?> WaitForRegisteredInstanceAsync(int pid, string? solution, int timeoutMs)
    {
        using CancellationTokenSource cts = new(timeoutMs);
        while (!cts.Token.IsCancellationRequested)
        {
            try
            {
                IReadOnlyList<BridgeInstance> instances = await VsDiscovery.ListAsync().ConfigureAwait(false);
                BridgeInstance? instance = instances.FirstOrDefault(instance =>
                    instance.ProcessId == pid &&
                    (string.IsNullOrWhiteSpace(solution) ||
                     string.Equals(instance.SolutionPath, solution, StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(instance.SolutionName, System.IO.Path.GetFileName(solution), StringComparison.OrdinalIgnoreCase)));
                if (instance is not null)
                {
                    return instance;
                }
            }
        catch (Exception ex) when (ex is not null) // best-effort discovery query while waiting for VS
            {
                LogIgnoredException($"Failed to query discovery state while waiting for Visual Studio pid {pid}.", ex);
            }

            try
            {
                using Process process = Process.GetProcessById(pid);
                if (process.HasExited)
                {
                    break;
                }
            }
            catch
            {
                break;
            }

            try
            {
                await Task.Delay(500, cts.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                break;
            }
        }

        return null;
    }

    private static void CleanupFailedVsLaunch(int pid)
    {
        ClearPendingVsLaunch(pid);
        DeletePendingLaunchFlags(pid);
        try
        {
            using Process process = Process.GetProcessById(pid);
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
            }
        }
        catch (Exception ex) when (ex is not null) // best-effort process kill
        {
            LogIgnoredException($"Failed to terminate launched Visual Studio pid {pid}.", ex);
        }
    }

    private static int LaunchVisualStudio(string devenvPath, string? solution)
    {
        if (ShouldLaunchInInteractiveSession())
        {
            return LaunchVisualStudioInInteractiveSession(devenvPath, solution);
        }

        ProcessStartInfo psi = new()
        {
            FileName = devenvPath,
            UseShellExecute = true,
            WorkingDirectory = System.IO.Path.GetDirectoryName(devenvPath),
        };

        if (!string.IsNullOrWhiteSpace(solution))
        {
            psi.ArgumentList.Add(solution);
        }

        using Process? process = Process.Start(psi);
        _ = process ?? throw new InvalidOperationException("Visual Studio launch failed: Process.Start returned null.");

        return process.Id;
    }

    private static async Task<JsonNode> ListInstancesAsync(BridgeConnection bridge)
    {
        // Ensure the Windows service is running so discovery actually has something to find.
        EnsureServiceRunning();

        IReadOnlyList<BridgeInstance> instances = await VsDiscovery
            .ListAsync(bridge.Mode).ConfigureAwait(false);

        JsonArray resultItems = [];
        foreach (BridgeInstance instance in instances)
        {
            resultItems.Add(new JsonObject
            {
                ["instanceId"] = instance.InstanceId,
                ["label"] = instance.Label,
                [PipeNameKey] = instance.PipeName,
                ["processId"] = instance.ProcessId,
                [SolutionPathKey] = instance.SolutionPath ?? string.Empty,
                ["solutionName"] = instance.SolutionName ?? string.Empty,
                ["source"] = instance.Source,
            });
        }

        JsonObject result = new()
        {
            ["success"] = true,
            ["instances"] = resultItems,
        };
        return BridgeResult(result);
    }

    /// <summary>
    /// Ensures the VsIdeBridgeService Windows service is running.
    /// Called before vs_open and list_instances so that a stopped-but-installed
    /// service auto-resumes rather than silently returning empty results.
    /// Failures are swallowed — the caller proceeds regardless.
    /// </summary>
    private static void EnsureServiceRunning()
    {
        if (!OperatingSystem.IsWindows()) return;
        try
        {
            using System.ServiceProcess.ServiceController sc = new(ServiceName);
            if (sc.Status != System.ServiceProcess.ServiceControllerStatus.Running &&
                sc.Status != System.ServiceProcess.ServiceControllerStatus.StartPending)
            {
                sc.Start();
                sc.WaitForStatus(
                    System.ServiceProcess.ServiceControllerStatus.Running,
                    TimeSpan.FromSeconds(10));
            }
        }
        catch (InvalidOperationException ex)
        {
            McpServerLog.WriteException("failed to start VsIdeBridgeService because the service is unavailable", ex);
        }
        catch (Win32Exception ex)
        {
            McpServerLog.WriteException("failed to start VsIdeBridgeService because service control was denied", ex);
        }
        catch (System.ServiceProcess.TimeoutException ex)
        {
            McpServerLog.WriteException("VsIdeBridgeService did not report Running within the best-effort wait window", ex);
        }
    }
}
