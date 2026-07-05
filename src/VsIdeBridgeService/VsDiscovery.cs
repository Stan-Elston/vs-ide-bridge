using System.IO.MemoryMappedFiles;
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace VsIdeBridgeService;

// Discovers live VS IDE Bridge instances from memory-mapped stores and JSON pipe files.
internal static class VsDiscovery
{
    private const string LocalMapName = @"Local\VsIdeBridge.Discovery.v1";
    private const string LocalMutexName = @"Local\VsIdeBridge.Discovery.v1.mutex";
    private const string GlobalMapName = @"Global\VsIdeBridge.Discovery.v1";
    private const string GlobalMutexName = @"Global\VsIdeBridge.Discovery.v1.mutex";
    private const int MemoryCapacityBytes = 1024 * 1024;

    private static readonly TimeSpan MutexWait = TimeSpan.FromSeconds(2);
    private static readonly TimeSpan HeadlessProcessGracePeriod = TimeSpan.FromSeconds(30);
    private static readonly (string Map, string Mutex, string Source)[] MemoryStores =
    [
        (LocalMapName,  LocalMutexName,  "memory://local/VsIdeBridge.Discovery.v1"),
        (GlobalMapName, GlobalMutexName, "memory://global/VsIdeBridge.Discovery.v1"),
    ];

    public static async Task<IReadOnlyList<BridgeInstance>> ListAsync(DiscoveryMode mode = DiscoveryMode.MemoryFirst)
    {
        IReadOnlyList<BridgeInstance> jsonInstances = mode == DiscoveryMode.MemoryFirst
            || mode == DiscoveryMode.JsonOnly
            || mode == DiscoveryMode.Hybrid
            ? await ListJsonAsync().ConfigureAwait(false)
            : [];

        if (mode == DiscoveryMode.JsonOnly) return jsonInstances;

        List<BridgeInstance> memoryInstances = ListMemory();
        IReadOnlyList<BridgeInstance> merged = mode == DiscoveryMode.MemoryFirst
            ? Merge(memoryInstances, jsonInstances, preferPrimary: true)
            : Merge(memoryInstances, jsonInstances, preferPrimary: false);
        IReadOnlyList<BridgeInstance> liveInstances = FilterLiveInstances(merged);
        IReadOnlyList<BridgeInstance> enriched = await EnrichIncompleteInstancesAsync(liveInstances).ConfigureAwait(false);
        IReadOnlyList<BridgeInstance> withSolutionPaths =
            [.. enriched.Where(static instance => !string.IsNullOrWhiteSpace(instance.SolutionPath))];
        return withSolutionPaths.Count > 0 ? withSolutionPaths : enriched;
    }

    public static async Task<BridgeInstance> SelectAsync(BridgeInstanceSelector selector, DiscoveryMode mode = DiscoveryMode.MemoryFirst)
    {
        IReadOnlyList<BridgeInstance> instances = await ListAsync(mode).ConfigureAwait(false);

        if (instances.Count == 0)
            throw new BridgeException(
                "No live VS IDE Bridge instance found. " +
                "Visual Studio must be running with the VS IDE Bridge extension installed. " +
                "Ask the user to open Visual Studio with the desired solution and then call list_instances to get the instance_id, then call bind_instance to connect.");

        BridgeInstance[] matches = [.. Filter(instances, selector)];
        if (matches.Length == 1) return matches[0];

        if (matches.Length == 0)
            throw new BridgeException(selector.HasAny
                ? $"No live VS IDE Bridge instance matched {selector.Describe()}. " +
                  $"Call list_instances to see what is currently open, " +
                  $"then call bind_instance with the correct instance_id from that list to rebind. " +
                  $"If the target solution is not in that list it is not open in Visual Studio yet — " +
                  $"use glob with pattern **/*.sln and set the path to one level above the current solution root (the repos root) to search sibling repositories on disk. " +
                  $"Ask the user if they want to save any unsaved work in the current solution first. " +
                  $"Then tell the user which .sln you found and offer two options: " +
                  $"(a) open it in a new VS window — both solutions stay open and you can switch between them at any time using bind_instance, or " +
                  $"(b) swap it into the current VS window using open_solution — this closes the current solution in that window. " +
                  $"Once the target solution is open in VS, call list_instances to confirm it appeared, then call bind_instance with the new instance_id."
                : "Multiple VS IDE Bridge instances are available. Call list_instances to see them, then call bind_solution or bind_instance to select one.");

        throw new BridgeException(
            $"Multiple VS IDE Bridge instances matched {selector.Describe()}. Call list_instances to see them, then call bind_instance with a specific instance_id.");
    }

    public static IEnumerable<BridgeInstance> Filter(IEnumerable<BridgeInstance> instances, BridgeInstanceSelector sel)
    {
        foreach (BridgeInstance inst in instances)
        {
            if (!string.IsNullOrWhiteSpace(sel.InstanceId) &&
                !string.Equals(inst.InstanceId, sel.InstanceId, StringComparison.OrdinalIgnoreCase))
                continue;

            if (sel.ProcessId.HasValue && inst.ProcessId != sel.ProcessId.Value)
                continue;

            if (!string.IsNullOrWhiteSpace(sel.PipeName) &&
                !string.Equals(inst.PipeName, sel.PipeName, StringComparison.OrdinalIgnoreCase))
                continue;

            if (!string.IsNullOrWhiteSpace(sel.SolutionHint))
            {
                bool matched = inst.SolutionPath.Contains(sel.SolutionHint, StringComparison.OrdinalIgnoreCase)
                    || inst.SolutionName.Contains(sel.SolutionHint, StringComparison.OrdinalIgnoreCase);
                if (!matched) continue;
            }

            yield return inst;
        }
    }

    // ── Memory-mapped store discovery ──────────────────────────────────────────

    private static List<BridgeInstance> ListMemory()
    {
        if (!OperatingSystem.IsWindows()) return [];

        List<BridgeInstance> collected = [];
        foreach ((string mapName, string mutexName, string source) in MemoryStores)
        {
            List<BridgeInstance> fromStore = TryListMemoryStore(mapName, mutexName, source);
            collected.AddRange(fromStore);
        }

        if (collected.Count <= 1) return collected;

        return [.. collected
            .GroupBy(i => $"{i.InstanceId}:{i.ProcessId}", StringComparer.OrdinalIgnoreCase)
            .Select(g => g.OrderByDescending(i => i.LastWriteTimeUtc).First())
            .OrderByDescending(i => i.LastWriteTimeUtc)];
    }

    private static List<BridgeInstance> TryListMemoryStore(string mapName, string mutexName, string source)
    {
        if (!OperatingSystem.IsWindows()) return [];

        try
        {
            using System.Threading.Mutex mutex = new(false, mutexName);
            bool hasLock = false;
            try
            {
                hasLock = mutex.WaitOne(MutexWait);
                if (!hasLock) return [];

                using MemoryMappedFile mmf = MemoryMappedFile.OpenExisting(mapName);
                using MemoryMappedViewStream stream = mmf.CreateViewStream(0, MemoryCapacityBytes, MemoryMappedFileAccess.Read);
                byte[] raw = new byte[MemoryCapacityBytes];
                int read = stream.Read(raw, 0, raw.Length);
                int nullPos = Array.IndexOf(raw, (byte)0);
                int len = nullPos >= 0 ? nullPos : read;
                string json = Encoding.UTF8.GetString(raw, 0, len).Trim();
                if (string.IsNullOrWhiteSpace(json)) return [];

                return ParseInstancesFromJson(json, source);
            }
            finally
            {
                if (hasLock) mutex.ReleaseMutex();
            }
        }
        catch
        {
            return [];
        }
    }

    // ── JSON pipe-file discovery ───────────────────────────────────────────────

    private static async Task<IReadOnlyList<BridgeInstance>> ListJsonAsync()
    {
        foreach (string directory in EnumerateCandidatePipeDirectories())
        {
            try
            {
                if (!Directory.Exists(directory)) continue;

                FileInfo[] files = [.. Directory.GetFiles(directory, "bridge-*.json")
                    .Select(p => new FileInfo(p))
                    .OrderByDescending(f => f.LastWriteTimeUtc)];

                if (files.Length == 0) continue;

                List<BridgeInstance> instances = [];
                foreach (FileInfo file in files)
                {
                    BridgeInstance? inst = await TryLoadJsonFileAsync(file.FullName).ConfigureAwait(false);
                    if (inst is not null) instances.Add(inst);
                }

                if (instances.Count > 0) return instances;
            }
            catch (IOException ex) { McpServerLog.WriteException($"failed to scan discovery pipe directory '{directory}'", ex); }
            catch (UnauthorizedAccessException ex) { McpServerLog.WriteException($"failed to scan discovery pipe directory '{directory}'", ex); }
        }

        return [];
    }

    private static HashSet<string> EnumerateCandidatePipeDirectories()
    {
        HashSet<string> directories = new(StringComparer.OrdinalIgnoreCase);

        foreach (string tempDir in new[]
                 {
                     Environment.GetEnvironmentVariable("TEMP"),
                     Environment.GetEnvironmentVariable("TMP"),
                     Path.GetTempPath(),
                 }
                 .Where(static p => !string.IsNullOrWhiteSpace(p))
                 .Select(static p => p!))
        {
            directories.Add(Path.Combine(tempDir, "vs-ide-bridge", "pipes"));
        }

        string systemDrive = Environment.GetEnvironmentVariable("SystemDrive") ?? "C:";
        string usersRoot = Path.Combine(systemDrive, "Users");
        try
        {
            if (Directory.Exists(usersRoot))
            {
                foreach (string profileDir in Directory.GetDirectories(usersRoot))
                {
                    directories.Add(Path.Combine(profileDir, "AppData", "Local", "Temp", "vs-ide-bridge", "pipes"));
                }
            }
        }
        catch (IOException ex)
        {
            McpServerLog.WriteException($"failed to enumerate user profile temp discovery directories under '{usersRoot}'", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            McpServerLog.WriteException($"failed to enumerate user profile temp discovery directories under '{usersRoot}'", ex);
        }
        catch (NotSupportedException ex)
        {
            McpServerLog.WriteException($"failed to enumerate user profile temp discovery directories under '{usersRoot}'", ex);
        }

        return directories;
    }

    private static async Task<BridgeInstance?> TryLoadJsonFileAsync(string path)
    {
        try
        {
            string json = await File.ReadAllTextAsync(path).ConfigureAwait(false);
            if (JsonNode.Parse(json) is not JsonObject obj)
            {
                return null;
            }
            BridgeInstance? instance = ParseInstanceFromObject(obj, path, File.GetLastWriteTimeUtc(path));
            if (instance is null)
            {
                return null;
            }

            if (!IsLiveInstance(instance))
            {
                TryDeleteStaleDiscoveryFile(path);
                return null;
            }

            return instance;
        }
        catch
        {
            return null;
        }
    }

    // ── Parsing helpers ────────────────────────────────────────────────────────

    private static List<BridgeInstance> ParseInstancesFromJson(string json, string source)
    {
        try
        {
            JsonNode? node = JsonNode.Parse(json);
            List<BridgeInstance> result = [];
            if (node is JsonArray arr)
            {
                foreach (JsonNode? item in arr)
                {
                    if (item is JsonObject obj)
                    {
                        BridgeInstance? inst = ParseInstanceFromObject(obj, source);
                        if (inst is not null) result.Add(inst);
                    }
                }
            }
            else if (node is JsonObject single)
            {
                BridgeInstance? inst = ParseInstanceFromObject(single, source);
                if (inst is not null) result.Add(inst);
            }

            return result;
        }
        catch
        {
            return [];
        }
    }

    private static BridgeInstance? ParseInstanceFromObject(JsonObject obj, string source, DateTime? fallbackLastWriteTimeUtc = null)
    {
        string? instanceId = obj["instanceId"]?.GetValue<string>();
        string? pipeName = obj["pipeName"]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(instanceId) || string.IsNullOrWhiteSpace(pipeName))
            return null;

        return new BridgeInstance
        {
            InstanceId = instanceId,
            PipeName = pipeName,
            ProcessId = obj["pid"]?.GetValue<int>() ?? 0,
            SolutionPath = obj["solutionPath"]?.GetValue<string>() ?? string.Empty,
            SolutionName = Path.GetFileName(obj["solutionPath"]?.GetValue<string>() ?? string.Empty),
            Label = GetInstanceLabel(obj),
            Source = source,
            StartedAtUtc = obj["startedAtUtc"]?.GetValue<string>(),
            Version = obj["version"]?.GetValue<string>(),
            DiscoveryFile = source,
            LastWriteTimeUtc = DateTime.TryParse(
                obj["lastWriteTimeUtc"]?.GetValue<string>(), out DateTime dt)
                ? dt
                : fallbackLastWriteTimeUtc ?? DateTime.MinValue,
        };
    }

    private static string GetInstanceLabel(JsonObject obj)
    {
        string? explicitLabel = obj["label"]?.GetValue<string>();
        if (!string.IsNullOrWhiteSpace(explicitLabel))
        {
            return explicitLabel;
        }

        string solutionPath = obj["solutionPath"]?.GetValue<string>() ?? string.Empty;
        string pipeName = obj["pipeName"]?.GetValue<string>() ?? string.Empty;
        int processId = obj["pid"]?.GetValue<int>() ?? 0;
        string solutionBaseName = string.IsNullOrWhiteSpace(solutionPath)
            ? string.Empty
            : Path.GetFileNameWithoutExtension(solutionPath);
        string name = string.IsNullOrWhiteSpace(solutionBaseName)
            ? (string.IsNullOrWhiteSpace(pipeName) ? "Visual Studio" : pipeName)
            : solutionBaseName;

        return processId > 0 ? $"{name} ({processId})" : name;
    }

    // ── Merge logic ────────────────────────────────────────────────────────────

    private static List<BridgeInstance> Merge(
        List<BridgeInstance> primary,
        IReadOnlyList<BridgeInstance> secondary,
        bool preferPrimary)
    {
        if (primary.Count == 0) return [.. secondary];
        if (secondary.Count == 0) return primary;

        System.Collections.Generic.HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
        List<BridgeInstance> result = [];

        IEnumerable<BridgeInstance> first = preferPrimary ? primary : (IEnumerable<BridgeInstance>)secondary;
        IEnumerable<BridgeInstance> second = preferPrimary ? secondary : (IEnumerable<BridgeInstance>)primary;

        foreach (BridgeInstance inst in first)
        {
            if (seen.Add(inst.PipeName))
            {
                result.Add(inst);
            }
        }

        foreach (BridgeInstance inst in second)
        {
            int existingIndex = result.FindIndex(existing => string.Equals(existing.PipeName, inst.PipeName, StringComparison.OrdinalIgnoreCase));
            if (existingIndex >= 0)
            {
                result[existingIndex] = ChoosePreferredInstance(result[existingIndex], inst, preferPrimary);
                continue;
            }

            if (seen.Add(inst.PipeName))
            {
                result.Add(inst);
            }
        }

        return result;
    }

    private static BridgeInstance ChoosePreferredInstance(BridgeInstance first, BridgeInstance second, bool preferFirst)
    {
        bool firstHasSolution = !string.IsNullOrWhiteSpace(first.SolutionPath);
        bool secondHasSolution = !string.IsNullOrWhiteSpace(second.SolutionPath);
        if (firstHasSolution != secondHasSolution)
        {
            return firstHasSolution ? first : second;
        }

        if (first.LastWriteTimeUtc != second.LastWriteTimeUtc)
        {
            return first.LastWriteTimeUtc >= second.LastWriteTimeUtc ? first : second;
        }

        return preferFirst ? first : second;
    }

    private static IReadOnlyList<BridgeInstance> FilterLiveInstances(IReadOnlyList<BridgeInstance> instances)
    {
        if (instances.Count == 0)
        {
            return instances;
        }

        List<BridgeInstance> liveInstances = [];
        foreach (BridgeInstance instance in instances)
        {
            if (IsLiveInstance(instance))
            {
                liveInstances.Add(instance);
                continue;
            }

            if (instance.DiscoveryFile.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
            {
                TryDeleteStaleDiscoveryFile(instance.DiscoveryFile);
            }
        }

        return liveInstances;
    }

    private static async Task<IReadOnlyList<BridgeInstance>> EnrichIncompleteInstancesAsync(IReadOnlyList<BridgeInstance> instances)
    {
        if (instances.Count == 0)
        {
            return instances;
        }

        List<BridgeInstance> enriched = [];
        foreach (BridgeInstance instance in instances)
        {
            if (!string.IsNullOrWhiteSpace(instance.SolutionPath))
            {
                enriched.Add(instance);
                continue;
            }

            BridgeInstance refreshed = await TryRefreshInstanceStateAsync(instance).ConfigureAwait(false);
            enriched.Add(refreshed);
        }

        return enriched;
    }

    private static async Task<BridgeInstance> TryRefreshInstanceStateAsync(BridgeInstance instance)
    {
        try
        {
            if (instance.DiscoveryFile.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
            {
                BridgeInstance? refreshedFromFile = await TryLoadJsonFileAsync(instance.DiscoveryFile).ConfigureAwait(false);
                if (refreshedFromFile is not null && !string.IsNullOrWhiteSpace(refreshedFromFile.SolutionPath))
                {
                    return refreshedFromFile;
                }
            }

            await using VsPipeClient client = await VsPipeClient.CreateAsync(instance.PipeName, 3_000, 500).ConfigureAwait(false);
            JsonObject response = await client.SendAsync(new JsonObject
            {
                ["id"] = "discover",
                ["command"] = "state",
                ["args"] = new JsonObject(),
            }).ConfigureAwait(false);

            JsonObject payload = response["Data"] as JsonObject
                ?? response["data"] as JsonObject
                ?? response;
            string? solutionPath = payload["solutionPath"]?.GetValue<string>();
            if (string.IsNullOrWhiteSpace(solutionPath))
            {
                return instance;
            }

            string solutionName = Path.GetFileName(solutionPath);
            string label = string.IsNullOrWhiteSpace(instance.Label) || string.Equals(instance.Label, instance.PipeName, StringComparison.OrdinalIgnoreCase)
                ? BuildFallbackLabel(solutionPath, instance.ProcessId, instance.PipeName)
                : instance.Label;

            return instance with
            {
                SolutionPath = solutionPath,
                SolutionName = solutionName,
                Label = label,
            };
        }
        catch
        {
            return instance;
        }
    }

    private static string BuildFallbackLabel(string solutionPath, int processId, string pipeName)
    {
        string solutionBaseName = Path.GetFileNameWithoutExtension(solutionPath);
        string name = string.IsNullOrWhiteSpace(solutionBaseName) ? pipeName : solutionBaseName;
        return processId > 0 ? $"{name} ({processId})" : name;
    }

    private static bool IsLiveInstance(BridgeInstance instance)
    {
        if (instance.ProcessId <= 0)
        {
            return false;
        }

        try
        {
            using Process process = Process.GetProcessById(instance.ProcessId);
            if (process.HasExited)
            {
                return false;
            }

            // Window handles are session-scoped: when the Windows service (session 0)
            // probes an interactive VS process, MainWindowHandle is always zero, so the
            // zombie-window heuristic below would misread every real instance as dead
            // and delete its discovery file. Cross-session, a running process is the
            // strongest signal available - treat it as live.
            if (!IsInCurrentSession(process))
            {
                return true;
            }

            if (HasUsableMainWindow(process))
            {
                return true;
            }

            return IsWithinHeadlessGracePeriod(instance);
        }
        catch (ArgumentException ex)
        {
            McpServerLog.WriteException($"failed to probe process liveness for '{instance.ProcessId}'", ex);
            return false;
        }
        catch (InvalidOperationException ex)
        {
            McpServerLog.WriteException($"failed to probe process liveness for '{instance.ProcessId}'", ex);
            return false;
        }
        catch (NotSupportedException ex)
        {
            McpServerLog.WriteException($"failed to probe process liveness for '{instance.ProcessId}'", ex);
            return false;
        }
        catch (System.ComponentModel.Win32Exception ex)
        {
            McpServerLog.WriteException($"failed to probe process liveness for '{instance.ProcessId}'", ex);
            return false;
        }
    }

    internal static bool IsProcessGone(int processId)
    {
        if (processId <= 0)
            return false;
        try
        {
            using Process process = Process.GetProcessById(processId);
            return process.HasExited;
        }
        catch (ArgumentException)
        {
            return true; // Process not found — definitely gone.
        }
        catch
        {
            return false; // Unknown state — don't assume the session is lost.
        }
    }

    private static bool HasUsableMainWindow(Process process)
    {
        try
        {
            return process.MainWindowHandle != IntPtr.Zero;
        }
        catch
        {
            return false;
        }
    }

    private static readonly int CurrentSessionId = GetCurrentSessionId();

    private static bool IsInCurrentSession(Process process)
    {
        try
        {
            return process.SessionId == CurrentSessionId;
        }
        catch
        {
            // Session unknown: assume cross-session so the window heuristic is skipped
            // rather than misclassifying a live instance as dead.
            return false;
        }
    }

    private static int GetCurrentSessionId()
    {
        try
        {
            using Process current = Process.GetCurrentProcess();
            return current.SessionId;
        }
        catch
        {
            return -1;
        }
    }

    private static bool IsWithinHeadlessGracePeriod(BridgeInstance instance)
    {
        if (instance.LastWriteTimeUtc == DateTime.MinValue)
        {
            return false;
        }

        return DateTime.UtcNow - instance.LastWriteTimeUtc <= HeadlessProcessGracePeriod;
    }

    private static void TryDeleteStaleDiscoveryFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException ex)
        {
            McpServerLog.WriteException($"failed to delete stale discovery file '{path}'", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            McpServerLog.WriteException($"failed to delete stale discovery file '{path}'", ex);
        }
        catch (NotSupportedException ex)
        {
            McpServerLog.WriteException($"failed to delete stale discovery file '{path}'", ex);
        }
    }
}
