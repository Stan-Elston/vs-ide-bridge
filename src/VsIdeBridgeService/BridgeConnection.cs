using System.Text.Json.Nodes;
using VsIdeBridgeService.Diagnostics;
using static VsIdeBridgeService.BridgeConnectionDefaults;

namespace VsIdeBridgeService;

// Manages a VS bridge connection for one MCP session.
// Caches the discovered instance to avoid repeated discovery on every tool call.
// Thread-safe: multiple concurrent tool calls may use the same connection.
internal sealed class BridgeConnection
{
    private readonly DiscoveryMode _discoveryMode;
    private readonly int? _timeoutOverrideMs;
    private readonly object _gate = new();
    private readonly DocumentDiagnosticsCoordinator _documentDiagnostics;
    private readonly ConnectionState _state = new();

    public BridgeConnection(string[] args)
    {
        _discoveryMode = BridgeConnectionArgs.ResolveMode(args);
        _timeoutOverrideMs = BridgeConnectionArgs.GetOptionalPositiveInt(args, "timeout-ms");
        _documentDiagnostics = new DocumentDiagnosticsCoordinator(this);
    }

    internal enum ToolTimeoutProfile
    {
        Fast,
        Interactive,
        Heavy,
        BuildWait,
    }

    private sealed class ConnectionState
    {
        public BridgeInstanceSelector Selector { get; set; } = new();
        public BridgeInstance? Cached { get; set; }
        public string LastSolutionPath { get; set; } = string.Empty;
        public string? PendingBindingNotice { get; set; }
    }

    // ── Public API used by tool handlers ──────────────────────────────────────

    public Task<JsonObject> SendAsync(JsonNode? id, string command, string args)
        => SendCoreAsync(id, command, JsonValue.Create(args), ignoreSolutionHint: false);

    public Task<JsonObject> SendAsync(JsonNode? id, string command, string args, ToolTimeoutProfile timeoutProfile)
        => SendCoreAsync(id, BuildRequest(command, JsonValue.Create(args)), ignoreSolutionHint: false, timeoutProfile);

    public Task<JsonObject> SendAsync(JsonNode? id, string command, JsonObject? args)
        => SendCoreAsync(id, command, args?.DeepClone(), ignoreSolutionHint: false);

    public Task<JsonObject> SendAsync(JsonNode? id, string command, JsonObject? args, ToolTimeoutProfile timeoutProfile)
        => SendCoreAsync(id, BuildRequest(command, args?.DeepClone()), ignoreSolutionHint: false, timeoutProfile);

    public Task<JsonObject> SendBatchAsync(JsonNode? id, JsonArray steps, bool stopOnError = false)
        => SendCoreAsync(id, BuildBatchRequest(steps, stopOnError), ignoreSolutionHint: false,
            BridgeConnectionArgs.SelectTimeoutProfile("batch"));

    public Task<JsonObject> SendIgnoringSolutionHintAsync(JsonNode? id, string command, string args)
        => SendCoreAsync(id, command, JsonValue.Create(args), ignoreSolutionHint: true);

    public Task<JsonObject> SendIgnoringSolutionHintAsync(JsonNode? id, string command, JsonObject? args)
        => SendCoreAsync(id, command, args?.DeepClone(), ignoreSolutionHint: true);

    public string CurrentSolutionPath
    {
        get { lock (_gate) { return _state.LastSolutionPath; } }
    }

    public BridgeInstance? CurrentInstance
    {
        get { lock (_gate) { return _state.Cached; } }
    }

    public BridgeInstanceSelector CurrentSelector
    {
        get { lock (_gate) { return _state.Selector; } }
    }

    public DiscoveryMode Mode => _discoveryMode;

    public DocumentDiagnosticsCoordinator DocumentDiagnostics => _documentDiagnostics;

    // Bind to a specific instance and return binding info.
    public async Task<JsonObject> BindAsync(JsonNode? id, JsonObject? args)
    {
        BridgeInstanceSelector newSelector = BridgeConnectionArgs.ParseSelector(args);
        lock (_gate)
        {
            _state.Selector = newSelector;
            _state.Cached = null;
            _state.PendingBindingNotice = null;
        }

        try
        {
            BridgeInstance discovered = await GetInstanceAsync(ignoreSolutionHint: false).ConfigureAwait(false);
            RememberSolutionPath(discovered.SolutionPath);
            _documentDiagnostics.QueueRefreshAndGetSnapshot("bind", clearCached: true);
            JsonObject bindResponse = new()
            {
                ["success"] = true,
                ["binding"] = InstanceToJson(discovered),
                ["selector"] = SelectorToJson(CurrentSelector),
            };
            ToolCatalog.AttachBoundSessionGuidance(bindResponse);
            return bindResponse;
        }
        catch (BridgeException ex)
        {
            throw new McpRequestException(id, BridgeError, ex.Message);
        }
    }

    // Prefer a solution for future discover without full rebind.
    public void PreferSolution(string? solutionHint)
    {
        lock (_gate)
        {
            _state.Selector = new BridgeInstanceSelector
            {
                InstanceId = _state.Selector.InstanceId,
                ProcessId = _state.Selector.ProcessId,
                PipeName = _state.Selector.PipeName,
                SolutionHint = solutionHint,
            };
            _state.Cached = null;
            _state.PendingBindingNotice = null;
        }
    }

    // ── Internal send logic ────────────────────────────────────────────────────

    private async Task<JsonObject> SendCoreAsync(JsonNode? id, string command, JsonNode? args, bool ignoreSolutionHint)
        => await SendCoreAsync(id, BuildRequest(command, args), ignoreSolutionHint,
            BridgeConnectionArgs.SelectTimeoutProfile(command)).ConfigureAwait(false);

    private async Task<JsonObject> SendCoreAsync(
        JsonNode? id,
        JsonObject request,
        bool ignoreSolutionHint,
        ToolTimeoutProfile timeoutProfile)
    {
        try
        {
            JsonObject response = await SendPipeAsync(request, ignoreSolutionHint, timeoutProfile).ConfigureAwait(false);
            return await FinalizePipeResponseAsync(response, request, ignoreSolutionHint, timeoutProfile).ConfigureAwait(false);
        }
        catch (McpRequestException)
        {
            throw;
        }
        catch (BridgeException ex) { throw new McpRequestException(id, BridgeError, ex.Message); }
        catch (BridgeConnectTimeoutException ex)
        {
            // The request never reached VS, so a retry is safe for every profile. Evict the
            // cached instance and re-discover — a stale pipe (VS restarted) otherwise fails
            // on every call until the user manually rebinds.
            return await RetryAfterFailureAsync(id, request, ex, ignoreSolutionHint, timeoutProfile).ConfigureAwait(false);
        }
        catch (TimeoutException ex) { throw MapResponseTimeout(id, ex); }
        catch (UnauthorizedAccessException ex) when (ShouldRetry(timeoutProfile))
        {
            return await RetryAfterFailureAsync(id, request, ex, ignoreSolutionHint, timeoutProfile).ConfigureAwait(false);
        }
        catch (IOException ex) when (ShouldRetry(timeoutProfile))
        {
            return await RetryAfterFailureAsync(id, request, ex, ignoreSolutionHint, timeoutProfile).ConfigureAwait(false);
        }
        catch (UnauthorizedAccessException ex) { throw new McpRequestException(id, CommError, $"VS bridge communication failed: {ex.Message}"); }
        catch (IOException ex) { throw new McpRequestException(id, CommError, $"VS bridge communication failed: {ex.Message}"); }
    }

    private async Task<JsonObject> SendPipeAsync(JsonObject request, bool ignoreSolutionHint, ToolTimeoutProfile timeoutProfile)
    {
        BridgeInstance instance = await GetInstanceAsync(ignoreSolutionHint).ConfigureAwait(false);
        await using VsPipeClient client = await VsPipeClient.CreateAsync(
            instance.PipeName,
            GetCommandTimeoutMs(timeoutProfile, _timeoutOverrideMs),
            GetPipeGateTimeoutMs(timeoutProfile, _timeoutOverrideMs)).ConfigureAwait(false);
        return await client.SendAsync((JsonObject)request.DeepClone()).ConfigureAwait(false);
    }

    private async Task<JsonObject> RetryAfterFailureAsync(
        JsonNode? id,
        JsonObject request,
        Exception ex,
        bool ignoreSolutionHint,
        ToolTimeoutProfile timeoutProfile)
    {
        BridgeInstance? evicted = ClearCached();
        _ = evicted ?? throw new McpRequestException(id, CommError, $"VS bridge communication failed: {ex.Message}");

        // If the VS process is gone this is a session loss, not a transient pipe error.
        // Don't silently re-discover and bind to a different VS instance.
        if (VsDiscovery.IsProcessGone(evicted.ProcessId))
        {
            throw new McpRequestException(id, SessionLostError,
                $"VS session lost: '{evicted.Label}' (pid {evicted.ProcessId}) has closed. " +
                $"Call bridge_health to see available instances, then use bind_solution or bind_instance to reconnect.");
        }

        try
        {
            JsonObject response = await SendPipeAsync(request, ignoreSolutionHint, timeoutProfile).ConfigureAwait(false);
            return await FinalizePipeResponseAsync(response, request, ignoreSolutionHint, timeoutProfile).ConfigureAwait(false);
        }
        catch (BridgeException retryEx) { throw new McpRequestException(id, BridgeError, retryEx.Message); }
        catch (TimeoutException retryEx) { throw new McpRequestException(id, TimeoutError, $"Timed out: {retryEx.Message}"); }
        catch (Exception retryEx) when (retryEx is not null) { throw new McpRequestException(id, CommError, $"VS bridge retry failed: {retryEx.Message}"); }
    }

    private McpRequestException MapResponseTimeout(JsonNode? id, TimeoutException ex)
    {
        // The request reached VS, so don't re-send it and don't silently rebind to another
        // instance — VS is usually just busy. But if the VS process died mid-command, evict
        // the cached instance so the next call re-discovers instead of timing out against a
        // dead pipe again.
        BridgeInstance? current = CurrentInstance;
        if (current is not null && VsDiscovery.IsProcessGone(current.ProcessId))
        {
            ClearCached();
            return new McpRequestException(id, SessionLostError,
                $"VS session lost: '{current.Label}' (pid {current.ProcessId}) closed while the command was in flight. " +
                "Call bridge_health to see available instances, then use bind_solution or bind_instance to reconnect.");
        }

        return new McpRequestException(id, TimeoutError, $"Timed out: {ex.Message}");
    }

    private async Task<JsonObject> FinalizePipeResponseAsync(JsonObject response, JsonObject request, bool ignoreSolutionHint, ToolTimeoutProfile timeoutProfile)
    {
        response = await RetryImplicitBindingCancellationAsync(response, request, ignoreSolutionHint, timeoutProfile).ConfigureAwait(false);
        AttachPendingBindingNotice(response);
        RememberSolutionPath(response["Data"]?["solutionPath"]?.GetValue<string>());
        return response;
    }

    private async Task<BridgeInstance> GetInstanceAsync(bool ignoreSolutionHint)
    {
        BridgeInstanceSelector selectorSnapshot;
        lock (_gate)
        {
            if (_state.Cached is not null) return _state.Cached;
            selectorSnapshot = _state.Selector;
        }

        BridgeInstanceSelector effectiveSelector = ignoreSolutionHint
            ? new BridgeInstanceSelector
            {
                InstanceId = selectorSnapshot.InstanceId,
                ProcessId = selectorSnapshot.ProcessId,
                PipeName = selectorSnapshot.PipeName,
                SolutionHint = null,
            }
            : selectorSnapshot;

        BridgeInstance discovered = await VsDiscovery.SelectAsync(effectiveSelector, _discoveryMode).ConfigureAwait(false);

        lock (_gate)
        {
            if (_state.Cached is null && ReferenceEquals(_state.Selector, selectorSnapshot))
            {
                _state.Cached = discovered;
                if (!selectorSnapshot.HasAny)
                {
                    _state.PendingBindingNotice = $"Auto-bound to {discovered.Label}.";
                }
            }
        }

        return discovered;
    }

    internal BridgeInstance? DropCachedInstanceIfStale(IReadOnlyList<BridgeInstance> visibleInstances)
    {
        lock (_gate)
        {
            BridgeInstance? cached = _state.Cached;
            if (cached is null)
            {
                return null;
            }

            BridgeInstance? visibleMatch = visibleInstances.FirstOrDefault(visible =>
                string.Equals(visible.InstanceId, cached.InstanceId, StringComparison.OrdinalIgnoreCase)
                || string.Equals(visible.PipeName, cached.PipeName, StringComparison.OrdinalIgnoreCase)
                || visible.ProcessId == cached.ProcessId);
            if (visibleMatch is not null)
            {
                _state.Cached = visibleMatch;
                _state.LastSolutionPath = visibleMatch.SolutionPath;
                return null;
            }

            _state.Cached = null;
            _state.LastSolutionPath = string.Empty;
            _state.PendingBindingNotice = null;
            return cached;
        }
    }

    private BridgeInstance? ClearCached()
    {
        lock (_gate)
        {
            BridgeInstance? prev = _state.Cached;
            _state.Cached = null;
            return prev;
        }
    }

    private void RememberSolutionPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        lock (_gate)
        {
            _state.LastSolutionPath = path;
            if (_state.Cached is not null
                && !string.Equals(_state.Cached.SolutionPath, path, StringComparison.OrdinalIgnoreCase))
            {
                _state.Cached = _state.Cached with
                {
                    SolutionPath = path,
                    SolutionName = Path.GetFileName(path),
                };
            }
        }
    }

    private void AttachPendingBindingNotice(JsonObject response)
    {
        string? notice;
        lock (_gate)
        {
            notice = _state.PendingBindingNotice;
            _state.PendingBindingNotice = null;
        }

        if (string.IsNullOrWhiteSpace(notice))
        {
            return;
        }

        response["BindingNotice"] = notice;
        response["Binding"] = InstanceToJson(CurrentInstance ?? throw new InvalidOperationException("Current instance should exist when attaching a binding notice."));
        ToolCatalog.AttachBoundSessionGuidance(response);
    }

    private async Task<JsonObject> RetryImplicitBindingCancellationAsync(
        JsonObject response,
        JsonObject request,
        bool ignoreSolutionHint,
        ToolTimeoutProfile timeoutProfile)
    {
        bool shouldRetry;
        lock (_gate)
        {
            shouldRetry = !string.IsNullOrWhiteSpace(_state.PendingBindingNotice)
                && IsInterruptedOperationResponse(response);
            if (shouldRetry)
            {
                _state.PendingBindingNotice += " Retried once after the initial command was interrupted.";
            }
        }

        if (!shouldRetry)
        {
            return response;
        }

        JsonObject retryRequest = (JsonObject)request.DeepClone();
        retryRequest["id"] = Guid.NewGuid().ToString("N")[..8];
        return await SendPipeAsync(retryRequest, ignoreSolutionHint, timeoutProfile).ConfigureAwait(false);
    }

    private static JsonObject BuildRequest(string command, JsonNode? args) => new()
    {
        ["id"] = Guid.NewGuid().ToString("N")[..8],
        ["command"] = command,
        ["args"] = args?.DeepClone(),
    };

    private static JsonObject BuildBatchRequest(JsonArray steps, bool stopOnError)
    {
        JsonObject request = new()
        {
            ["id"] = Guid.NewGuid().ToString("N")[..8],
            ["command"] = "batch",
            ["batch"] = steps.DeepClone(),
        };

        if (stopOnError)
        {
            request["stopOnError"] = true;
        }

        return request;
    }

    private static bool IsInterruptedOperationResponse(JsonObject response)
    {
        bool success = response["Success"]?.GetValue<bool>() ?? false;
        if (success)
        {
            return false;
        }

        string? summary = response["Summary"]?.GetValue<string>();
        if (string.Equals(summary, "The operation was canceled.", StringComparison.Ordinal))
        {
            return true;
        }

        string? errorMessage = response["Error"]?["message"]?.GetValue<string>();
        return string.Equals(errorMessage, "Bridge server interrupted: The operation was canceled.", StringComparison.Ordinal);
    }

    // ── JSON helpers ───────────────────────────────────────────────────────────

    private static JsonObject InstanceToJson(BridgeInstance inst) => new()
    {
        ["instanceId"] = inst.InstanceId,
        ["pipeName"] = inst.PipeName,
        ["pid"] = inst.ProcessId,
        ["solutionPath"] = inst.SolutionPath,
        ["solutionName"] = inst.SolutionName,
        ["label"] = inst.Label,
        ["source"] = inst.Source,
        ["version"] = inst.Version ?? "pre-3.0.6",
    };

    private static JsonObject SelectorToJson(BridgeInstanceSelector sel) => new()
    {
        ["instanceId"] = sel.InstanceId,
        ["pid"] = sel.ProcessId,
        ["pipeName"] = sel.PipeName,
        ["solutionHint"] = sel.SolutionHint,
    };

}
