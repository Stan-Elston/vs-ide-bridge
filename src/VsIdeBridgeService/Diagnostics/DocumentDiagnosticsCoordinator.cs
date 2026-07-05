using System.Text.Json.Nodes;
using VsIdeBridge.Diagnostics;

using static VsIdeBridgeService.ArgBuilder;

namespace VsIdeBridgeService.Diagnostics;

internal sealed class DocumentDiagnosticsCoordinator(BridgeConnection bridge)
{
    private static readonly TimeSpan RefreshDebounceInterval = TimeSpan.FromSeconds(2);
    private static readonly TimeSpan RefreshCompletionTimeout = TimeSpan.FromSeconds(10);
    private const string ServiceCacheBypassArg = "service_cache_bypass";

    private readonly BridgeConnection _bridge = bridge;
    private readonly object _gate = new();
    private readonly CachedDiagnosticsState _cached = new();

    private Task? _refreshTask;
    private bool _refreshRequested;
    private readonly RefreshTimingState _timing = new();

    public JsonObject QueueRefreshAndGetSnapshot(string reason, bool clearCached = false)
    {
        DateTimeOffset now = DateTimeOffset.UtcNow;

        lock (_gate)
        {
            if (clearCached)
            {
                _cached.Errors = null;
                _cached.Warnings = null;
                _cached.Messages = null;
                _cached.LastError = null;
                _cached.Status = "idle";
            }

            if (_refreshTask is not null && !_refreshTask.IsCompleted)
            {
                if (clearCached)
                {
                    _refreshRequested = true;
                    _cached.Reason = reason;
                    _timing.LastQueuedUtc = now;
                }

                return CreateSnapshotLocked().ToJson();
            }

            if (!clearCached && _timing.LastCompletedUtc is not null && now - _timing.LastCompletedUtc < RefreshDebounceInterval)
            {
                return CreateSnapshotLocked().ToJson();
            }

            _refreshRequested = true;
            _cached.Reason = reason;
            _timing.LastQueuedUtc = now;

            if (_refreshTask is null || _refreshTask.IsCompleted)
            {
                _cached.Status = "queued";
                _refreshTask = Task.Run(RefreshLoopAsync);
            }

            return CreateSnapshotLocked().ToJson();
        }
    }

    public async Task<JsonObject> QueueRefreshAndWaitForSnapshotAsync(string reason, bool clearCached = false)
    {
        _ = QueueRefreshAndGetSnapshot(reason, clearCached);

        Task? refreshTask;
        lock (_gate)
        {
            refreshTask = _refreshTask;
        }

        if (refreshTask is not null)
        {
            _ = await Task.WhenAny(refreshTask, Task.Delay(RefreshCompletionTimeout)).ConfigureAwait(false);
        }

        lock (_gate)
        {
            return CreateSnapshotLocked().ToJson();
        }
    }

    public void Invalidate(string reason)
    {
        lock (_gate)
        {
            _cached.Errors = null;
            _cached.Warnings = null;
            _cached.Messages = null;
            _cached.LastError = null;
            _cached.Status = "idle";
            _cached.Reason = reason;
            _timing.LastQueuedUtc = null;
            _timing.LastStartedUtc = null;
            _timing.LastCompletedUtc = null;
        }
    }

    public bool TryGetCached(string severity, JsonObject? args, out JsonObject response)
    {
        lock (_gate)
        {
            DiagnosticBucket? bucket = GetCachedBucket(severity, out string kind);
            if (bucket is null || !CanServeCachedDiagnostics(args, severity) || !HasUsableCachedDiagnosticsLocked())
            {
                response = [];
                return false;
            }

            response = CreateCachedResponseLocked(bucket, kind);
            return true;
        }
    }

    public void RecordLiveRead(string severity, string kind, JsonObject response, bool replaceCachedBucket, bool invalidateCachedBucket)
    {
        if (response["Success"]?.GetValue<bool>() != true)
        {
            return;
        }

        lock (_gate)
        {
            if (replaceCachedBucket
                && !IsStaleDiagnosticsResponse(response)
                && response["Data"] is JsonObject data
                && data[DiagnosticJsonNames.Rows] is JsonArray)
            {
                SetCachedBucket(severity, DiagnosticBucket.FromResponse(response));
                _cached.Status = "completed";
                _cached.Reason = $"live-{kind}";
                _cached.LastError = null;
                _timing.LastCompletedUtc = DateTimeOffset.UtcNow;
                return;
            }

            if (invalidateCachedBucket)
            {
                ClearCachedBucket(severity);
                _cached.Status = "completed";
                _cached.Reason = $"filtered-live-{kind}";
                _cached.LastError = null;
                _timing.LastCompletedUtc = DateTimeOffset.UtcNow;
            }
        }
    }

    private DiagnosticBucket? GetCachedBucket(string severity, out string kind)
    {
        if (string.Equals(severity, "Error", StringComparison.OrdinalIgnoreCase))
        {
            kind = "errors";
            return _cached.Errors;
        }

        if (string.Equals(severity, "Warning", StringComparison.OrdinalIgnoreCase))
        {
            kind = "warnings";
            return _cached.Warnings;
        }

        kind = "messages";
        return string.Equals(severity, "Message", StringComparison.OrdinalIgnoreCase)
            ? _cached.Messages
            : null;
    }

    private void SetCachedBucket(string severity, DiagnosticBucket bucket)
    {
        if (string.Equals(severity, "Error", StringComparison.OrdinalIgnoreCase))
        {
            _cached.Errors = bucket;
            return;
        }

        if (string.Equals(severity, "Warning", StringComparison.OrdinalIgnoreCase))
        {
            _cached.Warnings = bucket;
            return;
        }

        if (string.Equals(severity, "Message", StringComparison.OrdinalIgnoreCase))
        {
            _cached.Messages = bucket;
        }
    }

    private void ClearCachedBucket(string severity)
    {
        if (string.Equals(severity, "Error", StringComparison.OrdinalIgnoreCase))
        {
            _cached.Errors = null;
            return;
        }

        if (string.Equals(severity, "Warning", StringComparison.OrdinalIgnoreCase))
        {
            _cached.Warnings = null;
            return;
        }

        if (string.Equals(severity, "Message", StringComparison.OrdinalIgnoreCase))
        {
            _cached.Messages = null;
        }
    }

    private static bool IsStaleDiagnosticsResponse(JsonObject response)
        => response["Cache"] is JsonObject cache
            && cache["mayBeStale"]?.GetValue<bool>() == true;

    private async Task RefreshLoopAsync()
    {
        while (true)
        {
            lock (_gate)
            {
                if (!_refreshRequested)
                {
                    _refreshTask = null;
                    return;
                }

                _refreshRequested = false;
                _cached.Status = "running";
                _timing.LastStartedUtc = DateTimeOffset.UtcNow;
                _cached.LastError = null;
            }

            try
            {
                JsonObject errors = await _bridge.SendAsync(null, "errors", BuildCachedListArgs("Error"))
                    .ConfigureAwait(false);
                JsonObject warnings = await _bridge.SendAsync(null, "warnings", BuildCachedListArgs("Warning"))
                    .ConfigureAwait(false);
                JsonObject messages = await _bridge.SendAsync(null, "messages", BuildCachedListArgs("Message"))
                    .ConfigureAwait(false);

                lock (_gate)
                {
                    _cached.Errors = DiagnosticBucket.FromResponse(errors);
                    _cached.Warnings = DiagnosticBucket.FromResponse(warnings);
                    _cached.Messages = DiagnosticBucket.FromResponse(messages);
                    _cached.Status = "completed";
                    _timing.LastCompletedUtc = DateTimeOffset.UtcNow;
                }
            }
            catch (Exception ex) when (ex is not null) // background diagnostics loop boundary
            {
                lock (_gate)
                {
                    _cached.Status = "failed";
                    _cached.LastError = ex.Message;
                    _timing.LastCompletedUtc = DateTimeOffset.UtcNow;
                }
            }
        }
    }

    private static bool CanServeCachedDiagnostics(JsonObject? args, string expectedSeverity)
    {
        if (args?[ServiceCacheBypassArg]?.GetValue<bool>() == true)
        {
            return false;
        }

        if (!WantsPassiveDiagnosticsRead(args))
        {
            return false;
        }

        if (args?["refresh"]?.GetValue<bool>() == true)
        {
            return false;
        }

        string? requestedSeverity = args?["severity"]?.GetValue<string>();
        if (!string.IsNullOrWhiteSpace(requestedSeverity)
            && !string.Equals(requestedSeverity, expectedSeverity, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (LooksLikeBridgeHandleFilter(GetOptionalString(args, "file"))
            || LooksLikeBridgeHandleFilter(GetOptionalString(args, "path")))
        {
            return false;
        }

        // quick and wait_for_intellisense are timing hints. Content filters are
        // applied by DiagnosticCollection before the tool boundary serializes JSON.
        return (args?["severity"] is null
            || string.Equals(requestedSeverity, expectedSeverity, StringComparison.OrdinalIgnoreCase));
    }

    private static bool LooksLikeBridgeHandleFilter(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        string trimmed = value.Trim();
        if (trimmed.Length < 3 || trimmed[1] != ':' || !char.IsLetter(trimmed[0]))
        {
            return false;
        }

        for (int i = 2; i < trimmed.Length; i++)
        {
            if (!char.IsDigit(trimmed[i]))
            {
                return false;
            }
        }

        return true;
    }

    private static string? GetOptionalString(JsonObject? args, string key)
    {
        return args?[key] is JsonValue value && value.TryGetValue(out string? text)
            ? text
            : null;
    }

    private static bool WantsPassiveDiagnosticsRead(JsonObject? args)
    {
        if (args?["quick"] is JsonNode quickNode)
        {
            return quickNode.GetValue<bool>();
        }

        return args?["refresh"]?.GetValue<bool>() != true;
    }

    private bool HasUsableCachedDiagnosticsLocked()
    {
        return string.Equals(_cached.Status, "completed", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(_cached.Reason, "startup", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(_cached.Reason, "bind", StringComparison.OrdinalIgnoreCase);
    }

    private DocumentDiagnosticsSnapshot CreateSnapshotLocked()
    {
        return new DocumentDiagnosticsSnapshot
        {
            Status = _cached.Status,
            Reason = _cached.Reason,
            LastError = _cached.LastError,
            Timing = new DocumentDiagnosticsTimingSnapshot
            {
                LastQueuedUtc = FormatUtc(_timing.LastQueuedUtc),
                LastStartedUtc = FormatUtc(_timing.LastStartedUtc),
                LastCompletedUtc = FormatUtc(_timing.LastCompletedUtc),
            },
            Results = new DocumentDiagnosticsResultSnapshot
            {
                Errors = _cached.Errors,
                Warnings = _cached.Warnings,
                Messages = _cached.Messages,
            },
        };
    }

    private JsonObject CreateCachedResponseLocked(DiagnosticBucket bucket, string kind)
    {
        JsonObject clone = bucket.ToResponseJson();
        JsonArray warnings = clone["Warnings"] as JsonArray is { } existingWarnings
            ? (JsonArray)existingWarnings.DeepClone()
            : [];
        warnings.Add("Using the passive diagnostics cache. This list may be stale relative to the current Visual Studio Error List. Use refresh=true for a fresh UI read.");
        clone["Warnings"] = warnings;
        clone["Cache"] = new JsonObject
        {
            ["source"] = "service-memory",
            ["kind"] = kind,
            ["mayBeStale"] = true,
            ["capturedAtUtc"] = FormatUtc(_timing.LastCompletedUtc),
            ["ageMs"] = _timing.LastCompletedUtc is DateTimeOffset completedUtc
                ? Math.Max(0, (DateTimeOffset.UtcNow - completedUtc).TotalMilliseconds)
                : null,
            ["snapshot"] = CreateSnapshotLocked().ToJson(),
        };
        return clone;
    }

    private static JsonObject BuildCachedListArgs(string severity)
        => new()
        {
            ["quick"] = true,
            ["wait_for_intellisense"] = false,
            ["severity"] = severity,
            [ServiceCacheBypassArg] = true,
        };

    private static string? FormatUtc(DateTimeOffset? value)
    {
        return value?.UtcDateTime.ToString("O");
    }

    private sealed class RefreshTimingState
    {
        public DateTimeOffset? LastQueuedUtc { get; set; }
        public DateTimeOffset? LastStartedUtc { get; set; }
        public DateTimeOffset? LastCompletedUtc { get; set; }
    }

    private sealed class CachedDiagnosticsState
    {
        public string Status { get; set; } = "idle";
        public string? Reason { get; set; }
        public string? LastError { get; set; }
        public DiagnosticBucket? Errors { get; set; }
        public DiagnosticBucket? Warnings { get; set; }
        public DiagnosticBucket? Messages { get; set; }
    }
}
