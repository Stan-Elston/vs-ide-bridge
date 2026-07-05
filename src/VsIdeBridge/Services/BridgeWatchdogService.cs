using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace VsIdeBridge.Services;

internal sealed class BridgeWatchdogService(
    AsyncPackage package,
    int probeIntervalMilliseconds = BridgeWatchdogService.DefaultProbeIntervalMilliseconds,
    int probeTimeoutMilliseconds = BridgeWatchdogService.DefaultProbeTimeoutMilliseconds) : IDisposable
{
    private const int DefaultProbeIntervalMilliseconds = 1_000;
    private const int DefaultProbeTimeoutMilliseconds = 2_000;
    private const int CommandTimeoutDegradedWindowMilliseconds = 15_000;

    private readonly AsyncPackage _package = package;
    private readonly object _sync = new();
    private readonly CancellationTokenSource _cts = new();
    private readonly int _probeIntervalMilliseconds = probeIntervalMilliseconds > 0 ? probeIntervalMilliseconds : DefaultProbeIntervalMilliseconds;
    private readonly int _probeTimeoutMilliseconds = probeTimeoutMilliseconds > 0 ? probeTimeoutMilliseconds : DefaultProbeTimeoutMilliseconds;
    private readonly BridgeWatchdogProbeState _probe = new();
    private readonly BridgeWatchdogCommandState _command = new();

    private Task? _monitorTask;
    private DateTimeOffset _startedAtUtc;
    private bool _started;
    private bool _disposed;

    public void Start()
    {
        lock (_sync)
        {
            if (_started || _disposed)
            {
                return;
            }

            _started = true;
            _startedAtUtc = DateTimeOffset.UtcNow;
        }

        _monitorTask = Task.Factory.StartNew(
            () => MonitorLoopAsync(_cts.Token),
            _cts.Token,
            TaskCreationOptions.LongRunning,
            TaskScheduler.Default).Unwrap();
    }

    public void RecordCommandStarted(string commandName, string? requestId)
    {
        lock (_sync)
        {
            _command.Active.Name = commandName ?? string.Empty;
            _command.Active.RequestId = requestId ?? string.Empty;
            _command.Active.StartedAtUtc = DateTimeOffset.UtcNow;
        }
    }

    public void RecordCommandCompleted(string commandName, string? requestId, bool success, double durationMs, string? errorCode)
    {
        DateTimeOffset now = DateTimeOffset.UtcNow;
        lock (_sync)
        {
            _command.Metrics.TotalCommands++;
            _command.Metrics.SumDurationMs += durationMs;
            _command.Metrics.MaxDurationMs = Math.Max(_command.Metrics.MaxDurationMs, durationMs);

            if (success)
            {
                _command.Metrics.SuccessfulCommands++;
            }
            else
            {
                _command.Metrics.FailedCommands++;
                if (IsTimeoutError(errorCode))
                {
                    _command.Metrics.TimeoutCommands++;
                    _probe.Health.LastTimeoutCommandAtUtc = now;
                    if (!_probe.Health.IsDegraded)
                    {
                        _probe.Health.DegradedSinceUtc = now;
                    }

                    _probe.Health.IsDegraded = true;
                    _probe.Health.DegradedReason = "command_timeout";
                }
            }

            _command.Last.Name = commandName ?? string.Empty;
            _command.Last.RequestId = requestId ?? string.Empty;
            _command.Last.Success = success;
            _command.Last.ErrorCode = errorCode ?? string.Empty;
            _command.Last.DurationMs = durationMs;
            _command.Last.CompletedAtUtc = now;

            _command.Active.Name = string.Empty;
            _command.Active.RequestId = string.Empty;
            _command.Active.StartedAtUtc = null;
        }
    }

    public JObject GetSnapshot()
    {
        lock (_sync)
        {
            DateTimeOffset now = DateTimeOffset.UtcNow;
            bool probeStalled = _probe.Lifecycle.Task is { IsCompleted: false } &&
                _probe.Lifecycle.TimeoutRecorded &&
                _probe.Lifecycle.StartedAtUtc != default;
            double probeStalledForMs = probeStalled
                ? Math.Max(0, (now - _probe.Lifecycle.StartedAtUtc).TotalMilliseconds)
                : 0;
            // A stalled probe while a foreground command holds the UI thread is expected
            // busy behavior, not a health failure (field report issue 4).
            bool commandActive = _command.Active.StartedAtUtc is not null;
            bool busy = probeStalled && commandActive;
            bool timeoutWindowOpen = _probe.Health.LastTimeoutCommandAtUtc is not null &&
                (now - _probe.Health.LastTimeoutCommandAtUtc.Value).TotalMilliseconds <= CommandTimeoutDegradedWindowMilliseconds;
            bool degraded = _probe.Health.IsDegraded || (probeStalled && !commandActive) || timeoutWindowOpen;
            string degradedReason = _probe.Health.DegradedReason;
            if (probeStalled && !commandActive)
            {
                degradedReason = "watchdog_probe_timeout";
            }
            else if (timeoutWindowOpen && string.IsNullOrWhiteSpace(degradedReason))
            {
                degradedReason = "command_timeout_recent";
            }

            long totalCompletedCommands = _command.Metrics.SuccessfulCommands + _command.Metrics.FailedCommands;
            double averageCommandDurationMs = totalCompletedCommands > 0
                ? _command.Metrics.SumDurationMs / totalCompletedCommands
                : 0.0;

            return new JObject
            {
                ["startedAtUtc"] = _startedAtUtc == default ? JValue.CreateNull() : _startedAtUtc.ToString("O"),
                ["probeIntervalMs"] = _probeIntervalMilliseconds,
                ["probeTimeoutMs"] = _probeTimeoutMilliseconds,
                ["isDegraded"] = degraded,
                ["degradedReason"] = degradedReason,
                ["isBusy"] = busy,
                ["totalBusyProbes"] = _probe.Metrics.TotalBusyProbes,
                ["degradedSinceUtc"] = ToNullableTimestamp(_probe.Health.DegradedSinceUtc),
                ["lastProbeStartedAtUtc"] = ToNullableTimestamp(_probe.Lifecycle.LastStartedAtUtc),
                ["lastProbeCompletedAtUtc"] = ToNullableTimestamp(_probe.Lifecycle.LastCompletedAtUtc),
                ["lastHealthyAtUtc"] = ToNullableTimestamp(_probe.Health.LastHealthyAtUtc),
                ["lastProbeDurationMs"] = Math.Round(_probe.Metrics.LastDurationMs, 1),
                ["maxProbeDurationMs"] = Math.Round(_probe.Metrics.MaxDurationMs, 1),
                ["totalProbeTimeouts"] = _probe.Metrics.TotalTimeouts,
                ["totalProbeFailures"] = _probe.Metrics.TotalFailures,
                ["consecutiveUnhealthyProbes"] = _probe.Health.ConsecutiveUnhealthyProbes,
                ["probeStalled"] = probeStalled,
                ["probeStalledForMs"] = Math.Round(probeStalledForMs, 1),
                ["lastProbeError"] = _probe.Health.LastError,
                ["lastProbeTimeoutCommand"] = _probe.Health.LastProbeTimeoutCommand.DetectedAtUtc is null
                    ? JValue.CreateNull()
                    : new JObject
                    {
                        ["name"] = _probe.Health.LastProbeTimeoutCommand.Name,
                        ["requestId"] = _probe.Health.LastProbeTimeoutCommand.RequestId,
                        ["startedAtUtc"] = ToNullableTimestamp(_probe.Health.LastProbeTimeoutCommand.StartedAtUtc),
                        ["detectedAtUtc"] = ToNullableTimestamp(_probe.Health.LastProbeTimeoutCommand.DetectedAtUtc),
                        ["elapsedMs"] = Math.Round(_probe.Health.LastProbeTimeoutCommand.ElapsedMs, 1),
                    },
                ["activeCommand"] = _command.Active.StartedAtUtc is null
                    ? JValue.CreateNull()
                    : new JObject
                    {
                        ["name"] = _command.Active.Name,
                        ["requestId"] = _command.Active.RequestId,
                        ["startedAtUtc"] = _command.Active.StartedAtUtc.Value.ToString("O"),
                        ["elapsedMs"] = Math.Round((now - _command.Active.StartedAtUtc.Value).TotalMilliseconds, 1),
                    },
                ["lastCommand"] = _command.Last.CompletedAtUtc is null
                    ? JValue.CreateNull()
                    : new JObject
                    {
                        ["name"] = _command.Last.Name,
                        ["requestId"] = _command.Last.RequestId,
                        ["success"] = _command.Last.Success,
                        ["errorCode"] = _command.Last.ErrorCode,
                        ["durationMs"] = Math.Round(_command.Last.DurationMs, 1),
                        ["completedAtUtc"] = _command.Last.CompletedAtUtc.Value.ToString("O"),
                    },
                ["totals"] = new JObject
                {
                    ["commands"] = _command.Metrics.TotalCommands,
                    ["successfulCommands"] = _command.Metrics.SuccessfulCommands,
                    ["failedCommands"] = _command.Metrics.FailedCommands,
                    ["timeoutCommands"] = _command.Metrics.TimeoutCommands,
                    ["averageCommandDurationMs"] = Math.Round(averageCommandDurationMs, 1),
                    ["maxCommandDurationMs"] = Math.Round(_command.Metrics.MaxDurationMs, 1),
                },
            };
        }
    }

    public void Dispose()
    {
        lock (_sync)
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
        }

        _cts.Cancel();
        try
        {
            _monitorTask?.Wait(TimeSpan.FromSeconds(2));
        }
        catch (AggregateException ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);
        }

        _cts.Dispose();
    }

    private async Task MonitorLoopAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                await TickAsync(cancellationToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                break;
            }
            catch (Exception ex) when (ex is not null) // monitor loop: swallow probe failures to keep the watchdog alive
            {
                RecordProbeFailure(ex.Message);
            }

            try
            {
                await Task.Delay(_probeIntervalMilliseconds, cancellationToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                break;
            }
        }
    }

    private async Task TickAsync(CancellationToken cancellationToken)
    {
        Task<double>? completedProbe = null;
        bool completedProbeWasTimedOut = false;
        bool shouldStartProbe = false;
        bool shouldRecordTimeout = false;
        double timeoutElapsedMs = 0.0;
        DateTimeOffset now = DateTimeOffset.UtcNow;

        lock (_sync)
        {
            if (_probe.Lifecycle.Task is null)
            {
                shouldStartProbe = true;
            }
            else if (_probe.Lifecycle.Task.IsCompleted)
            {
                completedProbe = _probe.Lifecycle.Task;
                completedProbeWasTimedOut = _probe.Lifecycle.TimeoutRecorded;
                _probe.Lifecycle.Task = null;
                _probe.Lifecycle.TimeoutRecorded = false;
            }
            else
            {
                timeoutElapsedMs = Math.Max(0, (now - _probe.Lifecycle.StartedAtUtc).TotalMilliseconds);
                if (!_probe.Lifecycle.TimeoutRecorded && timeoutElapsedMs >= _probeTimeoutMilliseconds)
                {
                    _probe.Lifecycle.TimeoutRecorded = true;
                    shouldRecordTimeout = true;
                }
            }
        }

        if (shouldRecordTimeout)
        {
            RecordProbeTimeout(timeoutElapsedMs);
        }

        if (completedProbe is not null)
        {
            try
            {
                double durationMs = await completedProbe.ConfigureAwait(false);
                RecordProbeSuccess(durationMs, completedProbeWasTimedOut);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                return;
            }
            catch (Exception ex) when (ex is not null) // probe tick: swallow failures to keep the watchdog alive
            {
                RecordProbeFailure(ex.Message);
            }

            shouldStartProbe = true;
        }

        if (shouldStartProbe)
        {
            StartProbe(cancellationToken);
        }
    }

    private void StartProbe(CancellationToken cancellationToken)
    {
        DateTimeOffset startedAtUtc = DateTimeOffset.UtcNow;
        Task<double> task = _package.JoinableTaskFactory.RunAsync(async () =>
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            await _package.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            stopwatch.Stop();
            return stopwatch.Elapsed.TotalMilliseconds;
        }).Task;

        lock (_sync)
        {
            if (_probe.Lifecycle.Task is not null)
            {
                return;
            }

            _probe.Lifecycle.Task = task;
            _probe.Lifecycle.StartedAtUtc = startedAtUtc;
            _probe.Lifecycle.LastStartedAtUtc = startedAtUtc;
            _probe.Lifecycle.TimeoutRecorded = false;
        }
    }

    private void RecordProbeTimeout(double elapsedMs)
    {
        DateTimeOffset now = DateTimeOffset.UtcNow;
        lock (_sync)
        {
            bool foregroundCommandActive = _command.Active.StartedAtUtc is not null;
            double activeCommandElapsedMs = _command.Active.StartedAtUtc is null
                ? 0.0
                : Math.Max(0, (now - _command.Active.StartedAtUtc.Value).TotalMilliseconds);
            _probe.Health.LastProbeTimeoutCommand.Name = _command.Active.Name;
            _probe.Health.LastProbeTimeoutCommand.RequestId = _command.Active.RequestId;
            _probe.Health.LastProbeTimeoutCommand.StartedAtUtc = _command.Active.StartedAtUtc;
            _probe.Health.LastProbeTimeoutCommand.DetectedAtUtc = now;
            _probe.Health.LastProbeTimeoutCommand.ElapsedMs = activeCommandElapsedMs;

            if (foregroundCommandActive)
            {
                // The UI thread is legitimately held by a foreground command (build, long
                // diagnostics refresh). That is busy, not unhealthy (field report issue 4):
                // count it separately and leave the degraded state untouched.
                _probe.Metrics.TotalBusyProbes++;
                _probe.Health.LastError = $"UI responsiveness probe exceeded {_probeTimeoutMilliseconds} ms while foreground command " +
                    $"{_command.Active.Name} was running for {Math.Round(activeCommandElapsedMs, 1)} ms (busy, not a failure).";
                return;
            }

            _probe.Metrics.TotalTimeouts++;
            _probe.Health.ConsecutiveUnhealthyProbes++;
            _probe.Health.LastError = $"UI responsiveness probe exceeded {_probeTimeoutMilliseconds} ms (elapsed={Math.Round(elapsedMs, 1)} ms) with no bridge command running.";
            if (!_probe.Health.IsDegraded)
            {
                _probe.Health.DegradedSinceUtc = now;
            }

            _probe.Health.IsDegraded = true;
            _probe.Health.DegradedReason = "watchdog_probe_timeout";
        }
    }

    private void RecordProbeFailure(string message)
    {
        DateTimeOffset now = DateTimeOffset.UtcNow;
        lock (_sync)
        {
            _probe.Metrics.TotalFailures++;
            _probe.Health.ConsecutiveUnhealthyProbes++;
            _probe.Lifecycle.LastCompletedAtUtc = now;
            _probe.Health.LastError = message ?? string.Empty;
            if (!_probe.Health.IsDegraded)
            {
                _probe.Health.DegradedSinceUtc = now;
            }

            _probe.Health.IsDegraded = true;
            _probe.Health.DegradedReason = "watchdog_probe_failure";
        }
    }

    private void RecordProbeSuccess(double durationMs, bool recoveredFromTimeout)
    {
        DateTimeOffset now = DateTimeOffset.UtcNow;
        lock (_sync)
        {
            _probe.Lifecycle.LastCompletedAtUtc = now;
            _probe.Health.LastHealthyAtUtc = now;
            _probe.Metrics.LastDurationMs = durationMs;
            _probe.Metrics.SuccessfulCount++;
            if (_probe.Metrics.SuccessfulCount % 120 == 0)
            {
                _probe.Metrics.MaxDurationMs = durationMs;
            }

            _probe.Metrics.MaxDurationMs = Math.Max(_probe.Metrics.MaxDurationMs, durationMs);
            _probe.Health.LastError = string.Empty;
            _probe.Health.ConsecutiveUnhealthyProbes = 0;

            bool timeoutWindowOpen = _probe.Health.LastTimeoutCommandAtUtc is not null &&
                (now - _probe.Health.LastTimeoutCommandAtUtc.Value).TotalMilliseconds <= CommandTimeoutDegradedWindowMilliseconds;
            if (!timeoutWindowOpen)
            {
                _probe.Health.IsDegraded = false;
                _probe.Health.DegradedReason = string.Empty;
                _probe.Health.DegradedSinceUtc = null;
            }
            else if (recoveredFromTimeout && string.Equals(_probe.Health.DegradedReason, "watchdog_probe_timeout", StringComparison.Ordinal))
            {
                _probe.Health.DegradedReason = "command_timeout_recent";
            }
        }
    }

    private static bool IsTimeoutError(string? errorCode)
    {
        return string.Equals(errorCode, "timeout", StringComparison.OrdinalIgnoreCase)
            || string.Equals(errorCode, "ide_blocked", StringComparison.OrdinalIgnoreCase);
    }

    private static JToken ToNullableTimestamp(DateTimeOffset? value)
    {
        return value is null ? JValue.CreateNull() : value.Value.ToString("O");
    }
}
