using System;
using System.Threading.Tasks;

namespace VsIdeBridge.Services;

internal sealed class BridgeWatchdogProbeLifecycleState
{
    public Task<double>? Task;
    public DateTimeOffset StartedAtUtc;
    public DateTimeOffset? LastStartedAtUtc;
    public DateTimeOffset? LastCompletedAtUtc;
    public bool TimeoutRecorded;
}

internal sealed class BridgeWatchdogProbeHealthState
{
    public DateTimeOffset? LastHealthyAtUtc;
    public DateTimeOffset? DegradedSinceUtc;
    public DateTimeOffset? LastTimeoutCommandAtUtc;
    public BridgeWatchdogProbeTimeoutCommandState LastProbeTimeoutCommand = new();
    public bool IsDegraded;
    public string DegradedReason = string.Empty;
    public string LastError = string.Empty;
    public int ConsecutiveUnhealthyProbes;
}

internal sealed class BridgeWatchdogProbeMetrics
{
    public double LastDurationMs;
    public double MaxDurationMs;
    public long TotalTimeouts;
    public long TotalBusyProbes;
    public long TotalFailures;
    public long SuccessfulCount;
}

internal sealed class BridgeWatchdogProbeState
{
    public BridgeWatchdogProbeLifecycleState Lifecycle = new();
    public BridgeWatchdogProbeHealthState Health = new();
    public BridgeWatchdogProbeMetrics Metrics = new();
}

internal sealed class BridgeWatchdogActiveCommandState
{
    public string Name = string.Empty;
    public string RequestId = string.Empty;
    public DateTimeOffset? StartedAtUtc;
}

internal sealed class BridgeWatchdogProbeTimeoutCommandState
{
    public string Name = string.Empty;
    public string RequestId = string.Empty;
    public DateTimeOffset? StartedAtUtc;
    public DateTimeOffset? DetectedAtUtc;
    public double ElapsedMs;
}

internal sealed class BridgeWatchdogLastCommandState
{
    public string Name = string.Empty;
    public string RequestId = string.Empty;
    public string ErrorCode = string.Empty;
    public bool? Success;
    public double DurationMs;
    public DateTimeOffset? CompletedAtUtc;
}

internal sealed class BridgeWatchdogCommandMetrics
{
    public long TotalCommands;
    public long SuccessfulCommands;
    public long FailedCommands;
    public long TimeoutCommands;
    public double SumDurationMs;
    public double MaxDurationMs;
}

internal sealed class BridgeWatchdogCommandState
{
    public BridgeWatchdogActiveCommandState Active = new();
    public BridgeWatchdogLastCommandState Last = new();
    public BridgeWatchdogCommandMetrics Metrics = new();
}
