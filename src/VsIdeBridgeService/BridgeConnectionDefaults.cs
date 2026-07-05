namespace VsIdeBridgeService;

internal static class BridgeConnectionDefaults
{
    public const int FastTimeoutMs = 15_000;
    public const int InteractiveTimeoutMs = 45_000;
    public const int HeavyTimeoutMs = 130_000;
    public const int BuildWaitTimeoutMs = 610_000;
    public const int FastPipeGateTimeoutMs = 750;
    public const int InteractivePipeGateTimeoutMs = 2_000;
    public const int HeavyPipeGateTimeoutMs = 5_000;
    public const int BridgeError = -32001;
    public const int TimeoutError = -32002;
    public const int CommError = -32003;
    public const int UnboundError = -32004;
    public const int SessionLostError = -32005;

    public static int GetCommandTimeoutMs(BridgeConnection.ToolTimeoutProfile timeoutProfile, int? timeoutOverrideMs)
    {
        return timeoutOverrideMs ?? timeoutProfile switch
        {
            BridgeConnection.ToolTimeoutProfile.Fast => FastTimeoutMs,
            BridgeConnection.ToolTimeoutProfile.Interactive => InteractiveTimeoutMs,
            BridgeConnection.ToolTimeoutProfile.Heavy => HeavyTimeoutMs,
            BridgeConnection.ToolTimeoutProfile.BuildWait => BuildWaitTimeoutMs,
            _ => InteractiveTimeoutMs,
        };
    }

    public static int GetPipeGateTimeoutMs(BridgeConnection.ToolTimeoutProfile timeoutProfile, int? timeoutOverrideMs)
    {
        int pipeGateTimeoutMs = timeoutProfile switch
        {
            BridgeConnection.ToolTimeoutProfile.Fast => FastPipeGateTimeoutMs,
            BridgeConnection.ToolTimeoutProfile.Interactive => InteractivePipeGateTimeoutMs,
            BridgeConnection.ToolTimeoutProfile.Heavy or BridgeConnection.ToolTimeoutProfile.BuildWait => HeavyPipeGateTimeoutMs,
            _ => InteractivePipeGateTimeoutMs,
        };

        return timeoutOverrideMs is int overrideMs
            ? Math.Min(pipeGateTimeoutMs, overrideMs)
            : pipeGateTimeoutMs;
    }

    public static bool ShouldRetry(BridgeConnection.ToolTimeoutProfile timeoutProfile)
        => timeoutProfile != BridgeConnection.ToolTimeoutProfile.Fast;
}
