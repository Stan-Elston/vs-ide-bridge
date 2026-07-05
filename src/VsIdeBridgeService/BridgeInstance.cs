namespace VsIdeBridgeService;

// Describes a discovered VS IDE Bridge instance.
internal sealed record BridgeInstance
{
    public required string InstanceId { get; init; }
    public required string PipeName { get; init; }
    public required int ProcessId { get; init; }
    public required string SolutionPath { get; init; }
    public required string SolutionName { get; init; }
    public required string Label { get; init; }
    public required string Source { get; init; }
    public string? StartedAtUtc { get; init; }
    // VSIX version stamped into the discovery record; null for pre-3.0.6 bridges.
    public string? Version { get; init; }
    public required string DiscoveryFile { get; init; }
    public required DateTime LastWriteTimeUtc { get; init; }
}

// Criteria for selecting one instance from the discovered list.
internal sealed class BridgeInstanceSelector
{
    public string? InstanceId { get; init; }
    public int? ProcessId { get; init; }
    public string? PipeName { get; init; }
    public string? SolutionHint { get; init; }

    public bool HasAny =>
        !string.IsNullOrWhiteSpace(InstanceId)
        || ProcessId.HasValue
        || !string.IsNullOrWhiteSpace(PipeName)
        || !string.IsNullOrWhiteSpace(SolutionHint);

    public string Describe()
    {
        System.Collections.Generic.List<string> parts = [];
        if (!string.IsNullOrWhiteSpace(InstanceId)) parts.Add($"instance_id={InstanceId}");
        if (ProcessId.HasValue) parts.Add($"pid={ProcessId}");
        if (!string.IsNullOrWhiteSpace(PipeName)) parts.Add($"pipe={PipeName}");
        if (!string.IsNullOrWhiteSpace(SolutionHint)) parts.Add($"solution={SolutionHint}");
        return parts.Count > 0 ? string.Join(", ", parts) : "(no criteria)";
    }
}

internal enum DiscoveryMode { MemoryFirst, JsonOnly, Hybrid }
