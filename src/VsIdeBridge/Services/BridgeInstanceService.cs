using System;
using System.Diagnostics;
using System.IO;
using Newtonsoft.Json.Linq;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed class BridgeInstanceService
{
    public BridgeInstanceService()
    {
        var process = Process.GetCurrentProcess();
        ProcessId = process.Id;
        ProcessStartedAtUtc = process.StartTime.ToUniversalTime();
        PipeName = $"VsIdeBridge18_{ProcessId}";
        InstanceId = $"vs18-{ProcessId}-{ProcessStartedAtUtc:yyyyMMddTHHmmssZ}";
    }

    public string InstanceId { get; }

    public int ProcessId { get; }

    public DateTime ProcessStartedAtUtc { get; }

    public string PipeName { get; }

    public object CreateDiscoveryRecord(string? solutionPath)
    {
        var normalizedSolutionPath = NormalizeSolutionPath(solutionPath);
        return new
        {
            instanceId = InstanceId,
            pid = ProcessId,
            startedAtUtc = ProcessStartedAtUtc.ToString("O"),
            pipeName = PipeName,
            solutionPath = normalizedSolutionPath,
            solutionName = GetSolutionName(normalizedSolutionPath),
        };
    }

    public JObject CreateStateData(string? solutionPath)
    {
        var normalizedSolutionPath = NormalizeSolutionPath(solutionPath);
        return new JObject
        {
            ["instanceId"] = InstanceId,
            ["pid"] = ProcessId,
            ["startedAtUtc"] = ProcessStartedAtUtc.ToString("O"),
            ["pipeName"] = PipeName,
            ["solutionPath"] = normalizedSolutionPath,
            ["solutionName"] = GetSolutionName(normalizedSolutionPath),
        };
    }

    private static string NormalizeSolutionPath(string? solutionPath)
    {
        if (string.IsNullOrWhiteSpace(solutionPath))
        {
            return string.Empty;
        }

        return PathNormalization.NormalizeFilePath(solutionPath);
    }

    private static string GetSolutionName(string solutionPath)
    {
        return string.IsNullOrWhiteSpace(solutionPath)
            ? string.Empty
            : Path.GetFileName(solutionPath);
    }
}
