using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed class BridgeInstanceService
{
    public BridgeInstanceService()
    {
        Process currentProcess = Process.GetCurrentProcess();
        ProcessId = currentProcess.Id;
        ProcessStartedAtUtc = currentProcess.StartTime.ToUniversalTime();
        PipeName = $"VsIdeBridge18_{ProcessId}";
        InstanceId = $"vs18-{ProcessId}-{ProcessStartedAtUtc:yyyyMMddTHHmmssZ}";
        Version = ResolveBridgeVersion();
    }

    public string InstanceId { get; }

    public int ProcessId { get; }

    public DateTime ProcessStartedAtUtc { get; }

    public string PipeName { get; }

    public string Version { get; }

    public object CreateDiscoveryRecord(string? solutionPath)
    {
        string normalizedSolutionPath = NormalizeSolutionPath(solutionPath);
        string label = BuildInstanceLabel(normalizedSolutionPath, ProcessId, PipeName);
        return new
        {
            instanceId = InstanceId,
            pid = ProcessId,
            startedAtUtc = ProcessStartedAtUtc.ToString("O"),
            pipeName = PipeName,
            solutionPath = normalizedSolutionPath,
            solutionName = GetSolutionName(normalizedSolutionPath),
            label,
            version = Version,
        };
    }

    public JObject CreateStateData(string? solutionPath)
    {
        string normalizedSolutionPath = NormalizeSolutionPath(solutionPath);
        string label = BuildInstanceLabel(normalizedSolutionPath, ProcessId, PipeName);
        return new JObject
        {
            ["instanceId"] = InstanceId,
            ["pid"] = ProcessId,
            ["startedAtUtc"] = ProcessStartedAtUtc.ToString("O"),
            ["pipeName"] = PipeName,
            ["solutionPath"] = normalizedSolutionPath,
            ["solutionName"] = GetSolutionName(normalizedSolutionPath),
            ["label"] = label,
            ["version"] = Version,
        };
    }

    private static string ResolveBridgeVersion()
    {
        string informational = typeof(BridgeInstanceService).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? string.Empty;
        int metadataStart = informational.IndexOf('+');
        if (metadataStart > 0)
        {
            informational = informational.Substring(0, metadataStart);
        }

        return string.IsNullOrWhiteSpace(informational)
            ? typeof(BridgeInstanceService).Assembly.GetName().Version?.ToString() ?? "unknown"
            : informational;
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

    private static string BuildInstanceLabel(string solutionPath, int processId, string pipeName)
    {
        string solutionBaseName = string.IsNullOrWhiteSpace(solutionPath)
            ? string.Empty
            : Path.GetFileNameWithoutExtension(solutionPath);

        string name = string.IsNullOrWhiteSpace(solutionBaseName)
            ? (string.IsNullOrWhiteSpace(pipeName) ? "Visual Studio" : pipeName)
            : solutionBaseName;

        return $"{name} ({processId})";
    }
}
