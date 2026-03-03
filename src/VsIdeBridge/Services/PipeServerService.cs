using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using VsIdeBridge.Commands;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

/// <summary>
/// Persistent named pipe server that eliminates per-call PowerShell overhead (~1500 ms → ~50 ms).
/// Discovery file: %TEMP%\vs-ide-bridge\pipes\bridge-{pid}.json
/// Protocol: newline-delimited JSON, one request per line, one response per line.
/// </summary>
internal sealed class PipeServerService : IDisposable
{
    private readonly VsIdeBridgePackage _package;
    private readonly IdeBridgeRuntime _runtime;
    private readonly string _discoveryFile;
    private readonly CancellationTokenSource _cts = new CancellationTokenSource();
    private Task? _listenTask;

    public PipeServerService(VsIdeBridgePackage package, IdeBridgeRuntime runtime)
    {
        _package = package;
        _runtime = runtime;
        var discoveryDir = Path.Combine(Path.GetTempPath(), "vs-ide-bridge", "pipes");
        Directory.CreateDirectory(discoveryDir);
        _discoveryFile = Path.Combine(discoveryDir, $"bridge-{runtime.BridgeInstanceService.ProcessId}.json");
    }

    public void Start()
    {
        PurgeStaleDiscoveryFiles();
        WriteDiscoveryFile(string.Empty);
        _ = _package.JoinableTaskFactory.RunAsync(async delegate
        {
            try
            {
                await _package.JoinableTaskFactory.SwitchToMainThreadAsync(_cts.Token);
                var dte = await _package.GetServiceAsync(typeof(SDTE)).ConfigureAwait(true) as DTE2;
                if (dte != null)
                {
                    WriteDiscoveryFile(GetSolutionPath(dte));
                }
            }
            catch (Exception ex)
            {
                ActivityLog.LogWarning(nameof(PipeServerService), $"Initial discovery refresh failed: {ex.Message}");
            }
        });
        _listenTask = Task.Run(() => ListenLoopAsync(_cts.Token));
    }

    private void PurgeStaleDiscoveryFiles()
    {
        var discoveryDir = Path.GetDirectoryName(_discoveryFile)!;
        try
        {
            foreach (var file in Directory.GetFiles(discoveryDir, "bridge-*.json"))
            {
                if (string.Equals(file, _discoveryFile, StringComparison.OrdinalIgnoreCase))
                    continue;

                var stem = Path.GetFileNameWithoutExtension(file); // "bridge-12345"
                var dash = stem.LastIndexOf('-');
                if (dash >= 0 && int.TryParse(stem.Substring(dash + 1), out var pid))
                {
                    try { Process.GetProcessById(pid); }
                    catch (ArgumentException) { File.Delete(file); } // process gone
                }
            }
        }
        catch (Exception ex)
        {
            ActivityLog.LogWarning(nameof(PipeServerService), $"Failed to purge stale discovery files: {ex.Message}");
        }
    }

    private void WriteDiscoveryFile(string? solutionPath)
    {
        try
        {
            var discoveryJson = JsonConvert.SerializeObject(_runtime.BridgeInstanceService.CreateDiscoveryRecord(solutionPath));
            File.WriteAllText(_discoveryFile, discoveryJson, new UTF8Encoding(false));
        }
        catch (Exception ex)
        {
            ActivityLog.LogWarning(nameof(PipeServerService), $"Failed to update discovery file: {ex.Message}");
        }
    }

    private static string GetSolutionPath(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            return dte.Solution?.FullName ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static JArray BuildBatchSteps(PipeRequest request)
    {
        var steps = new JArray();
        if (request.Batch == null)
        {
            return steps;
        }

        foreach (var batchRequest in request.Batch)
        {
            steps.Add(new JObject
            {
                ["id"] = batchRequest.Id is null ? JValue.CreateNull() : batchRequest.Id,
                ["command"] = batchRequest.Command ?? string.Empty,
                ["args"] = batchRequest.Args ?? string.Empty,
            });
        }

        return steps;
    }

    private static bool ShouldRevealActivity(string commandName)
    {
        return string.Equals(commandName, "Tools.IdeApplyUnifiedDiff", StringComparison.OrdinalIgnoreCase)
            || string.Equals(commandName, "apply-diff", StringComparison.OrdinalIgnoreCase)
            || commandName.IndexOf("Build", StringComparison.OrdinalIgnoreCase) >= 0
            || commandName.IndexOf("Debug", StringComparison.OrdinalIgnoreCase) >= 0
            || commandName.IndexOf("Open", StringComparison.OrdinalIgnoreCase) >= 0
            || commandName.IndexOf("Close", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private async Task ListenLoopAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            NamedPipeServerStream pipe;
            try
            {
                pipe = new NamedPipeServerStream(
                    _runtime.BridgeInstanceService.PipeName,
                    PipeDirection.InOut,
                    NamedPipeServerStream.MaxAllowedServerInstances,
                    PipeTransmissionMode.Byte,
                    PipeOptions.Asynchronous);
            }
            catch (Exception ex)
            {
                ActivityLog.LogError(nameof(PipeServerService), $"Failed to create pipe server instance: {ex.Message}");
                return;
            }

            try
            {
                await pipe.WaitForConnectionAsync(ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                pipe.Dispose();
                break;
            }
            catch (Exception ex)
            {
                ActivityLog.LogError(nameof(PipeServerService), $"Pipe accept error: {ex.Message}");
                pipe.Dispose();
                continue;
            }

            // Fire-and-forget: handle each connection on the thread pool
            _ = HandleConnectionAsync(pipe, ct);
        }
    }

    private async Task HandleConnectionAsync(NamedPipeServerStream pipe, CancellationToken ct)
    {
        using (pipe)
        {
            try
            {
                var reader = new StreamReader(pipe, Encoding.UTF8, detectEncodingFromByteOrderMarks: false, bufferSize: 4096, leaveOpen: true);
                var writer = new StreamWriter(pipe, new UTF8Encoding(false), bufferSize: 4096, leaveOpen: true)
                {
                    AutoFlush = true,
                    NewLine = "\n",
                };

                while (!ct.IsCancellationRequested && pipe.IsConnected)
                {
                    string? line;
                    try
                    {
                        line = await reader.ReadLineAsync().ConfigureAwait(false);
                    }
                    catch
                    {
                        break; // client disconnected mid-read
                    }

                    if (line == null) break; // clean EOF

                    var responseLine = await ExecuteRequestAsync(line, ct).ConfigureAwait(false);
                    await writer.WriteLineAsync(responseLine).ConfigureAwait(false);
                }
            }
            catch (Exception ex)
            {
                ActivityLog.LogError(nameof(PipeServerService), $"Pipe connection error: {ex.Message}");
            }
        }
    }

    private async Task<string> ExecuteRequestAsync(string requestJson, CancellationToken ct)
    {
        var startedAt = DateTimeOffset.UtcNow;
        string commandName = "";
        string? requestId = null;
        IdeCommandContext? failureContext = null;

        try
        {
            var request = JsonConvert.DeserializeObject<PipeRequest>(requestJson);
            if (request == null)
                throw new CommandErrorException("invalid_request", "Could not parse request JSON.");

            requestId = request.Id;
            var hasBatch = request.Batch is { Count: > 0 };
            commandName = hasBatch
                ? (!string.IsNullOrWhiteSpace(request.Command) ? request.Command : "Tools.IdeBatchCommands")
                : (request.Command ?? string.Empty);

            CommandExecutionResult result = null!;
            await _package.JoinableTaskFactory.RunAsync(async delegate
            {
                await _package.JoinableTaskFactory.SwitchToMainThreadAsync(ct);
                var dte = await _package.GetServiceAsync(typeof(SDTE)).ConfigureAwait(true) as DTE2;
                Assumes.Present(dte);
                WriteDiscoveryFile(GetSolutionPath(dte!));
                var ctx = new IdeCommandContext(_package, dte!, _runtime.Logger, _runtime, ct);
                failureContext = ctx;
                await _runtime.Logger.LogAsync($"IDE Bridge: {commandName} requested", ct).ConfigureAwait(true);

                if (hasBatch)
                {
                    var steps = BuildBatchSteps(request);
                    result = await IdeCoreCommands.ExecuteBatchAsync(ctx, steps, request.StopOnError ?? false).ConfigureAwait(true);
                    return;
                }

                if (!_runtime.TryGetCommand(commandName, out var cmd))
                    throw new CommandErrorException("command_not_found", $"Unknown command: '{commandName}'.");

                var args = CommandArgumentParser.Parse(request.Args);
                result = await cmd.ExecuteDirectAsync(ctx, args).ConfigureAwait(true);
            });

            await _runtime.Logger.LogAsync(
                $"IDE Bridge: {commandName} OK - {result.Summary}",
                ct,
                activatePane: ShouldRevealActivity(commandName)).ConfigureAwait(false);
            var envelope = BuildEnvelope(commandName, requestId, true, result.Summary, result.Data, result.Warnings, null, startedAt);
            return JsonConvert.SerializeObject(envelope);
        }
        catch (CommandErrorException ex)
        {
            await _runtime.Logger.LogAsync($"IDE Bridge: {commandName} FAIL - {ex.Code}", ct, activatePane: true).ConfigureAwait(false);
            var errorObj = new { code = ex.Code, message = ex.Message, details = ex.Details };
            var failureData = await _runtime.FailureContextService.CaptureAsync(failureContext).ConfigureAwait(false);
            var envelope = BuildEnvelope(commandName, requestId, false, ex.Message, failureData, new JArray(), errorObj, startedAt);
            return JsonConvert.SerializeObject(envelope);
        }
        catch (Exception ex)
        {
            await _runtime.Logger.LogAsync($"IDE Bridge: {commandName} FAIL - internal_error", ct, activatePane: true).ConfigureAwait(false);
            var errorObj = new { code = "internal_error", message = ex.Message, details = new { exception = ex.ToString() } };
            var failureData = await _runtime.FailureContextService.CaptureAsync(failureContext).ConfigureAwait(false);
            var envelope = BuildEnvelope(commandName, requestId, false, ex.Message, failureData, new JArray(), errorObj, startedAt);
            return JsonConvert.SerializeObject(envelope);
        }
    }

    private static CommandEnvelope BuildEnvelope(
        string command,
        string? requestId,
        bool success,
        string summary,
        JToken data,
        JArray warnings,
        object? error,
        DateTimeOffset startedAt)
    {
        return new CommandEnvelope
        {
            SchemaVersion = JsonSchemaVersioning.CurrentSchemaVersion,
            Command = command,
            RequestId = requestId,
            Success = success,
            StartedAtUtc = startedAt.UtcDateTime.ToString("O"),
            FinishedAtUtc = DateTimeOffset.UtcNow.UtcDateTime.ToString("O"),
            Summary = summary,
            Warnings = warnings,
            Error = error,
            Data = data,
        };
    }

    public void Dispose()
    {
        _cts.Cancel();
        try
        {
            if (File.Exists(_discoveryFile))
                File.Delete(_discoveryFile);
        }
        catch { }
    }
}
