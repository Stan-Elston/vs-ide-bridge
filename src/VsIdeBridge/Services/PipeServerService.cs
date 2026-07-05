using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using VsIdeBridge.Commands;
using VsIdeBridge.Infrastructure;
using VsIdeBridge.Shared;

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
    private readonly PipeServerDiscoveryCoordinator _discovery;
    private readonly CancellationTokenSource _cts = new();
    private readonly PipeServerMutableState _state = new();
    private readonly SemaphoreSlim _commandQueue = new(1, 1);
    private Task? _listenTask;
    private DTE2? _dte; // cached after first use; stable for the lifetime of the VS instance
    // Set by the timeout path while the command queue is held; consumed by the queue
    // release hand-off so the queue stays closed until the abandoned command finishes.
    private Task? _abandonedCommand;

    public PipeServerService(VsIdeBridgePackage package, IdeBridgeRuntime runtime)
    {
        _package = package;
        _runtime = runtime;
        _discovery = new PipeServerDiscoveryCoordinator(package, runtime, _cts);
    }

    public void Start()
    {
        _discovery.Start();
        _listenTask = Task.Factory.StartNew(
            () => ListenLoopAsync(_cts.Token),
            _cts.Token,
            TaskCreationOptions.LongRunning,
            TaskScheduler.Default).Unwrap();
    }

    public void UpdateDiscovery(string? solutionPath)
    {
        _discovery.UpdateDiscovery(solutionPath);
    }

    private static JArray BuildBatchSteps(PipeRequest request)
    {
        JArray steps = [];
        if (request.Batch == null)
        {
            return steps;
        }

        foreach (var batchRequest in request.Batch)
        {
            steps.Add(new JObject
            {
                ["id"] = (JToken?)batchRequest.Id ?? JValue.CreateNull(),
                ["command"] = batchRequest.Command ?? string.Empty,
                ["args"] = batchRequest.Args?.DeepClone() ?? JValue.CreateNull(),
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
                    PipeOptions.Asynchronous,
                     PipeServerConstants.PipeStreamBufferSize,
                     PipeServerConstants.PipeStreamBufferSize,
                    PipeServerSupport.CreatePipeSecurity());
            }
            catch (Exception ex) when (ex is not null) // pipe creation can throw Win32Exception, UnauthorizedAccessException, or IOException
            {
                ActivityLog.LogWarning(nameof(PipeServerService), $"Failed to create pipe server instance: {ex.Message}");
                try
                {
                    await Task.Delay(1000, ct).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    break;
                }

                continue;
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
            catch (Exception ex) when (ex is not null) // pipe accept boundary
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
        // net472 ReadLineAsync has no cancellation overload; disposing the pipe on shutdown
        // is the only way to unblock a pending read so this connection task can exit.
        using (ct.Register(() => SafeDisposePipe(pipe)))
        {
            try
            {
                await ServeConnectionRequestsAsync(pipe, ct).ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is not null) // pipe connection boundary
            {
                ActivityLog.LogError(nameof(PipeServerService), $"Pipe connection error: {ex.Message}");
            }
        }
    }

    private async Task ServeConnectionRequestsAsync(NamedPipeServerStream pipe, CancellationToken ct)
    {
        using StreamReader reader = new(pipe, Encoding.UTF8, detectEncodingFromByteOrderMarks: false, bufferSize: PipeServerConstants.PipeStreamBufferSize, leaveOpen: true);
        using StreamWriter writer = new(pipe, new UTF8Encoding(false), bufferSize: PipeServerConstants.PipeStreamBufferSize, leaveOpen: true)
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

            string responseLine;
            try
            {
                responseLine = await ExecuteRequestAsync(line, ct).ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is not null) // request execution boundary
            {
                // ExecuteRequestAsync should handle all exceptions internally, but if it
                // escapes (e.g. OperationCanceledException from WaitAsync during VS shutdown),
                // write an error response so the client never receives a raw EOF.
                responseLine = BuildInterruptedResponseLine(ex);
            }
            try
            {
                await writer.WriteLineAsync(responseLine).ConfigureAwait(false);
            }
            catch
            {
                break; // pipe broke during write — client already disconnected
            }
        }
    }

    // Serialize a real envelope: exception messages can contain quotes and newlines, and the
    // client matches the envelope's Summary/Error fields to detect interrupted operations.
    private static string BuildInterruptedResponseLine(Exception ex) =>
        JsonConvert.SerializeObject(new CommandEnvelope
        {
            SchemaVersion = JsonSchemaVersioning.CurrentSchemaVersion,
            Command = string.Empty,
            RequestId = null,
            Success = false,
            StartedAtUtc = DateTimeOffset.UtcNow.UtcDateTime.ToString("O"),
            FinishedAtUtc = DateTimeOffset.UtcNow.UtcDateTime.ToString("O"),
            Summary = ex.Message,
            Warnings = [],
            Error = new { code = "internal_error", message = $"Bridge server interrupted: {ex.Message}" },
            Data = new JObject(),
        });

    private async Task<string> ExecuteRequestAsync(string requestJson, CancellationToken ct)
    {
        PipeRequest request;
        try
        {
            request = JsonConvert.DeserializeObject<PipeRequest>(requestJson)
                ?? throw new CommandErrorException("invalid_request", "Could not parse request JSON.");
        }
        catch (JsonException ex)
        {
            CommandEnvelope envelope = new()
            {
                SchemaVersion = JsonSchemaVersioning.CurrentSchemaVersion,
                Command = string.Empty,
                RequestId = null,
                Success = false,
                StartedAtUtc = DateTimeOffset.UtcNow.UtcDateTime.ToString("O"),
                FinishedAtUtc = DateTimeOffset.UtcNow.UtcDateTime.ToString("O"),
                Summary = "Could not parse request JSON.",
                Warnings = [],
                Error = new { code = "invalid_request", message = ex.Message },
                Data = new JObject(),
            };
            return JsonConvert.SerializeObject(envelope);
        }

        bool hasBatch = request.Batch is { Count: > 0 };
        string commandName = hasBatch
            ? (!string.IsNullOrWhiteSpace(request.Command) ? request.Command : "Tools.IdeBatchCommands")
            : (request.Command ?? string.Empty);
        int timeoutMilliseconds = ResolveTimeoutMilliseconds(commandName, request.Args, hasBatch);
        bool isDiagnosticsCommand = IsDiagnosticsCommand(commandName);
        DateTimeOffset enqueuedAt = DateTimeOffset.UtcNow;
        int queueCount = Interlocked.Increment(ref _state.QueuedCommandCount);
        int positionAtEnqueue = Math.Max(0, queueCount - 1);
        bool acquired = false;
        try
        {
            // Bound the queue wait by the command's own timeout: a previous command that
            // timed out may still be running inside VS and holding the queue (see the
            // zombie hand-off in ReleaseCommandQueueAfterZombie).
            acquired = await _commandQueue.WaitAsync(timeoutMilliseconds, ct).ConfigureAwait(false);
            if (!acquired)
            {
                return PipeServerSupport.SerializeFailureEnvelope(
                    commandName,
                    request.Id,
                    "The bridge is still executing a previous command (it likely timed out but is still running inside VS). Wait for it to finish, then retry.",
                    new JObject(),
                    new JObject
                    {
                        ["code"] = "bridge_busy",
                        ["message"] = $"Queue wait exceeded {timeoutMilliseconds} ms while a previous command was still executing.",
                    },
                    enqueuedAt,
                    DateTimeOffset.UtcNow,
                    positionAtEnqueue,
                    (DateTimeOffset.UtcNow - enqueuedAt).TotalMilliseconds);
            }

            DateTimeOffset startedAt = DateTimeOffset.UtcNow;
            double queueWaitMs = (startedAt - enqueuedAt).TotalMilliseconds;
            return await ExecuteRequestCoreAsync(
                request,
                commandName,
                hasBatch,
                timeoutMilliseconds,
                isDiagnosticsCommand,
                ct,
                enqueuedAt,
                startedAt,
                positionAtEnqueue,
                queueWaitMs).ConfigureAwait(false);
        }
        finally
        {
            Interlocked.Decrement(ref _state.QueuedCommandCount);
            if (acquired)
            {
                ReleaseCommandQueueAfterZombie();
            }
        }
    }

    private void ReleaseCommandQueueAfterZombie()
    {
        Task? zombie = _abandonedCommand;
        _abandonedCommand = null;
        if (zombie is null || zombie.IsCompleted)
        {
            ReleaseCommandQueueSafe();
            return;
        }

        // A timed-out command is still executing inside VS. Keep the queue closed until it
        // finishes so the next command cannot mutate the IDE concurrently with it.
        _ = zombie.ContinueWith(
            _ => ReleaseCommandQueueSafe(),
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously,
            TaskScheduler.Default);
    }

    private void ReleaseCommandQueueSafe()
    {
        try
        {
            _commandQueue.Release();
        }
        catch (ObjectDisposedException ex)
        {
            // VS is shutting down; the queue is already gone. Expected during dispose races.
            ActivityLog.LogWarning(nameof(PipeServerService), $"Command queue released after dispose: {ex.Message}");
        }
    }

    private static void SafeDisposePipe(NamedPipeServerStream pipe)
    {
        try
        {
            pipe.Dispose();
        }
        catch (Exception ex) when (ex is not null) // dispose race during shutdown
        {
            System.Diagnostics.Debug.WriteLine(ex);
        }
    }

    private async Task<string> ExecuteRequestCoreAsync(
        PipeRequest request,
        string commandName,
        bool hasBatch,
        int timeoutMilliseconds,
        bool isDiagnosticsCommand,
        CancellationToken serverCancellationToken,
        DateTimeOffset enqueuedAtUtc,
        DateTimeOffset startedAtUtc,
        int queuePositionAtEnqueue,
        double queueWaitMs)
    {
        Stopwatch commandStopwatch = Stopwatch.StartNew();
        string? requestId = request.Id;
        bool completionRecorded = false;
        IdeCommandContext? failureContext = null;
        using CancellationTokenSource commandCts = CancellationTokenSource.CreateLinkedTokenSource(serverCancellationToken);
        commandCts.CancelAfter(timeoutMilliseconds);
        CancellationToken commandToken = commandCts.Token;
        _runtime.BridgeWatchdogService.RecordCommandStarted(commandName, requestId);

        try
        {
            await _runtime.Logger.LogAsync($"IDE Bridge Trace: {commandName} start (timeout={timeoutMilliseconds}ms)", commandToken).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is not null) // best-effort trace logging
        {
            System.Diagnostics.Debug.WriteLine(ex);
        }

        try
        {
            JoinableTask<CommandExecutionResult> executionTask = _package.JoinableTaskFactory.RunAsync(() => ExecuteCommandAsync(
                request,
                commandName,
                hasBatch,
                commandToken,
                ctx => failureContext = ctx));
            CommandExecutionResult commandResult = await AwaitCommandExecutionAsync(
                executionTask,
                timeoutMilliseconds,
                commandCts,
                commandToken,
                serverCancellationToken).ConfigureAwait(false);
            return await CompleteSuccessfulRequestAsync(
                commandName,
                requestId,
                commandResult,
                commandStopwatch,
                commandToken,
                enqueuedAtUtc,
                startedAtUtc,
                queuePositionAtEnqueue,
                queueWaitMs,
                onCompleted: () => completionRecorded = true).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (commandCts.IsCancellationRequested && !serverCancellationToken.IsCancellationRequested)
        {
            return await HandleTimedOutRequestAsync(
                commandName,
                requestId,
                isDiagnosticsCommand,
                timeoutMilliseconds,
                commandStopwatch,
                failureContext,
                enqueuedAtUtc,
                startedAtUtc,
                queuePositionAtEnqueue,
                queueWaitMs,
                onCompleted: () => completionRecorded = true).ConfigureAwait(false);
        }
        catch (CommandErrorException ex)
        {
            return await HandleCommandFailureAsync(
                commandName,
                requestId,
                ex,
                commandStopwatch,
                failureContext,
                enqueuedAtUtc,
                startedAtUtc,
                queuePositionAtEnqueue,
                queueWaitMs,
                onCompleted: () => completionRecorded = true).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is not null) // top-level request envelope boundary
        {
            return await HandleInternalRequestFailureAsync(
                commandName,
                requestId,
                ex,
                commandStopwatch,
                failureContext,
                enqueuedAtUtc,
                startedAtUtc,
                queuePositionAtEnqueue,
                queueWaitMs,
                onCompleted: () => completionRecorded = true).ConfigureAwait(false);
        }
        finally
        {
            if (!completionRecorded)
            {
                _runtime.BridgeWatchdogService.RecordCommandCompleted(commandName, requestId, success: false, durationMs: commandStopwatch.Elapsed.TotalMilliseconds, "internal_error");
            }
        }
    }

    private async Task<string> CompleteSuccessfulRequestAsync(
        string commandName,
        string? requestId,
        CommandExecutionResult commandResult,
        Stopwatch commandStopwatch,
        CancellationToken commandToken,
        DateTimeOffset enqueuedAtUtc,
        DateTimeOffset startedAtUtc,
        int queuePositionAtEnqueue,
        double queueWaitMs,
        Action onCompleted)
    {
        await _runtime.Logger.LogAsync(
            $"IDE Bridge: {commandName} OK - {commandResult.Summary}",
            commandToken,
            activatePane: ShouldRevealActivity(commandName)).ConfigureAwait(false);
        await _runtime.Logger.LogAsync($"IDE Bridge Trace: {commandName} end (duration={commandStopwatch.ElapsedMilliseconds}ms)", commandToken).ConfigureAwait(false);
        _runtime.BridgeWatchdogService.RecordCommandCompleted(commandName, requestId, success: true, durationMs: commandStopwatch.Elapsed.TotalMilliseconds, errorCode: null);
        onCompleted();
        return PipeServerSupport.SerializeSuccessEnvelope(
            commandName,
            requestId,
            commandResult,
            enqueuedAtUtc,
            startedAtUtc,
            queuePositionAtEnqueue,
            queueWaitMs);
    }

    private async Task<string> HandleTimedOutRequestAsync(
        string commandName,
        string? requestId,
        bool isDiagnosticsCommand,
        int timeoutMilliseconds,
        Stopwatch commandStopwatch,
        IdeCommandContext? failureContext,
        DateTimeOffset enqueuedAtUtc,
        DateTimeOffset startedAtUtc,
        int queuePositionAtEnqueue,
        double queueWaitMs,
        Action onCompleted)
    {
        string errorCode = isDiagnosticsCommand ? "ide_blocked" : "timeout";
        string summary = isDiagnosticsCommand
            ? "IDE appears blocked (modal dialog or busy state) while processing the command."
            : $"Command timed out after {timeoutMilliseconds} ms. It may still be running inside VS; " +
              "the bridge holds new commands until it completes.";
        if (!isDiagnosticsCommand && IsDocumentMutatingCommand(commandName))
        {
            summary += " A timeout does not undo the edit - it may still land after this response. When the bridge " +
                "responds again, verify the document content and its saved state (list-tabs flags unsaved tabs) and " +
                "run save-document if the tab is dirty.";
        }
        CommandTimeoutDetails details = new(
            TimeoutMs: timeoutMilliseconds,
            DurationMs: commandStopwatch.ElapsedMilliseconds,
            Reason: isDiagnosticsCommand ? "ui_not_responsive" : "command_timeout");
        CommandTimeoutError errorObj = new(Code: errorCode, Message: summary, Details: details);
        JObject failureData = await _runtime.FailureContextService.CaptureAsync(failureContext).ConfigureAwait(false);
        try
        {
            await _runtime.Logger.LogAsync($"IDE Bridge Trace: {commandName} timeout (duration={commandStopwatch.ElapsedMilliseconds}ms)", CancellationToken.None, activatePane: true).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is not null) // best-effort timeout logging
        {
            System.Diagnostics.Debug.WriteLine(ex);
        }

        _runtime.BridgeWatchdogService.RecordCommandCompleted(commandName, requestId, success: false, durationMs: commandStopwatch.Elapsed.TotalMilliseconds, errorCode);
        onCompleted();
        return PipeServerSupport.SerializeFailureEnvelope(
            commandName,
            requestId,
            summary,
            failureData,
            errorObj,
            enqueuedAtUtc,
            startedAtUtc,
            queuePositionAtEnqueue,
            queueWaitMs);
    }

    // Commands that mutate editor documents: a timeout does not undo the edit, so the
    // caller must be told to re-check content and saved state (field report issue 17).
    private static bool IsDocumentMutatingCommand(string commandName)
        => commandName is "apply-diff" or "apply-patch" or "write-file"
            or "Tools.IdeApplyUnifiedDiff" or "Tools.IdeWriteFile";

    private async Task<string> HandleCommandFailureAsync(
        string commandName,
        string? requestId,
        CommandErrorException ex,
        Stopwatch commandStopwatch,
        IdeCommandContext? failureContext,
        DateTimeOffset enqueuedAtUtc,
        DateTimeOffset startedAtUtc,
        int queuePositionAtEnqueue,
        double queueWaitMs,
        Action onCompleted)
    {
        await _runtime.Logger.LogAsync($"IDE Bridge: {commandName} FAIL - {ex.Code}", CancellationToken.None, activatePane: true).ConfigureAwait(false);
        JObject errorObj = new()
        {
            ["code"] = ex.Code,
            ["message"] = ex.Message,
            ["details"] = ex.Details is null ? null : JToken.FromObject(ex.Details),
        };
        JToken failureData = PipeServerSupport.IsCompactCommandError(ex.Code)
            ? PipeServerSupport.BuildCompactCommandErrorData(commandName, ex)
            : await _runtime.FailureContextService.CaptureAsync(failureContext).ConfigureAwait(false);
        await _runtime.Logger.LogAsync($"IDE Bridge Trace: {commandName} end (duration={commandStopwatch.ElapsedMilliseconds}ms, failed={ex.Code})", CancellationToken.None).ConfigureAwait(false);
        _runtime.BridgeWatchdogService.RecordCommandCompleted(commandName, requestId, success: false, durationMs: commandStopwatch.Elapsed.TotalMilliseconds, ex.Code);
        onCompleted();
        return PipeServerSupport.SerializeFailureEnvelope(
            commandName,
            requestId,
            ex.Message,
            failureData,
            errorObj,
            enqueuedAtUtc,
            startedAtUtc,
            queuePositionAtEnqueue,
            queueWaitMs);
    }

    private async Task<string> HandleInternalRequestFailureAsync(
        string commandName,
        string? requestId,
        Exception ex,
        Stopwatch commandStopwatch,
        IdeCommandContext? failureContext,
        DateTimeOffset enqueuedAtUtc,
        DateTimeOffset startedAtUtc,
        int queuePositionAtEnqueue,
        double queueWaitMs,
        Action onCompleted)
    {
        await _runtime.Logger.LogAsync($"IDE Bridge: {commandName} FAIL - internal_error", CancellationToken.None, activatePane: true).ConfigureAwait(false);
        JObject errorObj = new()
        {
            ["code"] = "internal_error",
            ["message"] = ex.Message,
            ["details"] = new JObject
            {
                ["exception"] = ex.ToString(),
            },
        };
        JObject failureData = await _runtime.FailureContextService.CaptureAsync(failureContext).ConfigureAwait(false);
        await _runtime.Logger.LogAsync($"IDE Bridge Trace: {commandName} end (duration={commandStopwatch.ElapsedMilliseconds}ms, failed=internal_error)", CancellationToken.None).ConfigureAwait(false);
        _runtime.BridgeWatchdogService.RecordCommandCompleted(commandName, requestId, success: false, durationMs: commandStopwatch.Elapsed.TotalMilliseconds, "internal_error");
        onCompleted();
        return PipeServerSupport.SerializeFailureEnvelope(
            commandName,
            requestId,
            ex.Message,
            failureData,
            errorObj,
            enqueuedAtUtc,
            startedAtUtc,
            queuePositionAtEnqueue,
            queueWaitMs);
    }

    private async Task<CommandExecutionResult> ExecuteCommandAsync(PipeRequest request, string commandName, bool hasBatch, CancellationToken commandToken, Action<IdeCommandContext> captureFailureContext)
    {
        DTE2? dte = await GetDteAsync(commandToken).ConfigureAwait(false);
        Assumes.Present(dte);

        string? solutionPath = await _discovery.CaptureSolutionPathAsync(dte!, commandToken).ConfigureAwait(false);
        _discovery.UpdateDiscovery(solutionPath);

        IdeCommandContext ctx = new(_package, dte!, _runtime.Logger, _runtime, commandToken);
        captureFailureContext(ctx);
        await _runtime.Logger.LogAsync($"IDE Bridge: {commandName} requested", commandToken).ConfigureAwait(false);

        if (hasBatch)
        {
            JArray steps = BuildBatchSteps(request);
            CommandExecutionResult batchResult = await IdeCoreCommands.ExecuteBatchAsync(ctx, steps, request.StopOnError ?? false).ConfigureAwait(false);
            await _discovery.RefreshAfterCommandAsync(dte!, solutionPath, commandToken).ConfigureAwait(false);
            return batchResult;
        }

        CommandArguments args = CommandArgumentParser.Parse(request.Args);
        if (IsWriteFileCommand(commandName))
        {
            CommandExecutionResult writeResult = await PatchCommands.ExecuteWriteFileAsync(ctx, args).ConfigureAwait(false);
            await _discovery.RefreshAfterCommandAsync(dte!, solutionPath, commandToken).ConfigureAwait(false);
            return writeResult;
        }

        if (!_runtime.TryGetCommand(commandName, out IdeCommandBase cmd))
        {
            throw new CommandErrorException("command_not_found", $"Unknown command: '{commandName}'. Call tool_help with no arguments to see all available tool names and their parameters, then retry with a valid command name.");
        }

        CommandExecutionResult result = await cmd.ExecuteDirectAsync(ctx, args).ConfigureAwait(false);
        await _discovery.RefreshAfterCommandAsync(dte!, solutionPath, commandToken).ConfigureAwait(false);
        return result;
    }

    private static bool IsWriteFileCommand(string commandName)
        => string.Equals(commandName, "write-file", StringComparison.OrdinalIgnoreCase)
            || string.Equals(commandName, "Tools.IdeWriteFile", StringComparison.OrdinalIgnoreCase);

    private async Task<CommandExecutionResult> AwaitCommandExecutionAsync(
        JoinableTask<CommandExecutionResult> executionTask,
        int timeoutMilliseconds,
        CancellationTokenSource commandCts,
        CancellationToken commandToken,
        CancellationToken serverCancellationToken)
    {
        Task<CommandExecutionResult> joinedExecutionTask = executionTask.JoinAsync(serverCancellationToken);
        Task completed = await Task.WhenAny(
            joinedExecutionTask,
            Task.Delay(timeoutMilliseconds, serverCancellationToken)).ConfigureAwait(false);
        if (ReferenceEquals(completed, joinedExecutionTask))
        {
            return await joinedExecutionTask.ConfigureAwait(false);
        }

        if (serverCancellationToken.IsCancellationRequested)
        {
            throw new OperationCanceledException(serverCancellationToken);
        }

        commandCts.Cancel();
        // The command may ignore cancellation and keep running inside VS. Observe its
        // eventual fault and hand its completion to the queue release path so no other
        // command executes concurrently with it.
        _abandonedCommand = joinedExecutionTask.ContinueWith(
            task => _ = task.Exception,
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously,
            TaskScheduler.Default);
        throw new OperationCanceledException(commandToken);
    }

    private async Task<DTE2?> GetDteAsync(CancellationToken cancellationToken)
    {
        if (_dte is not null) return _dte;
        await _package.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
        _dte = await _package.GetServiceAsync(typeof(SDTE)).ConfigureAwait(true) as DTE2;
        return _dte;
    }

    private static bool IsDiagnosticsCommand(string commandName)
    {
        return string.Equals(commandName, "Tools.IdeWaitForReady", StringComparison.OrdinalIgnoreCase)
            || string.Equals(commandName, "Tools.IdeGetErrorList", StringComparison.OrdinalIgnoreCase)
            || string.Equals(commandName, "Tools.IdeGetWarnings", StringComparison.OrdinalIgnoreCase)
            || string.Equals(commandName, "Tools.IdeGetMessages", StringComparison.OrdinalIgnoreCase)
            || string.Equals(commandName, "Tools.IdeBuildAndCaptureErrors", StringComparison.OrdinalIgnoreCase)
            || string.Equals(commandName, "Tools.IdeDiagnosticsSnapshot", StringComparison.OrdinalIgnoreCase)
            || string.Equals(commandName, "ready", StringComparison.OrdinalIgnoreCase)
            || string.Equals(commandName, "errors", StringComparison.OrdinalIgnoreCase)
            || string.Equals(commandName, "warnings", StringComparison.OrdinalIgnoreCase)
            || string.Equals(commandName, "messages", StringComparison.OrdinalIgnoreCase)
            || string.Equals(commandName, "build-errors", StringComparison.OrdinalIgnoreCase);
    }

    private static int ResolveTimeoutMilliseconds(string commandName, JToken? rawArgs, bool isBatch)
    {
        if (isBatch)
        {
             return PipeServerConstants.DefaultCommandTimeoutMilliseconds;
        }

        int defaultTimeout = IsDiagnosticsCommand(commandName)
             ? PipeServerConstants.DefaultDiagnosticsTimeoutMilliseconds
             : PipeServerConstants.DefaultCommandTimeoutMilliseconds;

        if (rawArgs is null || rawArgs.Type == JTokenType.Null)
        {
            return defaultTimeout;
        }

        try
        {
            CommandArguments args = CommandArgumentParser.Parse(rawArgs);
            int requested = args.GetInt32("timeout-ms", defaultTimeout);
            return requested > 0 ? requested : defaultTimeout;
        }
        catch (CommandErrorException)
        {
            return defaultTimeout;
        }
    }

    public void Dispose()
    {
        _cts.Cancel();
        _cts.Dispose();
        _commandQueue.Dispose();
        _discovery.Dispose();
    }
}

file static class PipeServerConstants
{
    internal const int PipeStreamBufferSize = 4096;
    internal const int DefaultCommandTimeoutMilliseconds = 120_000;
    internal const int DefaultDiagnosticsTimeoutMilliseconds = 120_000;
}

