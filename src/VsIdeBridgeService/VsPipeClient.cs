using System.IO.Pipes;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;

namespace VsIdeBridgeService;

// Connects to one VS bridge named pipe, sends a request, and reads the response.
// Acquires a per-pipe file lock to prevent concurrent requests on the same pipe.
internal sealed class VsPipeClient : IAsyncDisposable
{
    private readonly NamedPipeClientStream _pipe;
    private readonly StreamReader _reader;
    private readonly StreamWriter _writer;
    private readonly int _timeoutMs;
    private readonly FileStream _gate;

    private static readonly string LockDirectory =
        Path.Combine(Path.GetTempPath(), "vs-ide-bridge", "locks");

    private VsPipeClient(NamedPipeClientStream pipe, StreamReader reader, StreamWriter writer, int timeoutMs, FileStream gate)
    {
        _pipe = pipe;
        _reader = reader;
        _writer = writer;
        _timeoutMs = timeoutMs;
        _gate = gate;
    }

    // Connecting to a live pipe is fast; waiting the full command timeout on a dead pipe
    // (VS restarted, new pid) only delays the stale-binding recovery in BridgeConnection.
    private const int MaxConnectTimeoutMs = 10_000;

    public static async Task<VsPipeClient> CreateAsync(string pipeName, int timeoutMs, int gateTimeoutMs)
    {
        int resolvedTimeoutMs = Math.Max(1_000, timeoutMs);
        FileStream? gate = await AcquireGateAsync(pipeName, Math.Max(250, gateTimeoutMs)).ConfigureAwait(false);

        NamedPipeClientStream? pipe = null;
        try
        {
            pipe = new(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
            int connectTimeoutMs = Math.Min(resolvedTimeoutMs, MaxConnectTimeoutMs);
            using CancellationTokenSource cts = new(connectTimeoutMs);
            try
            {
                await pipe.ConnectAsync(cts.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                throw new BridgeConnectTimeoutException(pipeName, connectTimeoutMs);
            }

            StreamReader reader = new(pipe, new UTF8Encoding(false), detectEncodingFromByteOrderMarks: false, leaveOpen: true);
            StreamWriter writer = new(pipe, new UTF8Encoding(false), leaveOpen: true)
            {
                AutoFlush = true,
                NewLine = "\n",
            };

            VsPipeClient client = new(pipe, reader, writer, resolvedTimeoutMs, gate);
            pipe = null; // ownership transferred to the client
            gate = null;
            return client;
        }
        finally
        {
            // Reached with non-null locals only when construction failed. The gate lock
            // file must not outlive a failed connect, or every following call fails with
            // BridgeBusyException until the GC finalizes the leaked stream.
            pipe?.Dispose();
            gate?.Dispose();
        }
    }

    public async Task<JsonObject> SendAsync(JsonObject payload)
    {
        using CancellationTokenSource cts = new(_timeoutMs);
        try
        {
            await _writer.WriteLineAsync(payload.ToJsonString().AsMemory(), cts.Token).ConfigureAwait(false);
            string? line = await _reader.ReadLineAsync(cts.Token).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(line))
                throw new BridgeException("The VS bridge pipe returned an empty response.");
            return JsonNode.Parse(line) as JsonObject
                ?? throw new BridgeException("The VS bridge pipe returned malformed JSON.");
        }
        catch (OperationCanceledException)
        {
            throw new TimeoutException(
                $"Timed out waiting for VS bridge response after {_timeoutMs} ms. Visual Studio may be blocked.");
        }
    }

    public ValueTask DisposeAsync()
    {
        _reader.Dispose();
        _writer.Dispose();
        _pipe.Dispose();
        _gate.Dispose();
        return ValueTask.CompletedTask;
    }

    private static async Task<FileStream> AcquireGateAsync(string pipeName, int gateTimeoutMs)
    {
        Directory.CreateDirectory(LockDirectory);
        string hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(pipeName)));
        string lockFile = Path.Combine(LockDirectory, $"{hash}.lock");
        DateTime deadline = DateTime.UtcNow.AddMilliseconds(gateTimeoutMs);

        while (DateTime.UtcNow <= deadline)
        {
            try
            {
                return new FileStream(lockFile, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                await Task.Delay(25).ConfigureAwait(false);
            }
        }

        throw new BridgeBusyException(pipeName, gateTimeoutMs);
    }
}

internal sealed class BridgeBusyException(string pipeName, int gateTimeoutMs)
    : BridgeException($"Bridge pipe '{pipeName}' is busy with another request. Try again soon or avoid overlapping bridge calls. Waited {gateTimeoutMs} ms for exclusive access.");

// Thrown when the pipe cannot be connected at all - the request never reached VS, and the
// cached instance is most likely stale (VS restarted or closed). BridgeConnection evicts the
// cached instance and retries once with fresh discovery on this exception.
internal sealed class BridgeConnectTimeoutException(string pipeName, int timeoutMs)
    : TimeoutException($"Timed out connecting to VS bridge pipe '{pipeName}' after {timeoutMs} ms. The pipe is likely stale (VS restarted or closed).");

// Thrown for logical bridge errors (not I/O or timeout) so callers can distinguish error types.
internal class BridgeException(string message) : Exception(message);
