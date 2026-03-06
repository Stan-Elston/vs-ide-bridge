using System.IO.Pipes;
using System.ServiceProcess;
using System.Text;

namespace VsIdeBridgeService;

internal static class Program
{
    public static void Main(string[] args)
    {
        if (!OperatingSystem.IsWindows())
        {
            Console.Error.WriteLine("VsIdeBridgeService is Windows-only.");
            Environment.ExitCode = 1;
            return;
        }

        ServiceBase.Run(new BridgeService(args));
    }
}

internal sealed class BridgeService : ServiceBase
{
    private const string ServiceControlPipeName = "VsIdeBridgeServiceControl";

    private readonly TimeSpan _idleSoftTimeout;
    private readonly TimeSpan _idleHardTimeout;
    private readonly object _stateGate = new();
    private readonly string _logPath;

    private CancellationTokenSource? _stopCts;
    private Task? _acceptLoop;
    private Task? _idleLoop;

    private DateTime _lastActivityUtc;
    private int _connectedClients;
    private int _inFlightCommands;
    private bool _draining;

    public BridgeService(string[] args)
    {
        ServiceName = "VsIdeBridgeService";
        CanStop = true;
        CanPauseAndContinue = false;
        AutoLog = false;

        _idleSoftTimeout = TimeSpan.FromSeconds(GetIntArg(args, "idle-soft-seconds", 900));
        _idleHardTimeout = TimeSpan.FromSeconds(GetIntArg(args, "idle-hard-seconds", 1200));

        var programData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
        var logDir = Path.Combine(programData, "VsIdeBridge");
        Directory.CreateDirectory(logDir);
        _logPath = Path.Combine(logDir, "service.log");
    }

    protected override void OnStart(string[] args)
    {
        _stopCts = new CancellationTokenSource();
        lock (_stateGate)
        {
            _lastActivityUtc = DateTime.UtcNow;
            _connectedClients = 0;
            _inFlightCommands = 0;
            _draining = false;
        }

        Log($"service started; idle_soft={_idleSoftTimeout.TotalSeconds}s idle_hard={_idleHardTimeout.TotalSeconds}s");
        _acceptLoop = Task.Run(() => AcceptLoopAsync(_stopCts.Token));
        _idleLoop = Task.Run(() => IdleLoopAsync(_stopCts.Token));
    }

    protected override void OnStop()
    {
        Log("service stopping");
        _stopCts?.Cancel();

        try { _acceptLoop?.Wait(TimeSpan.FromSeconds(5)); } catch { }
        try { _idleLoop?.Wait(TimeSpan.FromSeconds(5)); } catch { }

        Log("service stopped");
        _stopCts?.Dispose();
        _stopCts = null;
    }

    private async Task AcceptLoopAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                using var server = new NamedPipeServerStream(
                    ServiceControlPipeName,
                    PipeDirection.In,
                    NamedPipeServerStream.MaxAllowedServerInstances,
                    PipeTransmissionMode.Byte,
                    PipeOptions.Asynchronous);

                await server.WaitForConnectionAsync(cancellationToken).ConfigureAwait(false);
                using var reader = new StreamReader(server, Encoding.UTF8, detectEncodingFromByteOrderMarks: false, leaveOpen: true);

                while (server.IsConnected && !cancellationToken.IsCancellationRequested)
                {
                    var line = await reader.ReadLineAsync(cancellationToken).ConfigureAwait(false);
                    if (line is null)
                    {
                        break;
                    }

                    HandleEvent(line.Trim());
                }
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                Log($"control pipe error: {ex.Message}");
                await Task.Delay(500, cancellationToken).ConfigureAwait(false);
            }
        }
    }

    private async Task IdleLoopAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                await Task.Delay(TimeSpan.FromSeconds(5), cancellationToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                break;
            }

            bool shouldLogDraining = false;
            bool shouldStop = false;
            TimeSpan idle;
            int clients;
            int inFlight;

            lock (_stateGate)
            {
                idle = DateTime.UtcNow - _lastActivityUtc;
                clients = _connectedClients;
                inFlight = _inFlightCommands;

                if (inFlight > 0 || clients > 0)
                {
                    continue;
                }

                if (!_draining && idle >= _idleSoftTimeout)
                {
                    _draining = true;
                    shouldLogDraining = true;
                }

                if (_draining && idle >= _idleHardTimeout)
                {
                    shouldStop = true;
                }
            }

            if (shouldLogDraining)
            {
                Log($"service going idle: draining started after {idle.TotalSeconds:F0}s inactivity");
            }

            if (shouldStop)
            {
                Log($"service going idle: stopping after {idle.TotalSeconds:F0}s inactivity (clients={clients}, inFlight={inFlight})");
                Stop();
                return;
            }
        }
    }

    private void HandleEvent(string evt)
    {
        if (string.IsNullOrWhiteSpace(evt))
        {
            return;
        }

        lock (_stateGate)
        {
            switch (evt.ToUpperInvariant())
            {
                case "MCP_REQUEST":
                    _lastActivityUtc = DateTime.UtcNow;
                    _draining = false;
                    break;
                case "COMMAND_START":
                    _inFlightCommands++;
                    _lastActivityUtc = DateTime.UtcNow;
                    _draining = false;
                    break;
                case "COMMAND_END":
                    if (_inFlightCommands > 0)
                    {
                        _inFlightCommands--;
                    }

                    _lastActivityUtc = DateTime.UtcNow;
                    _draining = false;
                    break;
                case "CLIENT_CONNECTED":
                    _connectedClients++;
                    _lastActivityUtc = DateTime.UtcNow;
                    _draining = false;
                    break;
                case "CLIENT_DISCONNECTED":
                    if (_connectedClients > 0)
                    {
                        _connectedClients--;
                    }

                    _lastActivityUtc = DateTime.UtcNow;
                    break;
                case "PING":
                    _lastActivityUtc = DateTime.UtcNow;
                    break;
                default:
                    Log($"unknown control event '{evt}'");
                    break;
            }
        }
    }

    private void Log(string message)
    {
        try
        {
            File.AppendAllText(_logPath, $"{DateTime.Now:O} {message}{Environment.NewLine}");
        }
        catch
        {
            // best effort logging
        }
    }

    private static int GetIntArg(string[] args, string name, int defaultValue)
    {
        for (var i = 0; i < args.Length; i++)
        {
            if (!args[i].Equals($"--{name}", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (i + 1 >= args.Length)
            {
                return defaultValue;
            }

            return int.TryParse(args[i + 1], out var parsed) && parsed > 0 ? parsed : defaultValue;
        }

        return defaultValue;
    }
}
