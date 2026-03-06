using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using VsIdeBridge.Services;

namespace VsIdeBridge;

[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
[InstalledProductRegistration("VS IDE Bridge", "Scriptable IDE control commands for Visual Studio.", "2.0.0")]
[ProvideMenuResource("Menus.ctmenu", 1)]
[ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
[ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
[Guid(PackageGuidString)]
public sealed class VsIdeBridgePackage : AsyncPackage
{
    public const string PackageGuidString = "D8F750B1-5FB7-4A52-8D75-ED5A7F576088";

    private IdeBridgeRuntime? _runtime;
    private PipeServerService? _pipeServer;

    internal IdeBridgeRuntime Runtime => _runtime ?? throw new InvalidOperationException("Runtime is not initialized.");

    protected override async Task InitializeAsync(
        CancellationToken cancellationToken,
        IProgress<ServiceProgressData> progress)
    {
        IdeBridgeRuntime runtime;
        try
        {
            runtime = await IdeBridgeRuntime.CreateAsync(this).ConfigureAwait(false);
            _runtime = runtime;
        }
        catch (Exception ex)
        {
            ActivityLog.LogError(nameof(VsIdeBridgePackage), $"Runtime initialization failed: {ex}");
            return;
        }

        try
        {
            await CommandRegistrar.InitializeAsync(this, runtime).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            ActivityLog.LogError(nameof(VsIdeBridgePackage), $"Command registration failed: {ex}");
            return;
        }

        try
        {
            runtime.BridgeWatchdogService.Start();
        }
        catch (Exception ex)
        {
            ActivityLog.LogWarning(nameof(VsIdeBridgePackage), $"Bridge watchdog failed to start: {ex.Message}");
        }

        // Start named pipe server (best-effort; failure does not break DTE commands)
        try
        {
            _pipeServer = new PipeServerService(this, runtime);
            _pipeServer.Start();
        }
        catch (Exception ex)
        {
            ActivityLog.LogWarning(nameof(VsIdeBridgePackage), $"Pipe server failed to start: {ex.Message}");
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _pipeServer?.Dispose();
            _runtime?.BridgeWatchdogService.Dispose();
        }
        base.Dispose(disposing);
    }
}
