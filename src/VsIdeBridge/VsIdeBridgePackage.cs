using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using VsIdeBridge.Services;

namespace VsIdeBridge;

[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
[InstalledProductRegistration("VS IDE Bridge", "Scriptable IDE control commands for Visual Studio.", "0.1.0")]
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
        _runtime = await IdeBridgeRuntime.CreateAsync(this).ConfigureAwait(false);
        await CommandRegistrar.InitializeAsync(this, _runtime).ConfigureAwait(false);

        // Start named pipe server (best-effort; failure does not break DTE commands)
        try
        {
            _pipeServer = new PipeServerService(this, _runtime);
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
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
            _pipeServer?.Dispose();
        base.Dispose(disposing);
    }
}
