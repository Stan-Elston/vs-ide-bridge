using System;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace VsIdeBridge.Infrastructure;

internal sealed class OutputPaneLogger
{
    private const string PaneName = "IDE Bridge";

    private readonly AsyncPackage _package;
    private readonly DTE2 _dte;

    public OutputPaneLogger(AsyncPackage package, DTE2 dte)
    {
        _package = package;
        _dte = dte;
    }

    public async Task LogAsync(string message, CancellationToken cancellationToken, bool activatePane = false)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

        var pane = GetOrCreatePane();
        pane.OutputString($"{message}{Environment.NewLine}");
        if (activatePane)
        {
            pane.Activate();
        }

        if (await _package.GetServiceAsync(typeof(SVsStatusbar)).ConfigureAwait(true) is IVsStatusbar statusbar)
        {
            statusbar.SetText(message);
        }
    }

    private EnvDTE.OutputWindowPane GetOrCreatePane()
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var panes = _dte.ToolWindows.OutputWindow.OutputWindowPanes;
        for (var i = 1; i <= panes.Count; i++)
        {
            var pane = panes.Item(i);
            if (string.Equals(pane.Name, PaneName, StringComparison.OrdinalIgnoreCase))
            {
                return pane;
            }
        }

        return panes.Add(PaneName);
    }
}
