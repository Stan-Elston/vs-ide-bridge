using System.Runtime.InteropServices;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed class IdeStateService
{
    private readonly BridgeInstanceService _bridgeInstanceService;

    public IdeStateService(BridgeInstanceService bridgeInstanceService)
    {
        _bridgeInstanceService = bridgeInstanceService;
    }

    public async Task<JObject> GetStateAsync(EnvDTE80.DTE2 dte)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var solutionPath = dte.Solution?.IsOpen == true ? PathNormalization.NormalizeFilePath(dte.Solution.FullName) : string.Empty;
        var data = new JObject
        {
            ["solutionPath"] = solutionPath,
            ["solutionName"] = dte.Solution?.IsOpen == true ? Path.GetFileName(dte.Solution.FullName) : string.Empty,
            ["solutionDirectory"] = dte.Solution?.IsOpen == true ? Path.GetDirectoryName(dte.Solution.FullName) ?? string.Empty : string.Empty,
            ["debugMode"] = dte.Debugger.CurrentMode.ToString(),
            ["activeWindow"] = dte.ActiveWindow?.Caption ?? string.Empty,
            ["activeWindowKind"] = dte.ActiveWindow?.Kind ?? string.Empty,
            ["activeDocument"] = string.IsNullOrWhiteSpace(dte.ActiveDocument?.FullName) ? string.Empty : PathNormalization.NormalizeFilePath(dte.ActiveDocument.FullName),
            ["openDocuments"] = new JArray(dte.Documents.Cast<Document>()
                .Where(document => !string.IsNullOrWhiteSpace(document.FullName))
                .Select(document => PathNormalization.NormalizeFilePath(document.FullName))),
            ["startupProjects"] = GetStartupProjects(dte),
            ["bridge"] = _bridgeInstanceService.CreateStateData(solutionPath),
        };

        if (TryGetActiveTextSelection(dte.ActiveDocument, out var selection))
        {
            data["caretLine"] = selection.ActivePoint.Line;
            data["caretColumn"] = selection.ActivePoint.DisplayColumn;
            data["selectionStartLine"] = selection.TopPoint.Line;
            data["selectionStartColumn"] = selection.TopPoint.DisplayColumn;
            data["selectionEndLine"] = selection.BottomPoint.Line;
            data["selectionEndColumn"] = selection.BottomPoint.DisplayColumn;
        }

        var activeConfig = dte.Solution?.SolutionBuild?.ActiveConfiguration;
        if (activeConfig is not null)
        {
            data["activeConfiguration"] = activeConfig.Name ?? string.Empty;
            data["activePlatform"] = (activeConfig as SolutionConfiguration2)?.PlatformName ?? string.Empty;
        }

        return data;
    }

    private static bool TryGetActiveTextSelection(Document? document, out TextSelection selection)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        selection = null!;
        if (document is null)
        {
            return false;
        }

        try
        {
            if (document.Object("TextDocument") is not TextDocument textDocument)
            {
                return false;
            }

            selection = textDocument.Selection;
            return selection is not null;
        }
        catch (COMException)
        {
            return false;
        }
    }

    private static JArray GetStartupProjects(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var startupProjects = dte.Solution?.SolutionBuild?.StartupProjects;
        return startupProjects switch
        {
            string singleProject when !string.IsNullOrWhiteSpace(singleProject) => new JArray(singleProject),
            object[] projects => new JArray(projects
                .OfType<string>()
                .Where(project => !string.IsNullOrWhiteSpace(project))),
            _ => new JArray(),
        };
    }
}
