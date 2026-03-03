using System.ComponentModel.Design;
using System.Threading.Tasks;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using VsIdeBridge.Commands;
using VsIdeBridge.Services;

namespace VsIdeBridge;

internal static class CommandRegistrar
{
    public static readonly System.Guid CommandSet = new("5C519A88-1F81-402D-BD2A-A0110F704494");

    public static async Task InitializeAsync(VsIdeBridgePackage package, IdeBridgeRuntime runtime)
    {
        var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)).ConfigureAwait(false) as OleMenuCommandService;
        Assumes.Present(commandService);

        runtime.RegisterCommand(new IdeCoreCommands.IdeHelpMenuCommand(package, runtime, commandService));
        runtime.RegisterCommand(new IdeCoreCommands.IdeToggleAllowBridgeEditsMenuCommand(package, runtime, commandService));
        runtime.RegisterCommand(new IdeCoreCommands.IdeToggleGoToEditedPartsMenuCommand(package, runtime, commandService));
        runtime.RegisterCommand(new IdeCoreCommands.IdeHelpCommand(package, runtime, commandService));
        runtime.RegisterCommand(new IdeCoreCommands.IdeSmokeTestCommand(package, runtime, commandService));
        runtime.RegisterCommand(new IdeCoreCommands.IdeGetStateCommand(package, runtime, commandService));
        runtime.RegisterCommand(new IdeCoreCommands.IdeWaitForReadyCommand(package, runtime, commandService));
        runtime.RegisterCommand(new IdeCoreCommands.IdeOpenSolutionCommand(package, runtime, commandService));

        runtime.RegisterCommand(new SearchNavigationCommands.IdeFindTextCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeFindFilesCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeOpenDocumentCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeListDocumentsCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeListOpenTabsCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeActivateDocumentCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeCloseDocumentCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeCloseFileCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeCloseAllExceptCurrentCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeActivateWindowCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeListWindowsCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeExecuteVsCommandCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeFindAllReferencesCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeShowCallHierarchyCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeGetDocumentSliceCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeGetSmartContextForQueryCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeGoToDefinitionCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeGoToImplementationCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeGetFileOutlineCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeSearchSymbolsCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeGetQuickInfoCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeGetDocumentSlicesCommand(package, runtime, commandService));
        runtime.RegisterCommand(new SearchNavigationCommands.IdeGetFileSymbolsCommand(package, runtime, commandService));
        runtime.RegisterCommand(new PatchCommands.IdeApplyUnifiedDiffCommand(package, runtime, commandService));

        runtime.RegisterCommand(new BreakpointCommands.IdeSetBreakpointCommand(package, runtime, commandService));
        runtime.RegisterCommand(new BreakpointCommands.IdeListBreakpointsCommand(package, runtime, commandService));
        runtime.RegisterCommand(new BreakpointCommands.IdeRemoveBreakpointCommand(package, runtime, commandService));
        runtime.RegisterCommand(new BreakpointCommands.IdeClearAllBreakpointsCommand(package, runtime, commandService));
        runtime.RegisterCommand(new BreakpointCommands.IdeEnableBreakpointCommand(package, runtime, commandService));
        runtime.RegisterCommand(new BreakpointCommands.IdeDisableBreakpointCommand(package, runtime, commandService));
        runtime.RegisterCommand(new BreakpointCommands.IdeEnableAllBreakpointsCommand(package, runtime, commandService));
        runtime.RegisterCommand(new BreakpointCommands.IdeDisableAllBreakpointsCommand(package, runtime, commandService));

        runtime.RegisterCommand(new DebugBuildCommands.IdeDebugGetStateCommand(package, runtime, commandService));
        runtime.RegisterCommand(new DebugBuildCommands.IdeDebugStartCommand(package, runtime, commandService));
        runtime.RegisterCommand(new DebugBuildCommands.IdeDebugStopCommand(package, runtime, commandService));
        runtime.RegisterCommand(new DebugBuildCommands.IdeDebugBreakCommand(package, runtime, commandService));
        runtime.RegisterCommand(new DebugBuildCommands.IdeDebugContinueCommand(package, runtime, commandService));
        runtime.RegisterCommand(new DebugBuildCommands.IdeDebugStepOverCommand(package, runtime, commandService));
        runtime.RegisterCommand(new DebugBuildCommands.IdeDebugStepIntoCommand(package, runtime, commandService));
        runtime.RegisterCommand(new DebugBuildCommands.IdeDebugStepOutCommand(package, runtime, commandService));
        runtime.RegisterCommand(new DebugBuildCommands.IdeBuildSolutionCommand(package, runtime, commandService));
        runtime.RegisterCommand(new DebugBuildCommands.IdeGetErrorListCommand(package, runtime, commandService));
        runtime.RegisterCommand(new DebugBuildCommands.IdeGetWarningsCommand(package, runtime, commandService));
        runtime.RegisterCommand(new DebugBuildCommands.IdeBuildAndCaptureErrorsCommand(package, runtime, commandService));
        runtime.RegisterCommand(new IdeCoreCommands.IdeCloseIdeCommand(package, runtime, commandService));
        runtime.RegisterCommand(new IdeCoreCommands.IdeBatchCommandsCommand(package, runtime, commandService));
    }
}
