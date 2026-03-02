using System.Threading.Tasks;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed class IdeBridgeRuntime
{
    private IdeBridgeRuntime(
        OutputPaneLogger logger,
        IdeStateService ideStateService,
        ReadinessService readinessService,
        SearchService searchService,
        DocumentService documentService,
        WindowService windowService,
        VsCommandService vsCommandService,
        PatchService patchService,
        BreakpointService breakpointService,
        DebuggerService debuggerService,
        BuildService buildService,
        ErrorListService errorListService)
    {
        Logger = logger;
        IdeStateService = ideStateService;
        ReadinessService = readinessService;
        SearchService = searchService;
        DocumentService = documentService;
        WindowService = windowService;
        VsCommandService = vsCommandService;
        PatchService = patchService;
        BreakpointService = breakpointService;
        DebuggerService = debuggerService;
        BuildService = buildService;
        ErrorListService = errorListService;
    }

    public OutputPaneLogger Logger { get; }

    public IdeStateService IdeStateService { get; }

    public ReadinessService ReadinessService { get; }

    public SearchService SearchService { get; }

    public DocumentService DocumentService { get; }

    public WindowService WindowService { get; }

    public VsCommandService VsCommandService { get; }

    public PatchService PatchService { get; }

    public BreakpointService BreakpointService { get; }

    public DebuggerService DebuggerService { get; }

    public BuildService BuildService { get; }

    public ErrorListService ErrorListService { get; }

    public static async Task<IdeBridgeRuntime> CreateAsync(VsIdeBridgePackage package)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
        var dte = await package.GetServiceAsync(typeof(SDTE)).ConfigureAwait(true) as DTE2;
        Assumes.Present(dte);

        var logger = new OutputPaneLogger(package, dte);
        var documentService = new DocumentService();
        var readinessService = new ReadinessService();
        var searchService = new SearchService();
        var errorListService = new ErrorListService(readinessService);
        var buildService = new BuildService(readinessService);

        return new IdeBridgeRuntime(
            logger,
            new IdeStateService(),
            readinessService,
            searchService,
            documentService,
            new WindowService(),
            new VsCommandService(),
            new PatchService(),
            new BreakpointService(),
            new DebuggerService(),
            buildService,
            errorListService);
    }
}
