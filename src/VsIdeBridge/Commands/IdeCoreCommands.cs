using System;
using System.ComponentModel.Design;
using System.IO;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using VsIdeBridge.Infrastructure;
using VsIdeBridge.Services;

namespace VsIdeBridge.Commands;

internal static class IdeCoreCommands
{
    private static Task<CommandExecutionResult> GetHelpResultAsync()
    {
        var commands = new JArray(
            "Tools.IdeGetState",
            "Tools.IdeWaitForReady",
            "Tools.IdeFindText",
            "Tools.IdeFindFiles",
            "Tools.IdeOpenDocument",
            "Tools.IdeListDocuments",
            "Tools.IdeListOpenTabs",
            "Tools.IdeActivateDocument",
            "Tools.IdeCloseDocument",
            "Tools.IdeCloseFile",
            "Tools.IdeCloseAllExceptCurrent",
            "Tools.IdeActivateWindow",
            "Tools.IdeListWindows",
            "Tools.IdeExecuteVsCommand",
            "Tools.IdeFindAllReferences",
            "Tools.IdeShowCallHierarchy",
            "Tools.IdeGetDocumentSlice",
            "Tools.IdeGetSmartContextForQuery",
            "Tools.IdeApplyUnifiedDiff",
            "Tools.IdeSetBreakpoint",
            "Tools.IdeListBreakpoints",
            "Tools.IdeRemoveBreakpoint",
            "Tools.IdeClearAllBreakpoints",
            "Tools.IdeDebugGetState",
            "Tools.IdeDebugStart",
            "Tools.IdeDebugStop",
            "Tools.IdeDebugBreak",
            "Tools.IdeDebugContinue",
            "Tools.IdeDebugStepOver",
            "Tools.IdeDebugStepInto",
            "Tools.IdeDebugStepOut",
            "Tools.IdeBuildSolution",
            "Tools.IdeGetErrorList",
            "Tools.IdeBuildAndCaptureErrors");

        return Task.FromResult(new CommandExecutionResult(
            "Command catalog written.",
            new JObject
            {
                ["commands"] = commands,
                ["example"] = @"Tools.IdeGetState --out ""C:\temp\ide-state.json""",
                ["documentSliceExample"] = @"Tools.IdeGetDocumentSlice --file ""C:\repo\src\foo.cpp"" --start-line 120 --end-line 180 --out ""C:\temp\slice.json""",
                ["smartContextExample"] = @"Tools.IdeGetSmartContextForQuery --query ""where is GUI_App::OnInit used"" --max-contexts 3 --out ""C:\temp\smart-context.json""",
                ["referencesExample"] = @"Tools.IdeFindAllReferences --file ""C:\repo\src\foo.cpp"" --line 42 --column 13 --out ""C:\temp\references.json""",
                ["callHierarchyExample"] = @"Tools.IdeShowCallHierarchy --file ""C:\repo\src\foo.cpp"" --line 42 --column 13 --out ""C:\temp\call-hierarchy.json""",
                ["applyDiffExample"] = @"Tools.IdeApplyUnifiedDiff --patch-file ""C:\temp\change.diff"" --out ""C:\temp\apply-diff.json""",
            }));
    }

    private static async Task<CommandExecutionResult> GetSmokeTestResultAsync(IdeCommandContext context)
    {
        var state = await context.Runtime.IdeStateService.GetStateAsync(context.Dte).ConfigureAwait(true);
        return new CommandExecutionResult(
            "Smoke test captured IDE state.",
            new JObject
            {
                ["success"] = true,
                ["state"] = state,
            });
    }

    internal sealed class IdeHelpMenuCommand : IdeCommandBase
    {
        public IdeHelpMenuCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0102, acceptsParameters: false)
        {
        }

        protected override string CanonicalName => "Tools.VsIdeBridgeHelpMenu";

        protected override Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            return GetHelpResultAsync();
        }
    }

    internal sealed class IdeSmokeTestMenuCommand : IdeCommandBase
    {
        public IdeSmokeTestMenuCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0103, acceptsParameters: false)
        {
        }

        protected override string CanonicalName => "Tools.VsIdeBridgeSmokeTestMenu";

        protected override Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            return GetSmokeTestResultAsync(context);
        }
    }

    internal sealed class IdeHelpCommand : IdeCommandBase
    {
        public IdeHelpCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0100)
        {
        }

        protected override string CanonicalName => "Tools.IdeHelp";

        protected override Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            return GetHelpResultAsync();
        }
    }

    internal sealed class IdeSmokeTestCommand : IdeCommandBase
    {
        public IdeSmokeTestCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0101)
        {
        }

        protected override string CanonicalName => "Tools.IdeSmokeTest";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            return await GetSmokeTestResultAsync(context).ConfigureAwait(true);
        }
    }

    internal sealed class IdeGetStateCommand : IdeCommandBase
    {
        public IdeGetStateCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0200)
        {
        }

        protected override string CanonicalName => "Tools.IdeGetState";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var state = await context.Runtime.IdeStateService.GetStateAsync(context.Dte).ConfigureAwait(true);
            return new CommandExecutionResult("IDE state captured.", state);
        }
    }

    internal sealed class IdeWaitForReadyCommand : IdeCommandBase
    {
        public IdeWaitForReadyCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0201)
        {
        }

        protected override string CanonicalName => "Tools.IdeWaitForReady";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var timeout = args.GetInt32("timeout-ms", 120000);
            var data = await context.Runtime.ReadinessService.WaitForReadyAsync(context, timeout).ConfigureAwait(true);
            return new CommandExecutionResult("Readiness wait completed.", data);
        }
    }
}
