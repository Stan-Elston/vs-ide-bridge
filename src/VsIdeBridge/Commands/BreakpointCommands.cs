using System.ComponentModel.Design;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using VsIdeBridge.Infrastructure;
using VsIdeBridge.Services;

namespace VsIdeBridge.Commands;

internal static class BreakpointCommands
{
    internal sealed class IdeSetBreakpointCommand : IdeCommandBase
    {
        public IdeSetBreakpointCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0206)
        {
        }

        protected override string CanonicalName => "Tools.IdeSetBreakpoint";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.BreakpointService.SetBreakpointAsync(
                context.Dte,
                args.GetRequiredString("file"),
                args.GetInt32("line", 1),
                args.GetInt32("column", 1),
                args.GetString("condition"),
                args.GetEnum("condition-type", "when-true", "when-true", "changed"),
                args.GetInt32("hit-count", 0),
                args.GetEnum("hit-type", "none", "none", "equal", "multiple", "greater-or-equal")).ConfigureAwait(true);

            if (args.GetBoolean("reveal", true))
            {
                var reveal = await context.Runtime.DocumentService.PositionTextSelectionAsync(
                    context.Dte,
                    args.GetRequiredString("file"),
                    documentQuery: null,
                    args.GetInt32("line", 1),
                    args.GetInt32("column", 1),
                    selectWord: false).ConfigureAwait(true);

                data["revealedInEditor"] = true;
                data["reveal"] = reveal;
            }

            return new CommandExecutionResult("Breakpoint set.", data);
        }
    }

    internal sealed class IdeListBreakpointsCommand : IdeCommandBase
    {
        public IdeListBreakpointsCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0207)
        {
        }

        protected override string CanonicalName => "Tools.IdeListBreakpoints";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.BreakpointService.ListBreakpointsAsync(context.Dte).ConfigureAwait(true);
            return new CommandExecutionResult($"Enumerated {data["count"]} breakpoint(s).", data);
        }
    }

    internal sealed class IdeRemoveBreakpointCommand : IdeCommandBase
    {
        public IdeRemoveBreakpointCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0208)
        {
        }

        protected override string CanonicalName => "Tools.IdeRemoveBreakpoint";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.BreakpointService.RemoveBreakpointAsync(
                context.Dte,
                args.GetRequiredString("file"),
                args.GetInt32("line", 1)).ConfigureAwait(true);

            return new CommandExecutionResult($"Removed {data["removedCount"]} breakpoint(s).", data);
        }
    }

    internal sealed class IdeClearAllBreakpointsCommand : IdeCommandBase
    {
        public IdeClearAllBreakpointsCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0209)
        {
        }

        protected override string CanonicalName => "Tools.IdeClearAllBreakpoints";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.BreakpointService.ClearAllBreakpointsAsync(context.Dte).ConfigureAwait(true);
            return new CommandExecutionResult($"Removed {data["removedCount"]} breakpoint(s).", data);
        }
    }
}
