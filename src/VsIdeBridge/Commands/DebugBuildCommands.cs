using System.ComponentModel.Design;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using VsIdeBridge.Infrastructure;
using VsIdeBridge.Services;

namespace VsIdeBridge.Commands;

internal static class DebugBuildCommands
{
    internal sealed class IdeDebugGetStateCommand : IdeCommandBase
    {
        public IdeDebugGetStateCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x020A)
        {
        }

        protected override string CanonicalName => "Tools.IdeDebugGetState";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DebuggerService.GetStateAsync(context.Dte).ConfigureAwait(true);
            return new CommandExecutionResult("Debugger state captured.", data);
        }
    }

    internal sealed class IdeDebugStartCommand : IdeCommandBase
    {
        public IdeDebugStartCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x020B)
        {
        }

        protected override string CanonicalName => "Tools.IdeDebugStart";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DebuggerService.StartAsync(
                context.Dte,
                args.GetBoolean("wait-for-break", false),
                args.GetInt32("timeout-ms", 120000)).ConfigureAwait(true);
            return new CommandExecutionResult("Debugger started.", data);
        }
    }

    internal sealed class IdeDebugStopCommand : IdeCommandBase
    {
        public IdeDebugStopCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x020C)
        {
        }

        protected override string CanonicalName => "Tools.IdeDebugStop";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DebuggerService.StopAsync(context.Dte).ConfigureAwait(true);
            return new CommandExecutionResult("Debugger stopped.", data);
        }
    }

    internal sealed class IdeDebugBreakCommand : IdeCommandBase
    {
        public IdeDebugBreakCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x020D)
        {
        }

        protected override string CanonicalName => "Tools.IdeDebugBreak";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DebuggerService.BreakAsync(context.Dte).ConfigureAwait(true);
            return new CommandExecutionResult("Debugger break requested.", data);
        }
    }

    internal sealed class IdeDebugContinueCommand : IdeCommandBase
    {
        public IdeDebugContinueCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x020E)
        {
        }

        protected override string CanonicalName => "Tools.IdeDebugContinue";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DebuggerService.ContinueAsync(
                context.Dte,
                args.GetBoolean("wait-for-break", false),
                args.GetInt32("timeout-ms", 120000)).ConfigureAwait(true);
            return new CommandExecutionResult("Debugger continued.", data);
        }
    }

    internal sealed class IdeDebugStepOverCommand : IdeCommandBase
    {
        public IdeDebugStepOverCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x020F)
        {
        }

        protected override string CanonicalName => "Tools.IdeDebugStepOver";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DebuggerService.StepOverAsync(
                context.Dte,
                args.GetInt32("timeout-ms", 120000)).ConfigureAwait(true);
            return new CommandExecutionResult("Debugger step over completed.", data);
        }
    }

    internal sealed class IdeDebugStepIntoCommand : IdeCommandBase
    {
        public IdeDebugStepIntoCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0210)
        {
        }

        protected override string CanonicalName => "Tools.IdeDebugStepInto";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DebuggerService.StepIntoAsync(
                context.Dte,
                args.GetInt32("timeout-ms", 120000)).ConfigureAwait(true);
            return new CommandExecutionResult("Debugger step into completed.", data);
        }
    }

    internal sealed class IdeDebugStepOutCommand : IdeCommandBase
    {
        public IdeDebugStepOutCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0211)
        {
        }

        protected override string CanonicalName => "Tools.IdeDebugStepOut";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DebuggerService.StepOutAsync(
                context.Dte,
                args.GetInt32("timeout-ms", 120000)).ConfigureAwait(true);
            return new CommandExecutionResult("Debugger step out completed.", data);
        }
    }

    internal sealed class IdeBuildSolutionCommand : IdeCommandBase
    {
        public IdeBuildSolutionCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0212)
        {
        }

        protected override string CanonicalName => "Tools.IdeBuildSolution";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.BuildService.BuildSolutionAsync(
                context,
                args.GetInt32("timeout-ms", 600000),
                args.GetString("configuration"),
                args.GetString("platform")).ConfigureAwait(true);

            return new CommandExecutionResult($"Build completed with LastBuildInfo={data["lastBuildInfo"]}.", data);
        }
    }

    internal sealed class IdeGetErrorListCommand : IdeCommandBase
    {
        public IdeGetErrorListCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0213)
        {
        }

        protected override string CanonicalName => "Tools.IdeGetErrorList";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.ErrorListService.GetErrorListAsync(
                context,
                args.GetBoolean("wait-for-intellisense", true),
                args.GetInt32("timeout-ms", 120000)).ConfigureAwait(true);

            return new CommandExecutionResult($"Captured {data["count"]} Error List row(s).", data);
        }
    }

    internal sealed class IdeBuildAndCaptureErrorsCommand : IdeCommandBase
    {
        public IdeBuildAndCaptureErrorsCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0214)
        {
        }

        protected override string CanonicalName => "Tools.IdeBuildAndCaptureErrors";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var timeout = args.GetInt32("timeout-ms", 600000);
            var build = await context.Runtime.BuildService.BuildAndCaptureErrorsAsync(
                context,
                timeout,
                args.GetBoolean("wait-for-intellisense", true)).ConfigureAwait(true);
            var errors = await context.Runtime.ErrorListService.GetErrorListAsync(
                context,
                false,
                timeout).ConfigureAwait(true);

            var data = new JObject
            {
                ["build"] = build,
                ["errors"] = errors,
            };

            return new CommandExecutionResult($"Build finished and captured {errors["count"]} Error List row(s).", data);
        }
    }
}
