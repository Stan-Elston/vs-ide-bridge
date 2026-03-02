using System.ComponentModel.Design;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using VsIdeBridge.Infrastructure;
using VsIdeBridge.Services;

namespace VsIdeBridge.Commands;

internal static class PatchCommands
{
    internal sealed class IdeApplyUnifiedDiffCommand : IdeCommandBase
    {
        public IdeApplyUnifiedDiffCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0221)
        {
        }

        protected override string CanonicalName => "Tools.IdeApplyUnifiedDiff";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            string? patchText = null;
            var patchTextBase64 = args.GetString("patch-text-base64");
            if (!string.IsNullOrWhiteSpace(patchTextBase64))
            {
                try
                {
                    patchText = Encoding.UTF8.GetString(System.Convert.FromBase64String(patchTextBase64));
                }
                catch (System.FormatException ex)
                {
                    throw new CommandErrorException(
                        "invalid_arguments",
                        "Value passed to --patch-text-base64 was not valid base64.",
                        new { exception = ex.Message });
                }
            }

            var data = await context.Runtime.PatchService.ApplyUnifiedDiffAsync(
                context.Dte,
                context.Runtime.DocumentService,
                args.GetString("patch-file"),
                patchText,
                args.GetString("base-directory"),
                args.GetBoolean("open-changed-files", true)).ConfigureAwait(true);

            return new CommandExecutionResult($"Applied unified diff to {data["count"]} file(s).", data);
        }
    }
}
