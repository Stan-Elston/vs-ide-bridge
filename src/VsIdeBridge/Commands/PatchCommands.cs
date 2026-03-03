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
            if (!context.Runtime.UiSettings.AllowBridgeEdits)
            {
                throw new CommandErrorException(
                    "edits_disabled",
                    "Bridge edits are disabled. Enable IDE Bridge > Allow Bridge Edits before applying changes.");
            }

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

            var openChangedFiles = args.Has("open-changed-files")
                ? args.GetBoolean("open-changed-files", context.Runtime.UiSettings.GoToEditedParts)
                : context.Runtime.UiSettings.GoToEditedParts;

            var data = await context.Runtime.PatchService.ApplyUnifiedDiffAsync(
                context.Dte,
                context.Runtime.DocumentService,
                args.GetString("patch-file"),
                patchText,
                args.GetString("base-directory"),
                openChangedFiles,
                args.GetBoolean("save-changed-files", false)).ConfigureAwait(true);

            return new CommandExecutionResult($"Applied unified diff to {data["count"]} file(s).", data);
        }
    }
}
