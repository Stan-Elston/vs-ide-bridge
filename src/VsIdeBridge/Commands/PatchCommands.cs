using System.ComponentModel.Design;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
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
            var approvedOnce = false;
            if (!context.Runtime.UiSettings.AllowBridgeEdits)
            {
                approvedOnce = await RequestOneTimeEditApprovalAsync(context).ConfigureAwait(true);
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

            data["approval"] = approvedOnce ? "one-time" : "persistent";

            return new CommandExecutionResult($"Applied unified diff to {data["count"]} file(s).", data);
        }

        private static async Task<bool> RequestOneTimeEditApprovalAsync(IdeCommandContext context)
        {
            await context.Logger.LogAsync(
                "IDE Bridge: waiting for edit approval in Visual Studio.",
                context.CancellationToken,
                activatePane: true).ConfigureAwait(true);

            var result = VsShellUtilities.ShowMessageBox(
                context.Package,
                "An external IDE Bridge request wants to edit files in this solution.\r\n\r\n" +
                "Yes: allow this edit request once.\r\n" +
                "No: keep edits blocked.\r\n\r\n" +
                "Use IDE Bridge > Allow Bridge Edits if you want to allow future edits without prompts.",
                "IDE Bridge Approval Required",
                OLEMSGICON.OLEMSGICON_QUERY,
                OLEMSGBUTTON.OLEMSGBUTTON_YESNO,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_SECOND);

            if (result == (int)VSConstants.MessageBoxResult.IDYES)
            {
                await context.Logger.LogAsync(
                    "IDE Bridge: one-time edit approval granted.",
                    context.CancellationToken,
                    activatePane: true).ConfigureAwait(true);
                return true;
            }

            throw new CommandErrorException(
                "edit_approval_denied",
                "Bridge edit approval was denied. Wait for a human to approve the Visual Studio prompt, or enable IDE Bridge > Allow Bridge Edits.",
                new { approvalRequested = true, persistentSettingEnabled = false });
        }
    }
}
