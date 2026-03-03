using System;
using System.ComponentModel.Design;
using System.IO;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using VsIdeBridge.Infrastructure;
using VsIdeBridge.Services;

namespace VsIdeBridge.Commands;

internal static class SearchNavigationCommands
{
    internal sealed class IdeFindTextCommand : IdeCommandBase
    {
        public IdeFindTextCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0202)
        {
        }

        protected override string CanonicalName => "Tools.IdeFindText";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var query = args.GetRequiredString("query");
            var scope = args.GetEnum("scope", "solution", "solution", "project", "document", "open");
            var project = args.GetString("project");
            var data = await context.Runtime.SearchService.FindTextAsync(
                context,
                query,
                scope,
                args.GetBoolean("match-case", false),
                args.GetBoolean("whole-word", false),
                args.GetBoolean("regex", false),
                args.GetInt32("results-window", 1),
                project,
                args.GetString("path")).ConfigureAwait(true);

            return new CommandExecutionResult($"Found {data["count"]} match(es).", data);
        }
    }

    internal sealed class IdeFindFilesCommand : IdeCommandBase
    {
        public IdeFindFilesCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0203)
        {
        }

        protected override string CanonicalName => "Tools.IdeFindFiles";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var query = args.GetRequiredString("query");
            var data = await context.Runtime.SearchService.FindFilesAsync(context, query).ConfigureAwait(true);
            return new CommandExecutionResult($"Found {data["count"]} file(s).", data);
        }
    }

    internal sealed class IdeOpenDocumentCommand : IdeCommandBase
    {
        public IdeOpenDocumentCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0204)
        {
        }

        protected override string CanonicalName => "Tools.IdeOpenDocument";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DocumentService.OpenDocumentAsync(
                context.Dte,
                args.GetRequiredString("file"),
                args.GetInt32("line", 1),
                args.GetInt32("column", 1)).ConfigureAwait(true);

            return new CommandExecutionResult("Document activated.", data);
        }
    }

    internal sealed class IdeListDocumentsCommand : IdeCommandBase
    {
        public IdeListDocumentsCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0216)
        {
        }

        protected override string CanonicalName => "Tools.IdeListDocuments";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DocumentService.ListOpenDocumentsAsync(context.Dte).ConfigureAwait(true);
            return new CommandExecutionResult($"Listed {data["count"]} open document(s).", data);
        }
    }

    internal sealed class IdeActivateDocumentCommand : IdeCommandBase
    {
        public IdeActivateDocumentCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0217)
        {
        }

        protected override string CanonicalName => "Tools.IdeActivateDocument";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DocumentService
                .ActivateOpenDocumentAsync(context.Dte, args.GetRequiredString("query"))
                .ConfigureAwait(true);
            return new CommandExecutionResult("Document tab activated.", data);
        }
    }

    internal sealed class IdeCloseDocumentCommand : IdeCommandBase
    {
        public IdeCloseDocumentCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0218)
        {
        }

        protected override string CanonicalName => "Tools.IdeCloseDocument";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DocumentService
                .CloseOpenDocumentsAsync(
                    context.Dte,
                    args.GetString("query"),
                    args.GetBoolean("all", false),
                    args.GetBoolean("save", false))
                .ConfigureAwait(true);
            return new CommandExecutionResult($"Closed {data["count"]} document(s).", data);
        }
    }

    internal sealed class IdeListOpenTabsCommand : IdeCommandBase
    {
        public IdeListOpenTabsCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x021D)
        {
        }

        protected override string CanonicalName => "Tools.IdeListOpenTabs";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DocumentService.ListOpenTabsAsync(context.Dte).ConfigureAwait(true);
            return new CommandExecutionResult($"Listed {data["count"]} open tab(s).", data);
        }
    }

    internal sealed class IdeCloseFileCommand : IdeCommandBase
    {
        public IdeCloseFileCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x021E)
        {
        }

        protected override string CanonicalName => "Tools.IdeCloseFile";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DocumentService
                .CloseFileAsync(
                    context.Dte,
                    args.GetString("file"),
                    args.GetString("query"),
                    args.GetBoolean("save", false))
                .ConfigureAwait(true);

            return new CommandExecutionResult($"Closed {data["count"]} file(s).", data);
        }
    }

    internal sealed class IdeCloseAllExceptCurrentCommand : IdeCommandBase
    {
        public IdeCloseAllExceptCurrentCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x021F)
        {
        }

        protected override string CanonicalName => "Tools.IdeCloseAllExceptCurrent";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DocumentService
                .CloseAllExceptCurrentAsync(context.Dte, args.GetBoolean("save", false))
                .ConfigureAwait(true);

            return new CommandExecutionResult($"Closed {data["count"]} file(s).", data);
        }
    }

    internal sealed class IdeActivateWindowCommand : IdeCommandBase
    {
        public IdeActivateWindowCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0205)
        {
        }

        protected override string CanonicalName => "Tools.IdeActivateWindow";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.WindowService.ActivateWindowAsync(context.Dte, args.GetRequiredString("window")).ConfigureAwait(true);
            return new CommandExecutionResult("Window activated.", data);
        }
    }

    internal sealed class IdeListWindowsCommand : IdeCommandBase
    {
        public IdeListWindowsCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0219)
        {
        }

        protected override string CanonicalName => "Tools.IdeListWindows";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.WindowService.ListWindowsAsync(context.Dte, args.GetString("query")).ConfigureAwait(true);
            return new CommandExecutionResult($"Listed {data["count"]} window(s).", data);
        }
    }

    internal sealed class IdeExecuteVsCommandCommand : IdeCommandBase
    {
        public IdeExecuteVsCommandCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x021A)
        {
        }

        protected override string CanonicalName => "Tools.IdeExecuteVsCommand";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.VsCommandService
                .ExecutePositionedCommandAsync(
                    context.Dte,
                    context.Runtime.DocumentService,
                    args.GetRequiredString("command"),
                    args.GetString("args"),
                    args.GetString("file"),
                    args.GetString("document"),
                    args.GetNullableInt32("line"),
                    args.GetNullableInt32("column"),
                    args.GetBoolean("select-word", false))
                .ConfigureAwait(true);
            return new CommandExecutionResult("Visual Studio command executed.", data);
        }
    }

    internal sealed class IdeFindAllReferencesCommand : IdeCommandBase
    {
        private static readonly string[] CandidateCommands = { "Edit.FindAllReferences" };

        public IdeFindAllReferencesCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x021B)
        {
        }

        protected override string CanonicalName => "Tools.IdeFindAllReferences";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.VsCommandService.ExecuteSymbolCommandAsync(
                    context.Dte,
                    context.Runtime.DocumentService,
                    context.Runtime.WindowService,
                    CandidateCommands,
                    args.GetString("file"),
                    args.GetString("document"),
                    args.GetNullableInt32("line"),
                    args.GetNullableInt32("column"),
                    args.GetBoolean("select-word", false),
                    "references",
                    args.GetBoolean("activate-window", true),
                    args.GetInt32("timeout-ms", 5000))
                .ConfigureAwait(true);

            return new CommandExecutionResult("Find All References executed.", data);
        }
    }

    internal sealed class IdeShowCallHierarchyCommand : IdeCommandBase
    {
        private static readonly string[] CandidateCommands = { "View.CallHierarchy" };

        public IdeShowCallHierarchyCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x021C)
        {
        }

        protected override string CanonicalName => "Tools.IdeShowCallHierarchy";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.VsCommandService.ExecuteSymbolCommandAsync(
                    context.Dte,
                    context.Runtime.DocumentService,
                    context.Runtime.WindowService,
                    CandidateCommands,
                    args.GetString("file"),
                    args.GetString("document"),
                    args.GetNullableInt32("line"),
                    args.GetNullableInt32("column"),
                    args.GetBoolean("select-word", false),
                    "Call Hierarchy",
                    args.GetBoolean("activate-window", true),
                    args.GetInt32("timeout-ms", 5000))
                .ConfigureAwait(true);

            return new CommandExecutionResult("Call Hierarchy executed.", data);
        }
    }

    internal sealed class IdeGetDocumentSliceCommand : IdeCommandBase
    {
        public IdeGetDocumentSliceCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0215)
        {
        }

        protected override string CanonicalName => "Tools.IdeGetDocumentSlice";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var baseLine = args.GetInt32("line", 1);
            var contextBefore = args.GetInt32("context-before", 0);
            var contextAfter = args.GetInt32("context-after", 0);
            var startLine = args.GetInt32("start-line", Math.Max(1, baseLine - contextBefore));
            var endLine = args.GetInt32("end-line", Math.Max(startLine, baseLine + contextAfter));

            var data = await context.Runtime.DocumentService.GetDocumentSliceAsync(
                context.Dte,
                args.GetString("file"),
                startLine,
                endLine,
                args.GetBoolean("include-line-numbers", true)).ConfigureAwait(true);

            return new CommandExecutionResult(
                $"Captured lines {data["actualStartLine"]}-{data["actualEndLine"]}.",
                data);
        }
    }

    internal sealed class IdeGetSmartContextForQueryCommand : IdeCommandBase
    {
        public IdeGetSmartContextForQueryCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0220)
        {
        }

        protected override string CanonicalName => "Tools.IdeGetSmartContextForQuery";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.SearchService.GetSmartContextForQueryAsync(
                context,
                args.GetRequiredString("query"),
                args.GetEnum("scope", "solution", "solution", "project", "document"),
                args.GetBoolean("match-case", false),
                args.GetBoolean("whole-word", false),
                args.GetBoolean("regex", false),
                args.GetString("project"),
                args.GetInt32("max-contexts", 5),
                args.GetInt32("context-before", 20),
                args.GetInt32("context-after", 20),
                args.GetBoolean("populate-results-window", true),
                args.GetInt32("results-window", 1)).ConfigureAwait(true);

            return new CommandExecutionResult(
                $"Captured {data["contextCount"]} smart context(s) from {data["totalMatchCount"]} match(es).",
                data);
        }
    }

    internal sealed class IdeGoToDefinitionCommand : IdeCommandBase
    {
        public IdeGoToDefinitionCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0222)
        {
        }

        protected override string CanonicalName => "Tools.IdeGoToDefinition";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DocumentService.GoToDefinitionAsync(
                    context.Dte,
                    args.GetString("file"),
                    args.GetString("document"),
                    args.GetNullableInt32("line"),
                    args.GetNullableInt32("column"))
                .ConfigureAwait(true);

            var found = (bool?)data["definitionFound"] == true;
            return new CommandExecutionResult(
                found ? "Navigated to definition." : "Go To Definition executed (location unchanged).",
                data);
        }
    }

    internal sealed class IdeGoToImplementationCommand : IdeCommandBase
    {
        public IdeGoToImplementationCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x022F)
        {
        }

        protected override string CanonicalName => "Tools.IdeGoToImplementation";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DocumentService.GoToImplementationAsync(
                    context.Dte,
                    args.GetString("file"),
                    args.GetString("document"),
                    args.GetNullableInt32("line"),
                    args.GetNullableInt32("column"))
                .ConfigureAwait(true);

            var found = (bool?)data["implementationFound"] == true;
            return new CommandExecutionResult(
                found ? "Navigated to implementation." : "Go To Implementation executed (location unchanged).",
                data);
        }
    }

    internal sealed class IdeGetFileOutlineCommand : IdeCommandBase
    {
        public IdeGetFileOutlineCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0223)
        {
        }

        protected override string CanonicalName => "Tools.IdeGetFileOutline";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DocumentService.GetFileOutlineAsync(
                    context.Dte,
                    args.GetString("file"),
                    args.GetInt32("max-depth", 3),
                    args.GetString("kind"))
                .ConfigureAwait(true);

            return new CommandExecutionResult($"Found {data["count"]} symbol(s).", data);
        }
    }

    internal sealed class IdeSearchSymbolsCommand : IdeCommandBase
    {
        public IdeSearchSymbolsCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0226)
        {
        }

        protected override string CanonicalName => "Tools.IdeSearchSymbols";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var symbolQuery = args.GetString("query") ?? args.GetString("name");
            if (string.IsNullOrWhiteSpace(symbolQuery))
            {
                throw new CommandErrorException("invalid_arguments", "Missing required argument --query.");
            }

            var data = await context.Runtime.SearchService.SearchSymbolsAsync(
                context,
                symbolQuery!,
                args.GetEnum("kind", "all", "all", "function", "class", "struct", "enum", "namespace", "interface", "member", "type"),
                args.GetEnum("scope", "solution", "solution", "project", "document", "open"),
                args.GetBoolean("match-case", false),
                args.GetString("project"),
                args.GetString("path"),
                args.GetInt32("max", 50)).ConfigureAwait(true);

            return new CommandExecutionResult($"Found {data["count"]} symbol match(es).", data);
        }
    }

    internal sealed class IdeGetQuickInfoCommand : IdeCommandBase
    {
        public IdeGetQuickInfoCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0227)
        {
        }

        protected override string CanonicalName => "Tools.IdeGetQuickInfo";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DocumentService.GetQuickInfoAsync(
                context.Dte,
                args.GetString("file"),
                args.GetString("document"),
                args.GetNullableInt32("line"),
                args.GetNullableInt32("column"),
                args.GetInt32("context-lines", 10)).ConfigureAwait(true);

            var found = (bool?)data["definitionFound"] == true;
            return new CommandExecutionResult(
                found ? $"Quick info: definition found for '{data["word"]}'." : $"Quick info: no definition found for '{data["word"]}'.",
                data);
        }
    }

    internal sealed class IdeGetDocumentSlicesCommand : IdeCommandBase
    {
        public IdeGetDocumentSlicesCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0228)
        {
        }

        protected override string CanonicalName => "Tools.IdeGetDocumentSlices";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            JArray ranges;
            var rangesJson = args.GetString("ranges");
            var rangesFile = args.GetString("ranges-file");
            if (!string.IsNullOrWhiteSpace(rangesJson))
            {
                try
                {
                    ranges = JArray.Parse(rangesJson);
                }
                catch (Exception ex)
                {
                    throw new CommandErrorException("invalid_json", $"Failed to parse --ranges JSON: {ex.Message}");
                }
            }
            else
            {
                if (string.IsNullOrWhiteSpace(rangesFile))
                {
                    throw new CommandErrorException("invalid_arguments", "Specify either --ranges or --ranges-file.");
                }

                if (!File.Exists(rangesFile))
                {
                    throw new CommandErrorException("file_not_found", $"Ranges file not found: {rangesFile}");
                }

                try
                {
                    ranges = JArray.Parse(File.ReadAllText(rangesFile!));
                }
                catch (Exception ex)
                {
                    throw new CommandErrorException("invalid_json", $"Failed to parse ranges file: {ex.Message}");
                }
            }

            var data = await context.Runtime.DocumentService.GetDocumentSlicesAsync(context.Dte, ranges).ConfigureAwait(true);
            return new CommandExecutionResult($"Captured {data["count"]} slice(s).", data);
        }
    }

    internal sealed class IdeGetFileSymbolsCommand : IdeCommandBase
    {
        public IdeGetFileSymbolsCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x022D)
        {
        }

        protected override string CanonicalName => "Tools.IdeGetFileSymbols";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var data = await context.Runtime.DocumentService.GetFileOutlineAsync(
                    context.Dte,
                    args.GetString("file"),
                    args.GetInt32("max-depth", 8),
                    args.GetString("kind"))
                .ConfigureAwait(true);

            return new CommandExecutionResult($"Found {data["count"]} symbol(s).", data);
        }
    }
}
