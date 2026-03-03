using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Threading.Tasks;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using VsIdeBridge.Infrastructure;
using VsIdeBridge.Services;

namespace VsIdeBridge.Commands;

internal static class IdeCoreCommands
{
    internal static async Task<CommandExecutionResult> ExecuteBatchAsync(IdeCommandContext context, JArray steps, bool stopOnError)
    {
        var results = new JArray();
        var successCount = 0;
        var failureCount = 0;
        var stoppedEarly = false;

        for (var i = 0; i < steps.Count; i++)
        {
            var step = steps[i] as JObject;
            JObject stepResult;

            if (step == null)
            {
                failureCount++;
                stepResult = new JObject
                {
                    ["index"] = i,
                    ["id"] = JValue.CreateNull(),
                    ["command"] = string.Empty,
                    ["success"] = false,
                    ["summary"] = "Batch entry must be a JSON object.",
                    ["warnings"] = new JArray(),
                    ["data"] = new JObject(),
                    ["error"] = new JObject { ["code"] = "invalid_batch_entry", ["message"] = "Batch entry must be a JSON object." },
                };
            }
            else
            {
                var stepId = (string?)step["id"];
                var commandName = (string?)step["command"] ?? string.Empty;
                var commandArgs = (string?)step["args"] ?? string.Empty;

                if (!context.Runtime.TryGetCommand(commandName, out var cmd))
                {
                    failureCount++;
                    stepResult = new JObject
                    {
                        ["index"] = i,
                        ["id"] = stepId is null ? JValue.CreateNull() : stepId,
                        ["command"] = commandName,
                        ["success"] = false,
                        ["summary"] = $"Unknown command: {commandName}",
                        ["warnings"] = new JArray(),
                        ["data"] = new JObject(),
                        ["error"] = new JObject { ["code"] = "unknown_command", ["message"] = $"Command not registered: {commandName}" },
                    };
                }
                else
                {
                    var parsedArgs = CommandArgumentParser.Parse(commandArgs);
                    try
                    {
                        var result = await cmd.ExecuteDirectAsync(context, parsedArgs).ConfigureAwait(true);
                        successCount++;
                        stepResult = new JObject
                        {
                            ["index"] = i,
                            ["id"] = stepId is null ? JValue.CreateNull() : stepId,
                            ["command"] = commandName,
                            ["success"] = true,
                            ["summary"] = result.Summary,
                            ["warnings"] = result.Warnings,
                            ["data"] = result.Data,
                            ["error"] = JValue.CreateNull(),
                        };
                    }
                    catch (CommandErrorException ex)
                    {
                        failureCount++;
                        stepResult = new JObject
                        {
                            ["index"] = i,
                            ["id"] = stepId is null ? JValue.CreateNull() : stepId,
                            ["command"] = commandName,
                            ["success"] = false,
                            ["summary"] = ex.Message,
                            ["warnings"] = new JArray(),
                            ["data"] = new JObject(),
                            ["error"] = new JObject { ["code"] = ex.Code, ["message"] = ex.Message },
                        };
                    }
                    catch (Exception ex)
                    {
                        failureCount++;
                        stepResult = new JObject
                        {
                            ["index"] = i,
                            ["id"] = stepId is null ? JValue.CreateNull() : stepId,
                            ["command"] = commandName,
                            ["success"] = false,
                            ["summary"] = ex.Message,
                            ["warnings"] = new JArray(),
                            ["data"] = new JObject(),
                            ["error"] = new JObject { ["code"] = "internal_error", ["message"] = ex.Message },
                        };
                    }
                }
            }

            results.Add(stepResult);

            if (stopOnError && !(stepResult.Value<bool?>("success") ?? false))
            {
                stoppedEarly = i < steps.Count - 1;
                break;
            }
        }

        var data = new JObject
        {
            ["batchCount"] = steps.Count,
            ["successCount"] = successCount,
            ["failureCount"] = failureCount,
            ["stoppedEarly"] = stoppedEarly,
            ["results"] = results,
        };

        return new CommandExecutionResult(
            $"Batch: {successCount}/{steps.Count} succeeded, {failureCount} failed.",
            data);
    }

    private static Task<CommandExecutionResult> GetHelpResultAsync()
    {
        var canonicalCommands = new[]
        {
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
            "Tools.IdeGetWarnings",
            "Tools.IdeBuildAndCaptureErrors",
            "Tools.IdeOpenSolution",
            "Tools.IdeGoToDefinition",
            "Tools.IdeGoToImplementation",
            "Tools.IdeGetFileOutline",
            "Tools.IdeGetFileSymbols",
            "Tools.IdeSearchSymbols",
            "Tools.IdeGetQuickInfo",
            "Tools.IdeGetDocumentSlices",
            "Tools.IdeEnableBreakpoint",
            "Tools.IdeDisableBreakpoint",
            "Tools.IdeEnableAllBreakpoints",
            "Tools.IdeDisableAllBreakpoints",
            "Tools.IdeBatchCommands",
        };

        var commands = new JArray();
        var legacyCommands = new JArray();
        foreach (var canonicalCommand in canonicalCommands)
        {
            commands.Add(PipeCommandNames.GetPrimaryName(canonicalCommand));
            legacyCommands.Add(canonicalCommand);
        }

        return Task.FromResult(new CommandExecutionResult(
            "Command catalog written.",
            new JObject
            {
                ["commands"] = commands,
                ["legacyCommands"] = legacyCommands,
                ["note"] = "Pipe requests accept the simple command names in commands[]. The legacy Tools.Ide* names still work in Visual Studio and over the pipe.",
                ["commandDetails"] = new JArray
                {
                    new JObject
                    {
                        ["name"] = "search-symbols",
                        ["legacyName"] = "Tools.IdeSearchSymbols",
                        ["description"] = "Search likely symbol definitions by name across the solution or a filtered path.",
                        ["example"] = @"search-symbols --query ""propose_export_file_name_and_path"" --kind function --path ""src\VsIdeBridge"""
                    },
                    new JObject
                    {
                        ["name"] = "quick-info",
                        ["legacyName"] = "Tools.IdeGetQuickInfo",
                        ["description"] = "Resolve a symbol at a file/line/column and return its definition location plus a surrounding code slice without leaving the source location selected.",
                        ["example"] = @"quick-info --file ""C:\repo\src\foo.cpp"" --line 42 --column 13"
                    },
                    new JObject
                    {
                        ["name"] = "goto-implementation",
                        ["legacyName"] = "Tools.IdeGoToImplementation",
                        ["description"] = "Navigate to one implementation of the symbol at the given source location.",
                        ["example"] = @"goto-implementation --file ""C:\repo\src\foo.cpp"" --line 42 --column 13"
                    },
                    new JObject
                    {
                        ["name"] = "document-slices",
                        ["legacyName"] = "Tools.IdeGetDocumentSlices",
                        ["description"] = "Fetch multiple code slices in one call from either --ranges-file or inline --ranges JSON.",
                        ["example"] = @"document-slices --ranges ""[{\""file\"":\""C:\\repo\\src\\foo.cpp\"",\""line\"":42,\""contextBefore\"":8,\""contextAfter\"":20}]"""
                    },
                    new JObject
                    {
                        ["name"] = "find-text",
                        ["legacyName"] = "Tools.IdeFindText",
                        ["description"] = "Find text across the solution, project, or current document, with optional --path subtree filtering.",
                        ["example"] = @"find-text --query ""OnInit"" --path ""src\libslic3r"""
                    },
                    new JObject
                    {
                        ["name"] = "file-symbols",
                        ["legacyName"] = "Tools.IdeGetFileSymbols",
                        ["description"] = "List symbols from one file with optional --kind filtering; alias over IdeGetFileOutline.",
                        ["example"] = @"file-symbols --file ""C:\repo\src\foo.cpp"" --kind function"
                    },
                    new JObject
                    {
                        ["name"] = "warnings",
                        ["legacyName"] = "Tools.IdeGetWarnings",
                        ["description"] = "Capture only warnings from the Error List, with optional code/path/project filtering.",
                        ["example"] = @"warnings --code C6031 --group-by code"
                    }
                },
                ["recipes"] = new JArray
                {
                    new JObject
                    {
                        ["name"] = "find-symbol-definition",
                        ["summary"] = "Use symbol search before text search when you know the identifier name.",
                        ["command"] = @"search-symbols --query ""propose_export_file_name_and_path"" --kind function"
                    },
                    new JObject
                    {
                        ["name"] = "inspect-symbol-at-location",
                        ["summary"] = "Use quick info to get the destination location and nearby definition context.",
                        ["command"] = @"quick-info --file ""C:\repo\src\foo.cpp"" --line 42 --column 13"
                    },
                    new JObject
                    {
                        ["name"] = "group-current-warnings",
                        ["summary"] = "Filter the Error List down to warnings and group them by code.",
                        ["command"] = @"warnings --group-by code"
                    },
                    new JObject
                    {
                        ["name"] = "fetch-multiple-slices",
                        ["summary"] = "Use inline ranges JSON when you need several code windows in one round-trip.",
                        ["command"] = @"document-slices --ranges ""[{\""file\"":\""C:\\repo\\src\\foo.cpp\"",\""line\"":42,\""contextBefore\"":8,\""contextAfter\"":20}]"""
                    }
                },
                ["example"] = @"state --out ""C:\temp\ide-state.json""",
                ["documentSliceExample"] = @"document-slice --file ""C:\repo\src\foo.cpp"" --start-line 120 --end-line 180 --out ""C:\temp\slice.json""",
                ["documentSlicesExample"] = @"document-slices --ranges ""[{\""file\"":\""C:\\repo\\src\\foo.cpp\"",\""line\"":42,\""contextBefore\"":8,\""contextAfter\"":20}]""",
                ["searchSymbolsExample"] = @"search-symbols --query ""propose_export_file_name_and_path"" --kind function",
                ["quickInfoExample"] = @"quick-info --file ""C:\repo\src\foo.cpp"" --line 42 --column 13",
                ["findTextPathExample"] = @"find-text --query ""OnInit"" --path ""src\libslic3r""",
                ["fileSymbolsExample"] = @"file-symbols --file ""C:\repo\src\foo.cpp"" --kind function",
                ["smartContextExample"] = @"smart-context --query ""where is GUI_App::OnInit used"" --max-contexts 3 --out ""C:\temp\smart-context.json""",
                ["referencesExample"] = @"find-references --file ""C:\repo\src\foo.cpp"" --line 42 --column 13 --out ""C:\temp\references.json""",
                ["callHierarchyExample"] = @"call-hierarchy --file ""C:\repo\src\foo.cpp"" --line 42 --column 13 --out ""C:\temp\call-hierarchy.json""",
                ["applyDiffExample"] = @"apply-diff --patch-file ""C:\temp\change.diff"" --out ""C:\temp\apply-diff.json""",
                ["openSolutionExample"] = @"open-solution --solution ""C:\path\to\solution.sln"" --out ""C:\temp\open-solution.json"""
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

    private static JObject GetUiSettingsData(IdeCommandContext context)
    {
        return new JObject
        {
            ["allowBridgeEdits"] = context.Runtime.UiSettings.AllowBridgeEdits,
            ["goToEditedParts"] = context.Runtime.UiSettings.GoToEditedParts,
        };
    }

    private static Task<CommandExecutionResult> ToggleAllowBridgeEditsAsync(IdeCommandContext context)
    {
        var enabled = !context.Runtime.UiSettings.AllowBridgeEdits;
        context.Runtime.UiSettings.AllowBridgeEdits = enabled;
        return Task.FromResult(new CommandExecutionResult(
            enabled ? "Bridge edits enabled." : "Bridge edits disabled.",
            GetUiSettingsData(context)));
    }

    private static Task<CommandExecutionResult> ToggleGoToEditedPartsAsync(IdeCommandContext context)
    {
        var enabled = !context.Runtime.UiSettings.GoToEditedParts;
        context.Runtime.UiSettings.GoToEditedParts = enabled;
        return Task.FromResult(new CommandExecutionResult(
            enabled ? "Go To Edited Parts enabled." : "Go To Edited Parts disabled.",
            GetUiSettingsData(context)));
    }

    private static string? TryResolveReadmePath(IdeCommandContext context)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            var solutionPath = context.Dte.Solution?.FullName;
            if (string.IsNullOrWhiteSpace(solutionPath))
            {
                return null;
            }

            var solutionDirectory = Path.GetDirectoryName(solutionPath);
            if (string.IsNullOrWhiteSpace(solutionDirectory))
            {
                return null;
            }

            var readmePath = Path.Combine(solutionDirectory, "README.md");
            return File.Exists(readmePath) ? readmePath : null;
        }
        catch
        {
            return null;
        }
    }

    private static async Task<CommandExecutionResult> ShowHelpMenuAsync(IdeCommandContext context)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);

        var readmePath = TryResolveReadmePath(context);
        if (!string.IsNullOrWhiteSpace(readmePath))
        {
            context.Dte.ItemOperations.OpenFile(readmePath);
        }

        var message = string.IsNullOrWhiteSpace(readmePath)
            ? "Use the Command Window with Tools.IdeHelp for the full command catalog. The README could not be resolved from the current solution."
            : $"Opened README: {readmePath}{Environment.NewLine}{Environment.NewLine}Use the Command Window with Tools.IdeHelp for the full command catalog.";

        VsShellUtilities.ShowMessageBox(
            context.Package,
            message,
            "VS IDE Bridge",
            OLEMSGICON.OLEMSGICON_INFO,
            OLEMSGBUTTON.OLEMSGBUTTON_OK,
            OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

        return new CommandExecutionResult(
            string.IsNullOrWhiteSpace(readmePath) ? "Displayed IDE Bridge help." : "Opened IDE Bridge help.",
            new JObject
            {
                ["readmePath"] = readmePath is null ? JValue.CreateNull() : readmePath,
                ["commandWindowHelp"] = "Tools.IdeHelp",
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
            return ShowHelpMenuAsync(context);
        }
    }

    internal sealed class IdeToggleAllowBridgeEditsMenuCommand : IdeCommandBase
    {
        public IdeToggleAllowBridgeEditsMenuCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0103, acceptsParameters: false)
        {
            MenuCommand.BeforeQueryStatus += (_, _) => MenuCommand.Checked = Runtime.UiSettings.AllowBridgeEdits;
        }

        protected override string CanonicalName => "Tools.VsIdeBridgeToggleAllowBridgeEdits";

        protected override Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            return ToggleAllowBridgeEditsAsync(context);
        }
    }

    internal sealed class IdeToggleGoToEditedPartsMenuCommand : IdeCommandBase
    {
        public IdeToggleGoToEditedPartsMenuCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0104, acceptsParameters: false)
        {
            MenuCommand.BeforeQueryStatus += (_, _) => MenuCommand.Checked = Runtime.UiSettings.GoToEditedParts;
        }

        protected override string CanonicalName => "Tools.VsIdeBridgeToggleGoToEditedParts";

        protected override Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            return ToggleGoToEditedPartsAsync(context);
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

    internal sealed class IdeOpenSolutionCommand : IdeCommandBase
    {
        public IdeOpenSolutionCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0224)
        {
        }

        protected override string CanonicalName => "Tools.IdeOpenSolution";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var solutionPath = args.GetRequiredString("solution");
            if (!File.Exists(solutionPath))
            {
                throw new CommandErrorException("file_not_found", $"Solution file not found: {solutionPath}");
            }
            if (!string.Equals(Path.GetExtension(solutionPath), ".sln", StringComparison.OrdinalIgnoreCase))
            {
                throw new CommandErrorException("invalid_file_type", $"File is not a solution file: {solutionPath}");
            }
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            context.Dte.Solution.Open(solutionPath);
            return new CommandExecutionResult("Solution opened.", new JObject { ["solutionPath"] = solutionPath });
        }
    }

    internal sealed class IdeBatchCommandsCommand : IdeCommandBase
    {
        public IdeBatchCommandsCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService)
            : base(package, runtime, commandService, 0x0225)
        {
        }

        protected override string CanonicalName => "Tools.IdeBatchCommands";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            var batchFile = args.GetRequiredString("batch-file");
            if (!File.Exists(batchFile))
            {
                throw new CommandErrorException("file_not_found", $"Batch file not found: {batchFile}");
            }

            var json = File.ReadAllText(batchFile);
            JArray steps;
            try
            {
                steps = JArray.Parse(json);
            }
            catch (JsonException ex)
            {
                throw new CommandErrorException("invalid_json", $"Failed to parse batch file: {ex.Message}");
            }

            var stopOnError = args.GetBoolean("stop-on-error", false);
            return await ExecuteBatchAsync(context, steps, stopOnError).ConfigureAwait(true);
        }
    }
}
