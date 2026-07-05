using System.Text;
using System.IO;
using System.Linq;
using System.Text.Json.Nodes;
using static VsIdeBridgeService.ArgBuilder;
using static VsIdeBridgeService.SchemaHelpers;
using VsIdeBridge.Diagnostics;
using VsIdeBridge.Shared;
using VsIdeBridge.Tooling.Handles;
using VsIdeBridge.Tooling.Patches;
using VsIdeBridgeService.SystemTools;

namespace VsIdeBridgeService;

internal static partial class ToolCatalog
{
    private const string ApplyDiffTool = "apply_diff";
    private const string ReadFileTool = "read_file";
    private const string ListTabsTool = "list_tabs";

    private static string EncodeUtf8ToBase64(string? text)
    {
        return Convert.ToBase64String(Encoding.UTF8.GetBytes(text ?? string.Empty));
    }

    /// <summary>
    /// Resolves a relative path to an existing file by searching upward from
    /// <paramref name="solutionDir"/>. When multiple candidates exist (e.g., both
    /// <c>src/foo.cpp</c> and <c>build/src/foo.cpp</c>), source-tree paths are
    /// preferred over build-artifact paths, matching the DocumentService behavior.
    /// Falls back to <c>Path.Combine(solutionDir, fileArg)</c> when no candidate exists.
    /// </summary>
    private static string ResolveExistingFilePath(string fileArg, string solutionDir)
    {
        if (System.IO.Path.IsPathRooted(fileArg))
            return fileArg;

        string normalizedArg = fileArg.Replace('/', System.IO.Path.DirectorySeparatorChar);
        List<string> candidates = [];
        string current = solutionDir;
        for (int depth = 0; depth < 6 && !string.IsNullOrWhiteSpace(current); depth++)
        {
            string candidate = System.IO.Path.Combine(current, normalizedArg);
            if (File.Exists(candidate))
                candidates.Add(candidate);
            string? parent = System.IO.Path.GetDirectoryName(current);
            if (string.IsNullOrWhiteSpace(parent) || string.Equals(parent, current, StringComparison.OrdinalIgnoreCase))
                break;
            current = parent;
        }

        if (candidates.Count == 0)
            return System.IO.Path.Combine(solutionDir, normalizedArg);

        return candidates
            .OrderByDescending(p => ContainsPathSegment(p, "src"))
            .ThenBy(p => ContainsPathSegment(p, "build") || ContainsPathSegment(p, "bin") || ContainsPathSegment(p, "obj"))
            .ThenBy(p => p.Length)
            .First();
    }

    private static bool ContainsPathSegment(string path, string segment)
    {
        string sep = System.IO.Path.DirectorySeparatorChar.ToString();
        string norm = path.Replace('/', System.IO.Path.DirectorySeparatorChar);
        return norm.Contains(sep + segment + sep, StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildApplyDiffArgs(ApplyDiffRequest request)
    {
        return Build(
            ("patch-text-base64", request.EncodedDiff),
            ("open-changed-files", "true"),
            ("save-changed-files", "true"),
            ("replace-all", request.ReplaceAll ? "true" : null));
    }

    private static string BuildWriteFileArgs(JsonObject? args)
    {
        return Build(
            (FileArg, OptionalString(args, FileArg)),
            ("content-base64", EncodeUtf8ToBase64(OptionalString(args, "content"))));
    }

    private static IEnumerable<ToolEntry> DocumentTools()
        =>
        DocumentEditTools()
            .Concat(DocumentTabTools())
            .Concat(WindowCommandTools())
            .Concat(FileOperationTools())
            .Concat(SolutionSystemTools());

    private static IEnumerable<ToolEntry> DocumentEditTools()
        =>
        ApplyDiffTools()
            .Concat(WriteFileTools());

    private static async Task<JsonObject> ApplyMultipleEditsAsync(
        JsonNode? id, JsonObject? args, JsonArray editsArray, BridgeConnection bridge)
    {
        if (args?.ContainsKey("replace_all") == true)
        {
            throw new McpRequestException(id, McpErrorCodes.InvalidParams,
                "apply_diff top-level 'replace_all' is only valid with a single 'file' + 'old_content' + " +
                "'new_content' edit. For batches, set 'replace_all' inside each edits[] item so it stays scoped " +
                "to that item's explicit file.");
        }

        JsonArray editResults = [];
        int applied = 0;
        int failed = 0;
        int skipped = 0;
        foreach (JsonNode? editNode in editsArray)
        {
            if (editNode is not JsonObject edit) continue;
            string? editFileArg = edit["file"]?.GetValue<string>();

            // Later edits often depend on earlier ones (e.g. a constant introduced by edit 1
            // and referenced by edit 3). After the first failure, skip the rest instead of
            // saving dependent edits around a missing prerequisite.
            if (failed > 0)
            {
                editResults.Add(new JsonObject
                {
                    ["file"] = editFileArg,
                    ["success"] = false,
                    ["skipped"] = true,
                    ["error"] = "Skipped: an earlier edit in this call failed.",
                });
                skipped++;
                continue;
            }

            JsonObject entry = await ApplySingleEditAsync(id, edit, editFileArg, bridge).ConfigureAwait(false);
            if (entry["success"]?.GetValue<bool>() == true) applied++; else failed++;
            editResults.Add(entry);
        }

        int totalEdits = applied + failed + skipped;
        string summary = failed == 0
            ? $"Applied {applied}/{totalEdits} edit(s)."
            : applied == 0
                ? $"Applied 0/{totalEdits} edit(s): the first edit failed and {skipped} later edit(s) were skipped. No files were modified by this call."
                : $"PARTIAL MUTATION: applied {applied}, failed 1, skipped {skipped} of {totalEdits} edit(s). " +
                  "The applied edits are already saved; the file may be inconsistent until the failed edit lands. " +
                  "Fix the failed edit now (see edits array) and resubmit ONLY the failed and skipped edits.";
        JsonObject combined = new()
        {
            ["Success"]  = failed == 0,
            ["Summary"]  = summary,
            ["Warnings"] = new JsonArray(),
            ["Data"]     = new JsonObject
            {
                ["count"]   = totalEdits,
                ["applied"] = applied,
                ["failed"]  = failed,
                ["skipped"] = skipped,
                ["partialMutation"] = failed > 0 && applied > 0,
                ["edits"]   = editResults,
            },
        };
        if (ArgBuilder.OptionalBool(args, PostCheck, false))
        {
            combined["postCheck"] = await RunDocumentPostCheckAsync(bridge, ApplyDiffTool).ConfigureAwait(false);
        }
        return combined;
    }

    private static async Task<JsonObject> ApplySingleEditAsync(
        JsonNode? id, JsonObject edit, string? editFileArg, BridgeConnection bridge)
    {
        ApplyDiffRequest editRequest;
        try
        {
            editRequest = ApplyDiffRequest.FromJsonObject(edit);
        }
        catch (ApplyDiffValidationException ex)
        {
            return new JsonObject
            {
                ["file"] = editFileArg,
                ["success"] = false,
                ["error"] = ex.Message,
            };
        }

        JsonObject editResult = await bridge.SendAsync(id, "apply-diff", BuildApplyDiffArgs(editRequest))
            .ConfigureAwait(false);
        bool success = editResult["Success"]?.GetValue<bool>() ?? false;
        bool alreadySatisfied = success &&
            editResult["Data"]?["items"] is JsonArray editItems &&
            editItems.Count > 0 &&
            editItems[0]?["alreadySatisfied"]?.GetValue<bool?>() == true;
        JsonObject entry = new()
        {
            ["file"] = editFileArg,
            ["success"] = success,
        };
        if (alreadySatisfied)
            entry["alreadySatisfied"] = true;
        if (editRequest.ReplaceAll)
            entry["replaceAll"] = true;
        if (!success)
            entry["error"] = editResult["Summary"]?.GetValue<string>();
        return entry;
    }

    private static IEnumerable<ToolEntry> ApplyDiffTools()
    {
        yield return new(
            ToolDefinitionCatalog.ApplyDiff(
                ObjectSchema(
                    Opt(FileArg, ApplyDiffFileDesc),
                    Opt("old_content",
                        "Exact source text block to replace. Copy the slice text from read_file without display line-number prefixes — whitespace-exact. " +
                        "Any mismatch causes a content-not-found error. " +
                        "Open files reload automatically before matching, so content is always current."),
                    Opt("new_content", "Replacement text to write in place of old_content."),
                    ("edits", new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] =
                            "Apply multiple targeted edits in one call. Each element is an independent edit object with its own " +
                            "file (handle), old_content, and new_content. Keep this array to 4 or fewer entries per call; " +
                            "split larger changes and re-check results between calls so Visual Studio is not overloaded. Each edit " +
                            "is dispatched as a separate command instance and applied in order; failures are reported per-entry " +
                            "without aborting the rest. Set replace_all on an individual edit only when every matching occurrence " +
                            "in that edit's explicit file should be replaced. For same-file edits, old_content is matched by " +
                            "content rather than original line number. Avoid duplicate old_content and overlapping replacements; " +
                            "entries already applied are not rolled back if a later entry fails.",
                        ["items"] = ObjectSchema(
                            Req(FileArg, "Handle (h:N, f:N) or plain path for the file to edit."),
                            Req("old_content", "Exact text block to replace — copy verbatim from read_file."),
                            Req("new_content", "Replacement text."),
                            OptBool("replace_all",
                                "When true, replace every non-overlapping occurrence of old_content, but only inside this " +
                                "edit item's explicit file or handle.")),
                        ["minItems"] = 1,
                    }, false),
                    Opt("diff",
                        "ONLY for multi-file or structural changes (add/move/delete files). " +
                        "For a single targeted edit omit this and use file + old_content + new_content instead. " +
                        "Single-file no-context replacement patches are rejected. " +
                        "Format for a multi-file content patch:\n" +
                        "*** Begin Patch\\n" +
                        "*** Update File: h:2\\n" +
                        "@@\\n" +
                        " context\\n" +
                        "-old line\\n" +
                        "+new line\\n" +
                        " context\\n" +
                        "*** Update File: f:3\\n" +
                        "@@\\n" +
                        " context\\n" +
                        "-old line\\n" +
                        "+new line\\n" +
                        " context\\n" +
                        "*** End Patch\n" +
                        "Structural patches may use *** Add File, *** Delete File, or *** Move to.\n" +
                        "Do not send unified diff headers like --- / +++."),
                    OptBool(PostCheck,
                        "Queue a quick diagnostics refresh after applying (default false)."),
                    OptBool("replace_all",
                        "When true, replace EVERY non-overlapping occurrence of old_content in the file instead of just the first. " +
                        "All replacements are grouped into a single VS undo transaction — one Ctrl+Z reverts all of them. " +
                        "Valid with a single file + old_content + new_content request, or inside an individual edits[] item. " +
                        "For batches, each replace_all applies only to that edit item's explicit file or handle; top-level " +
                        "replace_all is rejected with the edits array. Errors if old_content is not found at least once. " +
                        "IMPORTANT: after all edits are complete, call errors(), warnings(), and messages() with refresh=true " +
                        "for current diagnostics — do not leave diagnostics behind.")))
                .WithSearchHints(BuildSearchHints(
                    workflow:
                    [
                        ("errors", "Check current errors after edits; use refresh=true when you need a fresh UI read"),
                        ("warnings", "Check current warnings after edits; use refresh=true when you need a fresh UI read"),
                        ("messages", "Check current messages after edits; use refresh=true when you need a fresh UI read"),
                        ("reload_document", "Reload the file so VS picks up the changes"),
                    ],
                    related: [("write_file", "Overwrite the full file instead"), (ReadFileTool, "Read the file first to understand its current state")])),
            (id, args, bridge) => ExecuteApplyDiffAsync(id, args, bridge));
    }

    private static async Task<JsonNode> ExecuteApplyDiffAsync(JsonNode? id, JsonObject? args, BridgeConnection bridge)
    {
        // Multi-edit form: edits array — each element is its own command instance.
        if (args?["edits"] is JsonArray editsArray)
            return BridgeResult(await ApplyMultipleEditsAsync(id, args, editsArray, bridge).ConfigureAwait(false));

        // Single-edit form.
        ApplyDiffRequest request;
        try
        {
            request = ApplyDiffRequest.FromJsonObject(args);
        }
        catch (ApplyDiffValidationException ex)
        {
            throw new McpRequestException(id, McpErrorCodes.InvalidParams, ex.Message);
        }

        JsonObject result = await bridge.SendAsync(
            id,
            "apply-diff",
            BuildApplyDiffArgs(request))
            .ConfigureAwait(false);

        if (result["Data"] is JsonObject data)
        {
            data["validatedPatch"] = request.ToJsonObject();
        }

        if (ArgBuilder.OptionalBool(args, PostCheck, false))
        {
            result["postCheck"] = await RunDocumentPostCheckAsync(bridge, ApplyDiffTool).ConfigureAwait(false);
        }

        return BridgeResult(result);
    }

    private static IEnumerable<ToolEntry> WriteFileTools()
    {
        yield return new(
            ToolDefinitionCatalog.WriteFile(
                ObjectSchema(
                    Req(FileArg, FileDesc),
                    Req("content", "Full UTF-8 text content to write. This replaces the entire file; omitted text is removed."),
                    OptBool(PostCheck, "Queue a quick diagnostics refresh after writing (default false).")))
                .WithSearchHints(BuildSearchHints(
                    workflow: [("reload_document", "Reload the file so VS picks up the changes"), ("errors", "Check for diagnostics after writing")],
                    related: [(ApplyDiffTool, "Apply targeted changes instead of replacing the whole file"), (ReadFileTool, "Read the current file contents first")])),
            async (id, args, bridge) =>
            {
                JsonObject result = await bridge.SendAsync(id, "write-file", BuildWriteFileArgs(args))
                    .ConfigureAwait(false);

                if (ArgBuilder.OptionalBool(args, PostCheck, false))
                {
                    result["postCheck"] = await RunDocumentPostCheckAsync(bridge, "write_file").ConfigureAwait(false);
                }

                return BridgeResult(result);
            });
    }

    private static IEnumerable<ToolEntry> DocumentTabTools()
    {
        yield return BridgeTool("open_file",
            "Open a document by unique filename, solution-relative path, absolute path, or solution item name.",
            ObjectSchema(
                Req(FileArg, FileDesc),
                OptInt(Line, "Optional 1-based line number to navigate to."),
                OptInt(Column, "Optional 1-based column number.")),
            "open-document",
            a => Build(
                (FileArg, OptionalString(a, FileArg)),
                (Line, OptionalText(a, Line)),
                (Column, OptionalText(a, Column))),
            searchHints: BuildSearchHints(
                workflow: [(ReadFileTool, "Read the file contents"), ("file_outline", "Get the file symbol structure")],
                related: [("activate_document", "Switch to an already-open tab"), ("find_files", "Find the file path first")]));

        yield return BridgeTool("close_file",
            "Close one editor tab by exact FileArg path (preferred) or caption query. Use after list_tabs when trimming inactive tabs.",
            ObjectSchema(Opt(FileArg, "FileArg path to close."), Opt(Query, "Tab caption query.")),
            "close-file",
            a => Build(
                (FileArg, OptionalString(a, FileArg)),
                (Query, OptionalString(a, Query))),
            searchHints: BuildSearchHints(
                related: [("close_document", "Close by caption query"), ("close_others", "Close all except active"), (ListTabsTool, "List open tabs first")]));

        yield return BridgeTool("close_document",
            "Close editor tabs matching a caption/name query. Use to trim inactive tabs when list_tabs reports more than 7 open tabs. " +
            "Omit query with all: true to close every open tab.",
            ObjectSchema(Opt(Query, "Tab caption or name fragment. Omit with all: true to close all open tabs."), OptBool("all", "Close all matching tabs, or all open tabs when query is omitted.")),
            "close-document",
            a => Build(
                (Query, OptionalString(a, Query)),
                BoolArg("all", a, "all", false, true)),
            searchHints: BuildSearchHints(
                related: [("close_file", "Close by exact path"), ("close_others", "Close all except active"), (ListTabsTool, "List open tabs first")]));

        yield return BridgeTool("close_others",
            "Close all tabs except the active tab. Use this when list_tabs shows more than 7 tabs and only the active file is still needed.",
            ObjectSchema(OptBool("save", "Save before closing (default false).")),
            "close-others",
            a => Build(BoolArg("save", a, "save", false, true)),
            searchHints: BuildSearchHints(
                related: [("close_file", "Close a specific tab"), (ListTabsTool, "See what is open")]));

        yield return BridgeTool("save_document",
            "Save one document by path or handle, or save all open documents when file is omitted.",
            ObjectSchema(Opt(FileArg, "FileArg to save (path or handle). Omit to save all open documents.")),
            "save-document",
            a => OptionalString(a, FileArg) is { } file && !string.IsNullOrWhiteSpace(file)
                ? Build((FileArg, file))
                : Build(("all", "true")),
            searchHints: BuildSearchHints(
                workflow: [("errors", "Check diagnostics after saving")],
                related: [("reload_document", "Reload after external changes"), (ApplyDiffTool, "Apply changes before saving")]));

        yield return BridgeTool("reload_document",
            "Reload a document from disk — required after native Edit/Write tool changes. VS does not auto-detect external writes. " +
            "The file must already be open in an editor tab; call open_file first if it is not open, or check list_documents. Call " +
            "after every external edit, then check errors.",
            ObjectSchema(Req(FileArg, FileDesc)),
            "reload-document",
            a => Build((FileArg, OptionalString(a, FileArg))),
            searchHints: BuildSearchHints(
                workflow: [("errors", "Check for diagnostics after reloading")],
                related: [("save_document", "Save before reloading"), (ApplyDiffTool, "Apply changes that need reloading")]));

        yield return BridgeTool("list_documents",
            "List open documents.",
            EmptySchema(), "list-documents", _ => Empty(), Documents,
            searchHints: BuildSearchHints(
                workflow: [(ReadFileTool, "Read one of the listed documents"), ("activate_document", "Switch to a document")],
                related: [(ListTabsTool, "List open editor tabs")]));

        yield return BridgeTool(ListTabsTool,
            "List open editor tabs, identify the active tab, and report whether more than 7 tabs are open.",
            EmptySchema(), "list-tabs", _ => Empty(), Documents,
            searchHints: BuildSearchHints(
                workflow: [("close_file", "Close inactive tabs by path"), ("close_document", "Close matching inactive tabs"), ("close_others", "Keep only the active tab")],
                related: [("list_documents", "List open documents"), (ReadFileTool, "Read one of the listed files"), ("activate_document", "Switch to a tab")]));

        yield return BridgeTool("activate_document",
            "Activate an open document tab by query.",
            ObjectSchema(Req(Query, "FileArg name or tab caption fragment.")),
            "activate-document",
            a => Build((Query, OptionalString(a, Query))),
            searchHints: BuildSearchHints(
                workflow: [(ReadFileTool, "Read the activated file"), ("file_outline", "Get the file structure")],
                related: [("open_file", "Open a file that is not yet open"), (ListTabsTool, "List available tabs")]));
    }

    private static IEnumerable<ToolEntry> WindowCommandTools()
    {
        yield return BridgeTool("list_windows",
            "List Visual Studio tool windows (Solution Explorer, Error List, Output, etc.).",
            ObjectSchema(Opt(Query, "Optional caption filter.")),
            "list-windows",
            a => Build((Query, OptionalString(a, Query))),
            searchHints: BuildSearchHints(
                workflow: [("activate_window", "Bring a window to the foreground")],
                related: [("list_tabs", "List editor tabs")]));

        yield return BridgeTool("activate_window",
            "Bring a Visual Studio tool window to the foreground by caption fragment.",
            ObjectSchema(Req("window", "Window caption fragment.")),
            "activate-window",
            a => Build(("window", OptionalString(a, "window"))),
            searchHints: BuildSearchHints(
                related: [("list_windows", "List available windows"), ("execute_command", "Run a VS command")]));

        yield return BridgeTool("execute_command",
            "Execute an arbitrary Visual Studio command with optional arguments. Commands that may open or activate visible IDE " +
            "tool windows, such as TestExplorer.*, return visibleWindowSideEffect guidance so the model prompts the user before " +
            "closing or otherwise changing that window.",
            ObjectSchema(
                Req("command", "Visual Studio command name (e.g. Edit.FormatDocument)."),
                Opt("args", "Optional command arguments string."),
                Opt(FileArg, FileDesc),
                Opt("document", "Optional open-document query to position before running."),
                OptInt(Line, "Optional 1-based line number."),
                OptInt(Column, "Optional 1-based column number."),
                OptBool("select_word", "If true, select the word at the caret before executing.")),
            "execute-command",
            a => Build(
                ("command", OptionalString(a, "command")),
                ("args", OptionalString(a, "args")),
                (FileArg, OptionalString(a, FileArg)),
                ("document", OptionalString(a, "document")),
                (Line, OptionalText(a, Line)),
                (Column, OptionalText(a, Column)),
                BoolArg("select-word", a, "select_word", false, true)),
            searchHints: BuildSearchHints(
                related: [("format_document", "Format a document"), ("shell_exec", "Run an external process")]));

        yield return BridgeWrapperTool("format_document",
            "Format the current document or a specific FileArg.",
            ObjectSchema(
                Opt(FileArg, FileDesc),
                OptInt(Line, "Optional 1-based line."),
                OptInt(Column, "Optional 1-based column.")),
            "execute-command",
            a =>
            {
                string? file = OptionalString(a, FileArg);
                if (string.IsNullOrWhiteSpace(file))
                {
                    return Build(("name", "Edit.FormatDocument"));
                }

                return Build(
                    ("name", "Edit.FormatDocument"),
                    (FileArg, file),
                    (Line, OptionalText(a, Line)),
                    (Column, OptionalText(a, Column)));
            },
            searchHints: BuildSearchHints(
                workflow: [("errors", "Check diagnostics after formatting")],
                related: [("execute_command", "Run other VS commands"), ("apply_diff", "Make targeted edits")]));
    }

    private static IEnumerable<ToolEntry> FileOperationTools()
    {
        yield return new("delete_file",
            "Delete a FileArg from disk and close its editor tab. " +
            "SDK-style projects auto-update when a FileArg disappears from disk. " +
            "For legacy .csproj files use remove_file_from_project first.",
            ObjectSchema(
                Req(FileArg, "Absolute or solution-relative FileArg path to delete.")),
            Documents,
            async (id, args, bridge) =>
            {
                string fileArg = args?[FileArg]?.GetValue<string>() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(fileArg))
                    throw new McpRequestException(id, McpErrorCodes.InvalidParams, "Missing required argument 'FileArg'.");

                string solutionDir = ServiceToolPaths.ResolveSolutionDirectory(bridge);
                string resolvedPath = ResolveExistingFilePath(fileArg, solutionDir);

                // Close in editor first (best-effort)
                try
                {
                    await bridge.SendAsync(id, "close-file", Build((FileArg, resolvedPath)))
                        .ConfigureAwait(false);
                }
                catch (IOException ex)
                {
                    McpServerLog.WriteException($"failed to close '{resolvedPath}' before delete", ex);
                }
                catch (InvalidOperationException ex)
                {
                    McpServerLog.WriteException($"failed to close '{resolvedPath}' before delete", ex);
                }
                catch (McpRequestException ex)
                {
                    McpServerLog.WriteException($"failed to close '{resolvedPath}' before delete", ex);
                }

                System.IO.File.Delete(resolvedPath);

                JsonObject delPayload = new()
                {
                    ["deleted"] = true,
                    ["path"] = resolvedPath,
                };
                return ToolResultFormatter.StructuredToolResult(delPayload, args, successText: $"Deleted file '{resolvedPath}'.");
            },
            destructive: true,
            searchHints: BuildSearchHints(
                related: [("remove_file_from_project", "Remove from project first for legacy .csproj"), ("copy_file", "Copy instead of delete")]));

        yield return new("copy_file",
            "Copy a file to a new location on disk, creating parent directories as needed.",
            ObjectSchema(
                Req("source", "Absolute or solution-relative source path."),
                Req("destination", "Absolute or solution-relative destination path."),
                OptBool("overwrite", "Overwrite destination if it already exists (default false).")),
            Documents,
            async (id, args, bridge) =>
            {
                string sourceArg = args?["source"]?.GetValue<string>() ?? string.Empty;
                string destArg = args?["destination"]?.GetValue<string>() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(sourceArg))
                    throw new McpRequestException(id, McpErrorCodes.InvalidParams, "Missing required argument 'source'.");
                if (string.IsNullOrWhiteSpace(destArg))
                    throw new McpRequestException(id, McpErrorCodes.InvalidParams, "Missing required argument 'destination'.");

                bool overwrite = args?["overwrite"]?.GetValue<bool?>() ?? false;
                string solutionDir = ServiceToolPaths.ResolveSolutionDirectory(bridge);
                string resolvedSource = ResolveExistingFilePath(sourceArg, solutionDir);
                string resolvedDest = System.IO.Path.IsPathRooted(destArg) ? destArg : System.IO.Path.Combine(solutionDir, destArg);

                string? destDir = System.IO.Path.GetDirectoryName(resolvedDest);
                if (!string.IsNullOrEmpty(destDir))
                    Directory.CreateDirectory(destDir);

                System.IO.File.Copy(resolvedSource, resolvedDest, overwrite);

                JsonObject copyPayload = new()
                {
                    ["copied"] = true,
                    ["source"] = resolvedSource,
                    ["destination"] = resolvedDest,
                };
                return ToolResultFormatter.StructuredToolResult(copyPayload, args,
                    successText: $"Copied '{resolvedSource}' to '{resolvedDest}'.");
            },
            mutating: true,
            searchHints: BuildSearchHints(
                related: [("delete_file", "Delete the original after copying"), ("add_file_to_project", "Add the copied file to a project")]));
    }

    private static IEnumerable<ToolEntry> SolutionSystemTools()
    {
        yield return new("open_solution",
            "Open a specific existing .sln or .slnx file in the current Visual Studio instance.",
            ObjectSchema(
                Req("solution", "Absolute path to the .sln or .slnx file."),
                OptBool("wait_for_ready", "Wait for readiness after opening (default true).")),
            "documents",
            (id, args, bridge) => OpenSolutionAsync(id, args, bridge),
            searchHints: BuildSearchHints(
                workflow: [("wait_for_ready", "Wait for IntelliSense to load"), ("list_projects", "Inspect the loaded projects")],
                related: [("vs_open", "Launch a new VS instance"), ("bind_solution", "Bind to an already-open solution"), ("search_solutions", "Find the solution path")]));

        yield return BridgeTool("create_solution",
            "Create and open a new solution in the current Visual Studio instance.",
            ObjectSchema(
                Req("directory", "Absolute directory where the solution should be created."),
                Req("name", "Solution name ('.sln' is optional.)")),
            "create-solution",
            a => Build(
                ("directory", OptionalString(a, "directory")),
                ("name", OptionalString(a, "name"))),
            searchHints: BuildSearchHints(
                workflow: [("create_project", "Add a project to the new solution"), ("list_projects", "Inspect the solution structure")],
                related: [("open_solution", "Open an existing solution")]));

        yield return new("vs_close",
            "Close a Visual Studio instance by process id, or the currently bound instance.",
            ObjectSchema(
                OptInt("process_id", "Process ID of the VS instance to close. Defaults to bound instance."),
                OptBool("force", "Kill the process instead of gracefully closing (default false).")),
            "system",
            (id, args, bridge) => VsCloseAsync(id, args, bridge),
            searchHints: BuildSearchHints(
                related: [("vs_open", "Launch a VS instance"), ("bridge_health", "Check binding health")]));

        yield return new("run_tests",
            "Run .NET tests through dotnet test from the bound solution directory and return structured pass/fail counts plus output. " +
            "Use this for xUnit, NUnit, and MSTest projects before falling back to shell_exec. Supports project selection, " +
            "configuration, framework, runtime, VSTest filter expressions, loggers such as trx, results directories, and blame.",
            ObjectSchema(
                Opt("project", "Optional test project name or path. Omit to run the active solution with dotnet test."),
                Opt("configuration", "Optional build configuration passed to dotnet test (e.g. Release)."),
                Opt("framework", "Optional target framework passed with --framework."),
                Opt("runtime", "Optional runtime identifier passed with --runtime."),
                Opt("settings", "Optional .runsettings file passed with --settings."),
                Opt("filter", "Optional VSTest filter expression passed with --filter, e.g. FullyQualifiedName~MyTests."),
                Opt("logger", "Optional logger passed with --logger, e.g. trx or trx;LogFilePrefix=testResults."),
                Opt("results_directory", "Optional results directory passed with --results-directory."),
                Opt("collect", "Optional data collector passed with --collect, e.g. XPlat Code Coverage."),
                Opt("verbosity", "Optional dotnet test verbosity: quiet, minimal, normal, detailed, or diagnostic."),
                OptBool("no_restore", "Pass --no-restore (default true). Set false to allow restore."),
                OptBool("no_build", "Pass --no-build (default false)."),
                OptBool("blame", "Pass --blame to report tests in progress when the test host crashes (default false)."),
                OptInt("timeout_ms", "Timeout in milliseconds (default 120000)."),
                OptInt("head_lines", "If set, include only the first N lines of stdout and stderr."),
                OptInt("tail_lines", "If set, include only the last N lines of stdout and stderr. Combine with head_lines to see both ends of long output."),
                OptInt("max_lines", "Max total lines per stream when head_lines/tail_lines are not set (default 200).")),
            "test",
            (id, args, bridge) => RunTestsTool.ExecuteAsync(id, args, bridge),
            mutating: true,
            aliases: ["test", "dotnet_test", "unit_tests", "run_unit_tests", "vstest"],
            tags: ["test", "tests", "dotnet", "vstest", "xunit", "nunit", "mstest"],
            summary: "Run .NET tests through dotnet test and return structured results.",
            searchHints: BuildSearchHints(
                workflow: [("errors", "Check compile errors if tests fail before execution"), ("read_output", "Inspect VS output when a test run is linked to build output")],
                related: [("build_errors", "Build and return compiler errors"), ("shell_exec", "Run non-.NET test runners or custom scripts")]));

        yield return new("shell_exec",
            "Execute an arbitrary external process and capture stdout, stderr, and exit code. " +
            "Use for build scripts, package tools, non-.NET test runners, and CLI utilities. " +
            "Prefer named tools for common operations: git_* for version control, run_tests for .NET tests, " +
            "build / build_errors for compilation, delete_file / copy_file for FileArg operations. " +
            "Working directory defaults to the directory containing the .sln file. " +
            "For files in subdirectories use paths relative to that root (e.g. 'src/Foo/Bar.cs'), or pass cwd explicitly.",
            ObjectSchema(
                Req("exe", "Executable path or name (e.g. 'powershell', 'cmd', 'ISCC.exe')."),
                Opt("args", "Arguments string to pass to the executable."),
                Opt("cwd", "Working directory."),
                OptInt("timeout_ms", "Timeout in milliseconds (default 60000)."),
                OptInt("head_lines", "If set, include only the first N lines of stdout and stderr."),
                OptInt("tail_lines", "If set, include only the last N lines of stdout and stderr. Combine with head_lines to see both ends of long output."),
                OptInt("max_lines", "Max total lines per stream when head_lines/tail_lines are not set (default 200).")),
            "system",
            (id, args, bridge) => ShellExecTool.ExecuteAsync(id, args, bridge),
            aliases: ["bash", "shell", "run", "run_command", "run_shell_command", "terminal_command", "cmd", "powershell", "lint"],
            tags: ["shell", "bash", "run", "execute", "command", "test", "lint", "terminal"],
            summary: "Run shell commands, scripts, and lint/build helpers.",
            searchHints: BuildSearchHints(
                related: [("execute_command", "Run a VS command instead"), ("run_tests", "Run .NET tests"), ("build", "Use the build tool for compilation"), ("git_status", "Use git tools for version control")]));

    }

    private static async Task<JsonNode> OpenSolutionAsync(JsonNode? id, JsonObject? args, BridgeConnection bridge)
    {
        string solutionPath = System.IO.Path.GetFullPath(OptionalString(args, "solution")
            ?? throw new McpRequestException(id, McpErrorCodes.InvalidParams, "'solution' is required."));

        if (!File.Exists(solutionPath))
        {
            throw new McpRequestException(id, McpErrorCodes.InvalidParams, $"Solution file not found: {solutionPath}");
        }

        string extension = System.IO.Path.GetExtension(solutionPath);
        if (!string.Equals(extension, ".sln", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(extension, ".slnx", StringComparison.OrdinalIgnoreCase))
        {
            throw new McpRequestException(id, McpErrorCodes.InvalidParams, $"File is not a solution file: {solutionPath}");
        }

        IReadOnlyList<BridgeInstance> instances = await VsDiscovery.ListAsync(bridge.Mode).ConfigureAwait(false);
        BridgeInstance[] matchingInstances = [.. instances.Where(instance =>
            !string.IsNullOrWhiteSpace(instance.SolutionPath)
            && string.Equals(System.IO.Path.GetFullPath(instance.SolutionPath), solutionPath, StringComparison.OrdinalIgnoreCase))];

        if (matchingInstances.Length == 1)
        {
            JsonObject bindResult = await bridge.BindAsync(id, new JsonObject { ["instance_id"] = matchingInstances[0].InstanceId }).ConfigureAwait(false);
            bindResult["alreadyOpen"] = true;
            bindResult["solutionPath"] = solutionPath;

            if (OptionalBool(args, "wait_for_ready", true))
            {
                JsonObject ready = await bridge.SendAsync(id, "ready", []).ConfigureAwait(false);
                bindResult["ready"] = ready["Data"]?.DeepClone();
            }

            return ToolResultFormatter.StructuredToolResult(bindResult, args, successText: "Bound to already-open solution.");
        }

        if (matchingInstances.Length > 1)
        {
            throw new McpRequestException(id, McpErrorCodes.BridgeError,
                $"Multiple live VS IDE Bridge instances already have '{solutionPath}' open. Call bind_instance with a specific instance_id.");
        }

        if (bridge.CurrentInstance is null)
        {
            throw new McpRequestException(id, McpErrorCodes.BridgeError,
                "No live bound VS IDE Bridge instance is available to open the solution. Open Visual Studio with the VS IDE Bridge extension installed, then call bind_solution or bind_instance first.");
        }

        string openArgs = Build(
            ("solution", solutionPath),
            BoolArg("wait-for-ready", args, "wait_for_ready", true, true));
        JsonObject response = await bridge.SendIgnoringSolutionHintAsync(id, "open-solution", openArgs).ConfigureAwait(false);
        return BridgeResult(response);
    }

    private static async Task<JsonObject> RunDocumentPostCheckAsync(BridgeConnection bridge, string sourceTool)
    {
        JsonObject snapshot = await bridge.DocumentDiagnostics
            .QueueRefreshAndWaitForSnapshotAsync(sourceTool, clearCached: true)
            .ConfigureAwait(false);
        string status = snapshot["status"]?.GetValue<string>() ?? "unknown";
        JsonObject? errorsData = snapshot["errors"]?["Data"]?.AsObject();
        JsonObject? warningsData = snapshot["warnings"]?["Data"]?.AsObject();
        JsonObject? messagesData = snapshot["messages"]?["Data"]?.AsObject();
        bool completed = string.Equals(status, "completed", StringComparison.OrdinalIgnoreCase);
        bool hasCompleteCounts = completed
            && errorsData is not null
            && warningsData is not null
            && messagesData is not null;
        if (!hasCompleteCounts)
        {
            bool failed = string.Equals(status, "failed", StringComparison.OrdinalIgnoreCase);
            JsonObject pending = new()
            {
                ["mode"] = "refresh-and-wait",
                ["status"] = status,
                ["pending"] = !completed && !failed,
                ["summary"] = failed
                    ? "Diagnostics refresh failed; run errors, warnings, or messages with refresh=true for current counts."
                    : "Diagnostics refresh did not produce complete counts; run errors, warnings, or messages with refresh=true for current counts.",
                ["snapshot"] = snapshot.DeepClone(),
            };
            if (snapshot["lastError"] is JsonNode lastError)
            {
                pending["lastError"] = lastError.DeepClone();
            }

            return pending;
        }

        int errorCount   = GetSeverityCount(errorsData, "Error");
        int warningCount = GetSeverityCount(errorsData, "Warning");
        int messageCount = GetSeverityCount(errorsData, "Message");
        bool hasErrors = errorCount > 0 || (errorsData?["hasErrors"]?.GetValue<bool>() ?? false);
        string counts = $"{errorCount} error(s), {warningCount} warning(s), {messageCount} message(s)";
        // Warnings/messages don't block the build; only errors do. Still surface them so the
        // model doesn't declare victory while cleanup targets remain.
        string summary = hasErrors
            ? $"{counts} - fix errors before building."
            : warningCount + messageCount > 0
                ? $"{counts} - no errors; the build can succeed, but warnings/messages remain to fix."
                : $"{counts} - clean.";
        JsonObject result = new()
        {
            ["mode"] = "refresh-and-wait",
            ["status"] = status,
            ["pending"] = false,
            ["hasErrors"] = hasErrors,
            ["errorCount"] = errorCount,
            ["warningCount"] = warningCount,
            ["messageCount"] = messageCount,
            ["summary"] = summary,
        };
        if (hasErrors)
        {
            result["errors"] = errorsData?["rows"]?.DeepClone();
        }

        return result;
    }

    private static int GetSeverityCount(JsonObject? diagnosticsData, string severity)
        => diagnosticsData?["totalSeverityCounts"]?[severity]?.GetValue<int>() ?? 0;

}
