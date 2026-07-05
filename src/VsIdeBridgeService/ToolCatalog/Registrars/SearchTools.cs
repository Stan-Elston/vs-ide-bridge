using System.Text.Json.Nodes;
using static VsIdeBridgeService.ArgBuilder;
using static VsIdeBridgeService.SchemaHelpers;
using VsIdeBridge.Shared;
using VsIdeBridge.Tooling.Documents;
using VsIdeBridge.Tooling.Search;

namespace VsIdeBridgeService;

internal static partial class ToolCatalog
{
    private const string SearchDescriptionProperty = "description";
    private const string SearchExtensionsProperty = "extensions";
    private const string SearchFileOutlineTool = "file_outline";
    private const string SearchFindTextTool = "find_text";
    private const string MatchCase = "match_case";
    private const string SearchReadFileTool = "read_file";
    private const string SearchSymbolsTool = "search_symbols";
    private const string SearchReadFileBatchTool = "read_file_batch";
    private const int DefaultReadChunkSize = 10;
    private const int DefaultSearchChunkSize = 25;
    private const string ReadChunkSizeDescription = "Slices per returned chunk (default 10). Set 0 to return all filtered slices.";
    private const string ReadChunkIndexDescription = "Zero-based slice chunk index to return (default 0).";
    private const string ReadSortByDescription = "Optional slice sort field: path/file, name, requestedStartLine, requestedEndLine, actualStartLine, actualEndLine, lineCount, text, revealNote, or index (default index — preserves request order).";
    private const string ReadSortDirectionDescription = "Optional sort direction: asc or desc (default asc).";
    private const string ReadPathFilterDescription = "Optional path filter applied to returned slices.";
    private const string ReadTextFilterDescription = "Optional text filter applied to slice text, path, filename, and reveal note.";
    private const string ReadGroupByDescription = "Optional grouping mode: path/file or revealed.";
    private const string SearchChunkSizeDescription = "Rows per returned chunk (default 25). Set 0 to return all filtered rows.";
    private const string SearchChunkIndexDescription = "Zero-based result chunk index to return (default 0).";
    private const string SearchSortByDescription = "Optional result sort field: path/file, name, project, kind, source, line, column, score, text, preview, message, or signature.";
    private const string SearchSortDirectionDescription = "Optional sort direction: asc or desc (default asc).";
    private const string SearchTextFilterDescription = "Optional post-search text filter applied to names, paths, preview text, messages, and signatures.";
    private const string SearchSourceFilterDescription = "Optional result source filter.";
    private const string SearchGroupByDescription = "Optional grouping mode: path/file, name, project, kind, or source.";

    private static IEnumerable<ToolEntry> SearchTools()
        =>
        FileReadTools()
            .Concat(TextSearchTools())
            .Concat(CodeExplorationTools());

    private static IEnumerable<ToolEntry> FileReadTools()
    {
        yield return BridgeTool(
            ToolDefinitionCatalog.ReadFile(
                ObjectSchema(
                    Req("file", FileDesc),
                    OptInt("start_line", "First 1-based line to read. Omit to start at line 1. To continue past a previous slice: set to (last line shown + 1)."),
                    OptInt("end_line", "Last 1-based line (inclusive). Always a windowed range — there is no full-file mode."),
                    OptInt("line", "Anchor 1-based line. Use with context_before/context_after."),
                    OptInt("context_before", "Lines before anchor (default 10)."),
                    OptInt("context_after", "Lines after anchor (default 30)."),
                    OptBool("reveal_in_editor", "Reveal slice in editor (default true).")))
                .WithSearchHints(BuildSearchHints(
                    workflow: [("apply_diff", "Apply targeted edits to the file"), ("file_outline", "Get symbol structure of the file")],
                     related: 
                     [
                         ("read_file_batch", "Read multiple slices in one call — per range: start_line+end_line for a fixed range, or line+context_before+context_after for an anchor"),
                         ("file_outline", "Get symbol line numbers first, then re-call read_file with start_line/end_line to read only that slice"),
                         ("find_text", "Search for text to locate the right line before slicing"),
                     ])),
            "document-slice",
            a => Build(
                ("file", OptionalString(a, "file")),
                ("start-line", OptionalText(a, "start_line")),
                ("end-line", OptionalText(a, "end_line")),
                (Line, OptionalText(a, Line)),
                ("context-before", OptionalText(a, "context_before")),
                ("context-after", OptionalText(a, "context_after")),
                    BoolArg("reveal-in-editor", a, "reveal_in_editor", true, true)),
            transformResponse: NormalizeReadSliceResponse);

        yield return new(
            ToolDefinitionCatalog.ReadFileBatch(
                ObjectSchema(
                    WithReadSliceOptions(
                        (("ranges",
                            new JsonObject
                            {
                                ["type"] = "array",
                                ["description"] = "Ranges to read in order.",
                                ["items"] = ObjectSchema(
                                    Req("file", FileDesc),
                                    OptInt("start_line", "First 1-based line."),
                                    OptInt("end_line", "Last 1-based line."),
                                    OptInt("line", "Anchor line."),
                                    OptInt("context_before", "Lines before anchor."),
                                    OptInt("context_after", "Lines after anchor.")),
                                ["minItems"] = 1,
                            },
                            true)))))
                .WithSearchHints(BuildSearchHints(
                    workflow: [("apply_diff", "Apply changes after reading multiple files")],
                    related:
                    [
                        ("read_file", "Read a single slice with start_line/end_line"),
                        ("file_outline", "Get line numbers for each symbol first, then batch the slices here"),
                        ("find_text_batch", "Search multiple patterns to discover line numbers before batching"),
                    ])),
            async (id, args, bridge) =>
            {
                JsonArray ranges = args?["ranges"]?.AsArray() ?? [];
                JsonArray slices = [];
                foreach (JsonNode? rangeNode in ranges)
                {
                    if (rangeNode is not JsonObject range) continue;
                    string sliceArgs = Build(
                        ("file",             OptionalString(range, "file")),
                        ("start-line",       OptionalText(range, "start_line")),
                        ("end-line",         OptionalText(range, "end_line")),
                        (Line,               OptionalText(range, "line")),
                        ("context-before",   OptionalText(range, "context_before")),
                        ("context-after",    OptionalText(range, "context_after")),
                        ("reveal-in-editor", "false"));
                    JsonObject sliceResponse = await bridge.SendAsync(id, "document-slice", sliceArgs)
                        .ConfigureAwait(false);
                    NormalizeReadSliceResponse(sliceResponse, null);
                    if (sliceResponse["Data"] is JsonObject sliceData)
                        slices.Add(sliceData.DeepClone());
                    else
                        slices.Add(new JsonObject { ["error"] = sliceResponse["Summary"]?.GetValue<string>() ?? "Unknown error reading slice." });
                }
                JsonObject combined = new()
                {
                    ["Success"]  = true,
                    ["Summary"]  = $"Captured {slices.Count} slice(s).",
                    ["Warnings"] = new JsonArray(),
                    ["Data"]     = new JsonObject
                    {
                        ["count"]  = slices.Count,
                        ["slices"] = slices,
                    },
                };
                return BridgeResult(CompactReadSlicesResponse(combined, args));
            });
    }

    private static (string Name, JsonObject Schema, bool Required)[] WithReadSliceOptions(
        params (string Name, JsonObject Schema, bool Required)[] properties)
    {
        List<(string Name, JsonObject Schema, bool Required)> result = [.. properties];
        result.Add(OptInt("chunk_size", ReadChunkSizeDescription));
        result.Add(OptInt("chunk_index", ReadChunkIndexDescription));
        result.Add(Opt("sort_by", ReadSortByDescription));
        result.Add(Opt("sort_direction", ReadSortDirectionDescription));
        result.Add(Opt("path", ReadPathFilterDescription));
        result.Add(Opt("text", ReadTextFilterDescription));
        result.Add(Opt("group_by", ReadGroupByDescription));
        return [.. result];
    }

    private static JsonObject NormalizeReadSliceResponse(JsonObject response, JsonObject? args)
    {
        bool success = (response["Success"] ?? response["success"])?.GetValue<bool>() ?? false;
        if (success && response["Data"] is JsonObject data)
        {
            // Strip the lines array (each line as a {line, text} object) — it duplicates
            // the preformatted text field and balloons the structured content for large slices.
            // Also remove queue (pipe scheduling metadata), which is never useful to callers.
            data.Remove("lines");
            data.Remove("queue");
        }

        return response;
    }

    private static JsonObject CompactReadSlicesResponse(JsonObject response, JsonObject? args)
    {
        bool success = (response["Success"] ?? response["success"])?.GetValue<bool>() ?? false;
        if (!success || response["Data"] is not JsonObject data || !ReadSliceCollection.TryFromJsonObject(data, out ReadSliceCollection collection))
        {
            return response;
        }

        ReadQueryOptions options = ReadQueryOptions.FromJsonObject(args, DefaultReadChunkSize);
        response["Data"] = collection.ToJsonObject(options, data);

        // Replace the compact "Captured N slice(s)." summary with the actual slice content
        // so Claude can read the code directly from the tool result text.
        response["Summary"] = collection.BuildPageSummary(options);

        return response;
    }

    private static IEnumerable<ToolEntry> TextSearchTools()
        =>
        FileDiscoveryTools()
            .Concat(TextPatternTools())
            .Concat(SymbolSearchTools());

    private static (string Name, JsonObject Schema, bool Required)[] WithSearchResultOptions(
        params (string Name, JsonObject Schema, bool Required)[] properties)
    {
        List<(string Name, JsonObject Schema, bool Required)> result = [.. properties];
        result.Add(OptInt("chunk_size", SearchChunkSizeDescription));
        result.Add(OptInt("chunk_index", SearchChunkIndexDescription));
        result.Add(Opt("sort_by", SearchSortByDescription));
        result.Add(Opt("sort_direction", SearchSortDirectionDescription));
        result.Add(Opt("text", SearchTextFilterDescription));
        result.Add(Opt("source", SearchSourceFilterDescription));
        result.Add(Opt("group_by", SearchGroupByDescription));
        return [.. result];
    }

    private static JsonObject CompactSearchResponse(
        JsonObject response,
        JsonObject? args,
        bool includePathFilter,
        params string[] itemProperties)
    {
        bool success = (response["Success"] ?? response["success"])?.GetValue<bool>() ?? false;
        if (!success || response["Data"] is not JsonObject data)
        {
            return response;
        }

        SearchQueryOptions options = SearchQueryOptions.FromJsonObject(args, DefaultSearchChunkSize, includePathFilter);
        foreach (string itemProperty in itemProperties)
        {
            ApplySearchCollection(data, itemProperty, options);
        }

        CompactNestedSearchCollections(data, options);
        return response;
    }

    private static void CompactSearchPayload(
        JsonObject payload,
        JsonObject? args,
        bool includePathFilter,
        params string[] itemProperties)
    {
        SearchQueryOptions options = SearchQueryOptions.FromJsonObject(args, DefaultSearchChunkSize, includePathFilter);
        foreach (string itemProperty in itemProperties)
        {
            ApplySearchCollection(payload, itemProperty, options);
        }

        CompactNestedSearchCollections(payload, options);
    }

    private static void ApplySearchCollection(JsonObject data, string itemProperty, SearchQueryOptions options)
    {
        if (data[itemProperty] is not JsonArray)
        {
            return;
        }

        SearchResultCollection.FromJsonObject(data, itemProperty).WriteTo(data, itemProperty, options);
    }

    private static void CompactNestedSearchCollections(JsonObject data, SearchQueryOptions options)
    {
        if (data["queryResults"] is JsonArray queryResults)
        {
            foreach (JsonObject queryResult in queryResults.OfType<JsonObject>())
            {
                ApplySearchCollection(queryResult, SearchJsonNames.Matches, options);
            }
        }

        foreach (string groupProperty in new[] { SearchJsonNames.Results, SearchJsonNames.Contexts })
        {
            if (data[groupProperty] is not JsonArray groups)
            {
                continue;
            }

            foreach (JsonObject group in groups.OfType<JsonObject>())
            {
                ApplySearchCollection(group, SearchJsonNames.Matches, options);
                ApplySearchCollection(group, SearchJsonNames.Symbols, options);
                ApplySearchCollection(group, SearchJsonNames.Files, options);
            }
        }
    }

    private static IEnumerable<ToolEntry> FileDiscoveryTools()
    {
        yield return new(
            ToolDefinitionCatalog.FindFiles(
                ObjectSchema(WithSearchResultOptions(
                    Req(Query, "File name or path fragment. Use instead of Glob/find/ls when you know the filename or a partial path. For glob patterns like **/*.cs use the glob tool instead."),
                    Opt(Path, "Optional path fragment filter."),
                    Opt(Project, "Optional project name or path post-filter."),
                    OptArr(SearchExtensionsProperty, "Optional extension filters like ['.cmake','.txt']."),
                    OptBool("include_non_project", "Include disk files under solution root that are not in projects (default true)."),
                    OptInt("max_results", "Optional max result count (default 200)."),
                    Opt(Scope, "Optional scope: solution (default), project, document, or open."))))
                .WithSearchHints(BuildSearchHints(
                    workflow: [(SearchReadFileTool, "Read the found file"), (SearchReadFileBatchTool, "Read multiple found files at once")],
                    related: [("glob", "Find files by glob pattern"), (SearchFindTextTool, "Search file contents"), (SearchSymbolsTool, "Search by symbol name")]))
                .WithOutputSchema(BuildFindFilesOutputSchema()),
            async (id, args, bridge) =>
            {
                JsonObject response = await bridge.SendAsync(id, "find-files", Build(
                    (Query, OptionalString(args, Query)),
                    (Path, OptionalString(args, Path)),
                    (SearchExtensionsProperty, OptionalStringArray(args, SearchExtensionsProperty)),
                    BoolArg("include-non-project", args, "include_non_project", true, true),
                    ("max-results", OptionalText(args, "max_results")),
                    (Scope, OptionalString(args, Scope)))).ConfigureAwait(false);

                bool success = response["Success"]?.GetValue<bool>() ?? false;
                if (!success)
                {
                    return ToolResultFormatter.StructuredToolResult(response, args, isError: true);
                }

                response = CompactSearchResponse(response, args, includePathFilter: true, SearchJsonNames.Matches);
                JsonObject data = response["Data"] as JsonObject ?? [];
                int count = data["count"]?.GetValue<int>() ?? 0;
                int filteredCount = data[SearchJsonNames.FilteredCount]?.GetValue<int>() ?? count;
                int totalCount = data[SearchJsonNames.TotalCount]?.GetValue<int>() ?? filteredCount;
                string query = data[Query]?.GetValue<string>() ?? OptionalString(args, Query) ?? string.Empty;
                const string SearchToolsHint = " Tip: find_text and search_symbols also return h: handles you can pass directly to apply_diff. Call list_tools_by_category({\"category\":\"search\"}) to see all search tools.";
                string text = totalCount == 0 || filteredCount == 0
                    ? $"find_files: no file(s) found for '{query}'. If you expected a pattern match, try glob instead.{SearchToolsHint}"
                    : (ToolResultFormatter.StructuredToolResult(response, args).AsObject()["content"]?[0]?["text"]?.GetValue<string>()
                        ?? $"find_files: found {count} file(s) in this chunk.") + SearchToolsHint;

                return new JsonObject
                {
                    ["content"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["type"] = "text",
                            ["text"] = text,
                        },
                    },
                    ["isError"] = false,
                    ["structuredContent"] = StripBridgeEnvelopeFields(data),
                };
            });
    }

    private static IEnumerable<ToolEntry> TextPatternTools()
    {
        yield return new ToolEntry(
            ToolDefinitionCatalog.FindText(
                ObjectSchema(WithSearchResultOptions(
                    Req(Query, "Search text or regex pattern."),
                    Opt(Path, "Optional path or directory filter."),
                    Opt(Scope, "Scope: solution (default), project, or document."),
                    Opt(Project, ProjectFilterDesc),
                    OptBool(MatchCase, "Case-sensitive match (default false)."),
                    OptBool("whole_word", "Match whole word only (default false)."),
                    OptBool("regex", "Treat query as a regular expression (default false)."))))
                .WithSearchHints(BuildSearchHints(
                    workflow: [(SearchReadFileTool, "Read the file at the matched location"), ("goto_definition", "Navigate to the definition of the match")],
                    related: [("find_text_batch", "Search multiple patterns at once"), (SearchSymbolsTool, "Search by symbol name")])),
            async (id, args, bridge) =>
            {
                if (TryGetDiagnosticSearchCode(args, out string code))
                {
                    return CreateDiagnosticSearchRedirect(SearchFindTextTool, args, code);
                }

                JsonObject response = await bridge.SendAsync(id, "find-text", Build(
                    (Query, OptionalString(args, Query)),
                    (Path, OptionalString(args, Path)),
                    (Scope, OptionalString(args, Scope)),
                    (Project, OptionalString(args, Project)),
                    BoolArg("match-case", args, MatchCase, false, true),
                    BoolArg("whole-word", args, "whole_word", false, true),
                    BoolArg("regex", args, "regex", false, true))).ConfigureAwait(false);
                response = CompactSearchResponse(response, args, includePathFilter: true, SearchJsonNames.Matches);
                return BridgeResult(response, args);
            });

        yield return new ToolEntry(
            ToolDefinitionCatalog.FindTextBatch(
                ObjectSchema(WithSearchResultOptions(
                    (("queries",
                        new JsonObject
                        {
                            ["type"] = "array",
                            ["description"] = "Queries to search for in order. Each item is a string pattern or an object with a 'query' field.",
                            ["items"] = BuildBatchQueryItemSchema(),
                            ["minItems"] = 1,
                        },
                        true)),
                    Opt(Scope, "Optional scope: solution, project, document, or open."),
                    Opt(Project, ProjectFilterDesc),
                    Opt(Path, "Optional path filter."),
                    OptInt("results_window", "Optional Find Results window number."),
                    OptInt("max_queries_per_chunk", "Optional max query count per chunk (default 5)."),
                    OptBool(MatchCase, "Case-sensitive match (default false)."),
                    OptBool("whole_word", "Match whole word only (default false)."),
                    OptBool("regex", "Treat queries as regular expressions (default false)."))))
                .WithSearchHints(BuildSearchHints(
                    workflow: [(SearchReadFileTool, "Read the files at matched locations"), (SearchReadFileBatchTool, "Read multiple matched files at once")],
                    related: [(SearchFindTextTool, "Search a single pattern"), (SearchSymbolsTool, "Search by symbol name")])),
            async (id, args, bridge) =>
            {
                if (TryGetDiagnosticSearchCodeFromBatch(args, out string code, out string query))
                {
                    return CreateDiagnosticSearchRedirect("find_text_batch", args, code, query);
                }

                JsonObject response = await bridge.SendAsync(id, "find-text-batch", Build(
                    ("queries", NormalizeBatchQueryList(args?["queries"])),
                    (Scope, OptionalString(args, Scope)),
                    (Project, OptionalString(args, Project)),
                    (Path, OptionalString(args, Path)),
                    ("results-window", OptionalText(args, "results_window")),
                    ("max-queries-per-chunk", OptionalText(args, "max_queries_per_chunk")),
                    BoolArg("match-case", args, MatchCase, false, true),
                    BoolArg("whole-word", args, "whole_word", false, true),
                    BoolArg("regex", args, "regex", false, true))).ConfigureAwait(false);
                response = CompactSearchResponse(response, args, includePathFilter: true, SearchJsonNames.Matches);
                return BridgeResult(response, args);
            });
    }

    private static JsonObject BuildBatchQueryItemSchema() =>
        new()
        {
            ["oneOf"] = new JsonArray
            {
                new JsonObject { ["type"] = "string" },
                new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = new JsonObject
                    {
                        ["query"] = new JsonObject { ["type"] = "string", ["description"] = "The search pattern." },
                        ["path"] = new JsonObject { ["type"] = "string", ["description"] = "Optional path hint (informational; use the top-level path parameter to filter results)." },
                    },
                    ["required"] = new JsonArray { "query" },
                },
            }
        };

    private static string? NormalizeBatchQueryList(JsonNode? node)
    {
        if (node is not JsonArray arr)
        {
            return null;
        }

        // Accept both string[] and {query, path}[] — normalize objects to extract just the query string.
        JsonArray normalized = [];
        foreach (JsonNode? item in arr)
        {
            if (item is JsonValue v && v.TryGetValue<string>(out string? s))
            {
                normalized.Add(s);
            }
            else if (item is JsonObject obj && obj.TryGetPropertyValue("query", out JsonNode? q) && q is not null)
            {
                normalized.Add(q.GetValue<string>());
            }
        }

        return normalized.Count > 0 ? normalized.ToJsonString() : null;
    }

    internal static bool TryGetDiagnosticSearchCode(JsonObject? args, out string code)
    {
        return TryGetDiagnosticSearchCode(
            OptionalString(args, Query),
            OptionalString(args, Path),
            OptionalString(args, Project),
            OptionalString(args, Scope),
            out code);
    }

    internal static bool TryGetDiagnosticSearchCodeFromBatch(JsonObject? args, out string code, out string query)
    {
        code = string.Empty;
        query = string.Empty;

        if (args?["queries"] is not JsonArray queries)
        {
            return false;
        }

        // Query items may be plain strings or {query, path} objects (see NormalizeBatchQueryList);
        // calling GetValue<string>() on the object form throws before the batch even runs.
        foreach (string? item in queries.Select(GetBatchQueryText))
        {
            if (TryGetDiagnosticSearchCode(
                item,
                OptionalString(args, Path),
                OptionalString(args, Project),
                OptionalString(args, Scope),
                out code))
            {
                query = item ?? string.Empty;
                return true;
            }
        }

        return false;
    }

    private static string? GetBatchQueryText(JsonNode? node) => node switch
    {
        JsonObject obj when obj.TryGetPropertyValue("query", out JsonNode? q)
            && q is JsonValue qv && qv.TryGetValue(out string? objQuery) => objQuery,
        JsonValue value when value.TryGetValue(out string? text) => text,
        _ => null,
    };

    private static bool TryGetDiagnosticSearchCode(
        string? query,
        string? path,
        string? project,
        string? scope,
        out string code)
    {
        code = string.Empty;

        if (string.IsNullOrWhiteSpace(query)
            || !string.IsNullOrWhiteSpace(path)
            || !string.IsNullOrWhiteSpace(project))
        {
            return false;
        }

        string resolvedScope = scope ?? "solution";
        if (resolvedScope.Equals("document", StringComparison.OrdinalIgnoreCase)
            || resolvedScope.Equals("open", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return TryExtractLeadingDiagnosticCode(query, out code)
            && LooksLikeDiagnosticRowMessage(query);
    }

    private static JsonNode CreateDiagnosticSearchRedirect(string toolName, JsonObject? args, string code, string? matchedQuery = null)
    {
        string query = matchedQuery ?? OptionalString(args, Query) ?? string.Empty;
        string message = $"{toolName} searches source files, not Visual Studio Error List rows. Use warnings, errors, or messages with code '{code}' and a text filter instead of searching the solution text for '{query}'.";
        JsonObject response = new()
        {
            ["SchemaVersion"] = 1,
            ["Command"] = toolName,
            ["Success"] = false,
            ["Summary"] = "Rejected diagnostic-shaped text search.",
            ["Warnings"] = new JsonArray(),
            ["Error"] = new JsonObject
            {
                ["code"] = "invalid_arguments",
                ["message"] = message,
            },
            ["Data"] = new JsonObject
            {
                [Query] = query,
                ["diagnosticCode"] = code,
                ["suggestedTools"] = new JsonArray("warnings", "errors", "messages"),
            },
        };

        return ToolResultFormatter.StructuredToolResult(response, args, isError: true);
    }

    private static bool TryExtractLeadingDiagnosticCode(string query, out string code)
    {
        code = string.Empty;
        string trimmed = query.TrimStart();
        int end = 0;

        while (end < trimmed.Length && !char.IsWhiteSpace(trimmed[end]) && trimmed[end] != ':')
        {
            end++;
        }

        if (end == 0)
        {
            return false;
        }

        code = trimmed[..end];
        return IsDiagnosticCodeToken(code);
    }

    private static bool IsDiagnosticCodeToken(string code)
    {
        if (code.Length > 16)
        {
            return false;
        }

        bool hasLetter = false;
        bool hasDigit = false;
        foreach (char value in code)
        {
            if (char.IsLetter(value))
            {
                hasLetter = true;
                continue;
            }

            if (char.IsDigit(value))
            {
                hasDigit = true;
                continue;
            }

            if (value != '-')
            {
                return false;
            }
        }

        return hasLetter && hasDigit;
    }

    private static bool LooksLikeDiagnosticRowMessage(string query)
    {
        return query.Contains(".*", StringComparison.Ordinal)
            || query.Contains(" appears ", StringComparison.OrdinalIgnoreCase)
            || query.Contains(" warning", StringComparison.OrdinalIgnoreCase)
            || query.Contains(" error", StringComparison.OrdinalIgnoreCase)
            || query.Contains(" message", StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<ToolEntry> SymbolSearchTools()
    {
        yield return BridgeTool(
            ToolDefinitionCatalog.SearchSymbols(
                ObjectSchema(WithSearchResultOptions(
                    Req(Query, "Symbol search text."),
                    Opt("kind", "Optional symbol kind filter."),
                    Opt(Scope, "Optional scope: solution, project, document, or open."),
                    Opt(Project, ProjectFilterDesc),
                    Opt(Path, "Optional path or directory filter."),
                    OptInt(Max, "Optional max result count."),
                    OptBool(MatchCase, "Case-sensitive match (default false)."))))
                .WithSearchHints(BuildSearchHints(
                    workflow: [("goto_definition", "Navigate to the found symbol"), (SearchReadFileTool, "Read the file containing the symbol"), ("find_references", "Find all usages of the symbol")],
                    related: [(SearchFileOutlineTool, "Get all symbols in a known file"), (SearchFindTextTool, "Search by text instead of symbol name")])),
            "search-symbols",
            a => Build(
                (Query, OptionalString(a, Query)),
                ("kind", OptionalString(a, "kind")),
                (Scope, OptionalString(a, Scope)),
                (Project, OptionalString(a, Project)),
                (Path, OptionalString(a, Path)),
                (Max, OptionalText(a, Max)),
                BoolArg("match-case", a, MatchCase, false, true)),
            transformResponse: (response, args) => CompactSearchResponse(response, args, includePathFilter: true, SearchJsonNames.Matches));
    }

    private static JsonObject BuildFindFilesOutputSchema()
        => new()
        {
            ["type"] = "object",
            ["properties"] = new JsonObject
            {
                ["query"] = new JsonObject { ["type"] = "string", [SearchDescriptionProperty] = "The file name or path fragment search query." },
                ["pathFilter"] = new JsonObject { ["type"] = "string", [SearchDescriptionProperty] = "Optional path fragment filter applied." },
                [SearchExtensionsProperty] = new JsonObject
                {
                    ["type"] = "array",
                    [SearchDescriptionProperty] = "File extension filters applied.",
                    ["items"] = new JsonObject { ["type"] = "string" },
                },
                ["includeNonProject"] = new JsonObject { ["type"] = "boolean", [SearchDescriptionProperty] = "Whether non-project files under solution root were included." },
                ["count"] = new JsonObject { ["type"] = "integer", [SearchDescriptionProperty] = "Number of files found." },
                ["totalCount"] = new JsonObject { ["type"] = "integer", [SearchDescriptionProperty] = "Total source result count before chunking." },
                ["filteredCount"] = new JsonObject { ["type"] = "integer", [SearchDescriptionProperty] = "Result count after post-search filters." },
                ["chunkIndex"] = new JsonObject { ["type"] = "integer", [SearchDescriptionProperty] = "Zero-based chunk index returned." },
                ["chunkSize"] = new JsonObject { ["type"] = "integer", [SearchDescriptionProperty] = "Requested chunk size; zero means all filtered rows." },
                ["chunkCount"] = new JsonObject { ["type"] = "integer", [SearchDescriptionProperty] = "Number of available chunks after filtering." },
                ["chunkStart"] = new JsonObject { ["type"] = "integer", [SearchDescriptionProperty] = "Zero-based start offset of this chunk." },
                ["chunkEnd"] = new JsonObject { ["type"] = "integer", [SearchDescriptionProperty] = "Exclusive end offset of this chunk." },
                ["hasMoreChunks"] = new JsonObject { ["type"] = "boolean", [SearchDescriptionProperty] = "Whether another chunk is available." },
                ["truncated"] = new JsonObject { ["type"] = "boolean", [SearchDescriptionProperty] = "Whether results were truncated by source limits or chunking." },
                ["chunkOutOfRange"] = new JsonObject { ["type"] = "boolean", [SearchDescriptionProperty] = "Whether the requested chunk index is outside the filtered results." },
                ["sortBy"] = new JsonObject { ["type"] = "string", [SearchDescriptionProperty] = "Sort field applied to this result, when requested." },
                ["sortDirection"] = new JsonObject { ["type"] = "string", [SearchDescriptionProperty] = "Sort direction applied to this result, when requested." },
                ["groups"] = new JsonObject
                {
                    ["type"] = "array",
                    [SearchDescriptionProperty] = "Optional group summaries when group_by is requested.",
                    ["items"] = new JsonObject { ["type"] = "object" },
                },
                ["matches"] = new JsonObject
                {
                    ["type"] = "array",
                    [SearchDescriptionProperty] = "Array of file match objects sorted by relevance score.",
                    ["items"] = new JsonObject
                    {
                        ["type"] = "object",
                        ["properties"] = new JsonObject
                        {
                            ["path"] = new JsonObject { ["type"] = "string", [SearchDescriptionProperty] = "Absolute file system path to the file." },
                            ["name"] = new JsonObject { ["type"] = "string", [SearchDescriptionProperty] = "Just the file name component without directory." },
                            ["project"] = new JsonObject { ["type"] = "string", [SearchDescriptionProperty] = "The project unique name if the file is in a project, empty string otherwise." },
                            ["score"] = new JsonObject { ["type"] = "number", [SearchDescriptionProperty] = "Relevance score for sorting results (higher is better match)." },
                            ["source"] = new JsonObject { ["type"] = "string", [SearchDescriptionProperty] = "Source of the match: 'project' or 'disk'." },
                        },
                        ["required"] = new JsonArray { "path", "name", "project", "score", "source" },
                        ["additionalProperties"] = false,
                    },
                },
            },
            ["required"] = new JsonArray
            {
                "query", "pathFilter", SearchExtensionsProperty, "includeNonProject", "count", "totalCount",
                "filteredCount", "chunkIndex", "chunkSize", "chunkCount", "chunkStart", "chunkEnd",
                "hasMoreChunks", "truncated", "chunkOutOfRange", "matches",
            },
            ["additionalProperties"] = false,
        };

    // Strip bridge-internal envelope fields the infrastructure appends to every Data payload
    // (e.g. "queue" timing info) so the result passes the strict output schema validation.
    private static JsonObject StripBridgeEnvelopeFields(JsonObject data)
    {
        JsonObject clone = (JsonObject)data.DeepClone();
        clone.Remove("queue");
        return clone;
    }

    private static IEnumerable<ToolEntry> CodeExplorationTools()
    {
        yield return BridgeTool("search_solutions",
            "Search for solution files (.sln/.slnx) on disk under a given root directory. " +
            "Defaults to %USERPROFILE%\\source\\repos.",
            ObjectSchema(WithSearchResultOptions(
                Opt(Path, "Root directory to search."),
                Opt(Query, "Filter by solution name (case-insensitive substring)."),
                OptInt("max_depth", "Max directory depth to recurse (default 6)."),
                OptInt(Max, "Max results to return (default 200)."))),
            "search-solutions",
            a => Build(
                (Path, OptionalString(a, Path)),
                (Query, OptionalString(a, Query)),
                ("max-depth", OptionalText(a, "max_depth")),
                (Max, OptionalText(a, Max))),
            Search,
            searchHints: BuildSearchHints(
                workflow: [("open_solution", "Open the found solution")],
                related: [("list_instances", "List already-open instances"), ("bind_solution", "Bind to an open solution")]),
            transformResponse: (response, args) => CompactSearchResponse(response, args, includePathFilter: false, SearchJsonNames.Solutions, SearchJsonNames.Results, SearchJsonNames.Matches, SearchJsonNames.Files));

        yield return BridgeTool("smart_context",
            "First call for open-ended code exploration. " +
            "Collects focused context for a natural-language query — searches symbols, usages, and related definitions. " +
            "Prefer over read_file + find_text when you don't know exactly where to look. It is more expensive than direct read/search tools and avoids populating the VS Find Results window unless explicitly requested.",
            ObjectSchema(WithSearchResultOptions(
                Req(Query, "Natural-language description of what you are looking for."),
                OptInt("max_contexts", "Max context blocks to return (default 3)."))),
            "smart-context",
            a => Build(
                (Query, OptionalString(a, Query)),
                ("max-contexts", OptionalText(a, "max_contexts"))),
            Search,
            searchHints: BuildSearchHints(
                workflow: [(SearchReadFileTool, "Read a file from the returned context"), ("apply_diff", "Edit code after exploring")],
                related: [(SearchFindTextTool, "Search for specific text"), (SearchSymbolsTool, "Search by symbol name")]),
            transformResponse: (response, args) => CompactSearchResponse(response, args, includePathFilter: true, SearchJsonNames.Contexts, SearchJsonNames.Matches, SearchJsonNames.Results, SearchJsonNames.Symbols));

        yield return BridgeTool("file_symbols",
            "List symbols in one file with optional kind filtering. " +
            "More targeted than file_outline — use when you want only functions, " +
            "only classes, etc.",
            ObjectSchema(WithSearchResultOptions(
                Req("file", FileDesc),
                Opt("kind", "Optional symbol kind filter (e.g. function, class, member)."))),
            "file-symbols",
            a => Build(
                ("file", OptionalString(a, "file")),
                ("kind", OptionalString(a, "kind"))),
            Search,
            searchHints: BuildSearchHints(
                workflow: [(SearchReadFileTool, "Read the implementation of a listed symbol"), ("goto_definition", "Navigate to the symbol definition")],
                related: [(SearchFileOutlineTool, "Get the full outline with hierarchy"), (SearchSymbolsTool, "Search symbols across the solution")]),
            transformResponse: (response, args) => CompactSearchResponse(response, args, includePathFilter: true, SearchJsonNames.Symbols, SearchJsonNames.Matches));

        yield return new("glob",
            "Find files by glob pattern — use instead of Glob. " +
            "Supports **/*.cs, src/**/*.tsx, *.sln and any pattern supported by " +
            "Microsoft.Extensions.FileSystemGlobbing. Root defaults to solution directory. " +
            "Each match includes a \"handle\" field (f:N) — pass it as the file argument to read_file, file_outline, apply_diff, or write_file instead of copying the full path.",
            ObjectSchema(WithSearchResultOptions(
                Req("pattern", "Glob pattern, e.g. \"**/*.cs\", \"src/**/*.tsx\", \"*.sln\"."),
                Opt(Path, "Root directory. Defaults to solution directory."),
                OptInt(Max, "Max results (default 200)."))),
            Search,
            (id, args, bridge) => SystemTools.GlobTool.ExecuteAsync(id, args, bridge),
            aliases: ["find_by_pattern", "glob_files", "list_files"],
            summary: "Find files by glob pattern — use instead of Glob.",
            readOnly: true,
            mutating: false,
            destructive: false,
            searchHints: BuildSearchHints(
                workflow: [(SearchReadFileTool, "Read one of the matched files"), ("read_file_batch", "Read multiple matched files at once")],
                related: [("find_files", "Find files by name fragment"), (SearchFindTextTool, "Search file contents")]));
    }
}
