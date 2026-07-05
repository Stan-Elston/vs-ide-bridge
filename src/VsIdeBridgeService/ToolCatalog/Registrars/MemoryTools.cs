using VsIdeBridgeService.SystemTools;
using static VsIdeBridgeService.SchemaHelpers;

namespace VsIdeBridgeService;

internal static partial class ToolCatalog
{
    private const string MemoryCategory = "memory";

    private static IEnumerable<ToolEntry> MemoryTools()
    {
        yield return new("memory_search",
            "Search Codex memory files without writing anything to disk. Supports exact text and ranked " +
            "multi-term queries; results include memory-relative paths, line numbers, previews, and small " +
            "bounded context snippets.",
            ObjectSchema(
                Req("query", "Case-insensitive text or space-separated terms to search for in Codex memory files."),
                OptInt("max_results", "Maximum matches to return (default 25, max 100)."),
                OptInt("context_lines", "Context lines before and after each match (default 1, max 5)."),
                OptBool("include_rollouts", "Also search rollout summaries when true (default true).")),
            MemoryCategory,
            (id, args, bridge) => MemoryTool.SearchAsync(id, args, bridge),
            aliases: ["search_memory", "codex_memory_search", "memory_find"],
            tags: ["memory", "codex", "search", "read"],
            summary: "Search bounded Codex memory files.",
            readOnly: true,
            searchHints: BuildSearchHints(
                workflow: [("memory_read", "Read a bounded snippet around a match")],
                related: [("tool_help", "Inspect memory tool schemas")]));

        yield return new("memory_read",
            "Read a bounded snippet from a Codex memory file without writing anything to disk. Paths are relative " +
            "to the resolved memory root and cannot escape that root.",
            ObjectSchema(
                Opt("path", "Memory-relative path to read (default MEMORY.md)."),
                OptInt("start_line", "First 1-based line to read (default 1)."),
                OptInt("end_line", "Last 1-based line to read. Output is capped at 250 lines."),
                OptInt("line", "Center line to read around; overrides start_line/end_line when provided."),
                OptInt("context_lines", "Lines before and after line when line is provided (default 20, max 100).")),
            MemoryCategory,
            (id, args, bridge) => MemoryTool.ReadAsync(id, args, bridge),
            aliases: ["read_memory", "codex_memory_read", "memory_snippet"],
            tags: ["memory", "codex", "read", "snippet"],
            summary: "Read bounded Codex memory snippets.",
            readOnly: true,
            searchHints: BuildSearchHints(
                workflow: [("memory_search", "Find relevant memory files and lines first")],
                related: [("tool_help", "Inspect memory tool schemas")]));
    }
}
