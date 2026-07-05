using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using VsIdeBridge.Infrastructure;
using VsIdeBridge.Services;

namespace VsIdeBridge.Commands;

internal static partial class DebugBuildCommands
{
    private static ErrorListQuery CreateErrorListQuery(IdeCommandContext context, CommandArguments args, string? defaultSeverity = null)
    {
        return new ErrorListQuery
        {
            Severity      = args.GetString("severity") ?? defaultSeverity,
            Code          = args.GetString("code"),
            Project       = args.GetString("project"),
            Path          = ResolveDiagnosticPathFilter(context, args.GetString("path")),
            File          = ResolveDiagnosticPathFilter(context, args.GetString(FileArgument)),
            Text          = args.GetString("text"),
            // Try CLI form (group-by) first, then MCP JSON form (group_by).
            GroupBy       = args.GetString("group-by") ?? args.GetString(GroupByJsonArgument),
            Max           = GetDiagnosticsMax(args),
            ChunkSize     = GetNullableInt32(args, ChunkSizeArgument, ChunkSizeJsonArgument),
            ChunkIndex    = GetNullableInt32(args, ChunkIndexArgument, ChunkIndexJsonArgument),
            SortBy        = args.GetString(SortByArgument)        ?? args.GetString(SortByJsonArgument),
            SortDirection = args.GetString(SortDirectionArgument) ?? args.GetString(SortDirectionJsonArgument),
        };
    }

    private static string? ResolveDiagnosticPathFilter(IdeCommandContext context, string? filter)
    {
        if (string.IsNullOrWhiteSpace(filter) || !HandleService.IsHandle(filter))
        {
            return filter;
        }

        return context.Handles.ResolveFilePath(filter!);
    }

    // Max is the legacy per-request limit (--max); ChunkSize is now separate.
    private static int? GetDiagnosticsMax(CommandArguments args)
        => args.GetNullableInt32(MaxArgument);

    private static int? GetNullableInt32(CommandArguments args, params string[] names)
    {
        foreach (string name in names)
        {
            int? value = args.GetNullableInt32(name);
            if (value.HasValue)
            {
                return value;
            }
        }

        return null;
    }

    private static bool GetDiagnosticsForceRefresh(CommandArguments args)
        => args.GetBoolean("refresh", false);

    private static JObject FilterRowsBySeverity(JArray allRows, string severity, int? max)
    {
        IEnumerable<JToken> filtered = allRows
            .Where(r => string.Equals((string?)r["severity"], severity, StringComparison.OrdinalIgnoreCase));
        if (max is > 0)
            filtered = filtered.Take(max.Value);

        JToken[] commandResult = [.. filtered];
        return new JObject
        {
            ["count"] = commandResult.Length,
            ["rows"] = new JArray(commandResult.Select(r => r.DeepClone())),
        };
    }

    private static Task<JObject> GetSeverityDiagnosticsAsync(
        IdeCommandContext context,
        bool waitForIntellisense,
        int timeoutMilliseconds,
        bool quickSnapshot,
        string severity,
        int? max)
    {
        return GetDiagnosticsWithFallbackAsync(
            context,
            waitForIntellisense,
            timeoutMilliseconds,
            quickSnapshot,
            new ErrorListQuery
            {
                Severity = severity,
                Max = max,
            });
    }

    /// <summary>
    /// Builds a compact one-line summary. Unfiltered reads lead with whole Error List totals so
    /// models do not declare victory on 0 errors while warnings/messages remain. Content-filtered
    /// reads lead with the filtered population and label whole-list counts separately.
    /// </summary>
    internal static string BuildDiagnosticsCountSummary(JObject result)
    {
        JToken? total = result["totalSeverityCounts"];
        int errors   = total?["Error"]?.Value<int>()   ?? 0;
        int warnings = total?["Warning"]?.Value<int>() ?? 0;
        int messages = total?["Message"]?.Value<int>() ?? 0;
        string counts = $"{errors} error(s), {warnings} warning(s), {messages} message(s)";

        if (HasDiagnosticContentFilter(result))
        {
            string severity = NormalizeSummarySeverity(result["filter"]?["severity"]?.Value<string>());
            int filteredCount = result["totalCount"]?.Value<int>() ?? result["count"]?.Value<int>() ?? 0;
            string filteredCounts = $"{filteredCount} {severity.ToLowerInvariant()}(s) for current filter";
            string globalCounts = $"whole Error List: {counts}";
            return severity switch
            {
                "Error" when filteredCount > 0 => $"{filteredCounts} - fix filtered errors before building; {globalCounts}.",
                "Error" => $"{filteredCounts} - no errors in this filtered set; {globalCounts}.",
                _ when filteredCount > 0 => $"{filteredCounts} - filtered diagnostics remain; {globalCounts}.",
                _ => $"{filteredCounts} - filtered set is clean; {globalCounts}.",
            };
        }

        if (errors > 0)
        {
            return $"{counts} - fix errors before building.";
        }

        // Warnings/messages don't block the build, but don't let the model declare victory.
        return warnings == 0 && messages == 0
            ? $"{counts} - clean."
            : $"{counts} - no errors; the build can succeed, but warnings/messages remain to fix.";
    }

    private static string NormalizeSummarySeverity(string? severity)
    {
        if (string.Equals(severity, "Warning", StringComparison.OrdinalIgnoreCase)) return "Warning";
        if (string.Equals(severity, "Message", StringComparison.OrdinalIgnoreCase)) return "Message";
        if (string.Equals(severity, "Error", StringComparison.OrdinalIgnoreCase)) return "Error";
        return "diagnostic";
    }

    private static bool HasDiagnosticContentFilter(JObject result)
    {
        if (result["filter"] is not JObject filter)
        {
            return false;
        }

        return HasNonEmptyFilterValue(filter, "code")
            || HasNonEmptyFilterValue(filter, "project")
            || HasNonEmptyFilterValue(filter, "path")
            || HasNonEmptyFilterValue(filter, FileArgument)
            || HasNonEmptyFilterValue(filter, "text");
    }

    private static bool HasNonEmptyFilterValue(JObject filter, string key)
        => !string.IsNullOrWhiteSpace(filter[key]?.Value<string>());

    /// <summary>
    /// Produces bridge-level advisory warnings pointing to severity categories that were
    /// NOT the primary focus of this call, so the model is reminded to check them.
    /// </summary>
    private static JArray BuildDiagnosticsCrossReferenceAdvisories(JObject result, params string[] otherSeverities)
    {
        JToken? counts = result["totalSeverityCounts"];
        if (counts is null)
            return [];

        bool contentFiltered = HasDiagnosticContentFilter(result);
        JArray advisories = [];
        foreach (string severity in otherSeverities)
        {
            int count = counts[severity]?.Value<int>() ?? 0;
            if (count > 0)
            {
                string callHint = severity switch
                {
                    "Error" => "run errors",
                    "Warning" => "run warnings",
                    "Message" => "run warnings with severity=Message",
                    _ => $"run {severity.ToLowerInvariant()}",
                };
                string prefix = contentFiltered ? "Whole Error List has" : "Also found";
                advisories.Add($"{prefix} {count} {severity.ToLowerInvariant()}(s) - {callHint}.");
            }
        }
        return advisories;
    }
}
