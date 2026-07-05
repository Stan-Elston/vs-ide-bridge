using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json.Nodes;

namespace VsIdeBridge.Diagnostics;

public static class DiagnosticJsonNames
{
    public const string Severity = "severity";
    public const string Code = "code";
    public const string Project = "project";
    public const string File = "file";
    public const string Path = "path";
    public const string Line = "line";
    public const string Column = "column";
    public const string Message = "message";
    public const string Source = "source";
    public const string Tool = "tool";
    public const string Rows = "rows";
    public const string Count = "count";
    public const string TotalCount = "totalCount";
    public const string FilteredCount = "filteredCount";
    public const string Truncated = "truncated";
    public const string ChunkOutOfRange = "chunkOutOfRange";
    public const string ChunkIndex = "chunkIndex";
    public const string ChunkSize = "chunkSize";
    public const string ChunkCount = "chunkCount";
    public const string ChunkStart = "chunkStart";
    public const string ChunkEnd = "chunkEnd";
    public const string HasMoreChunks = "hasMoreChunks";
    public const string Groups = "groups";
    public const string Key = "key";
    public const string GroupBy = "groupBy";
    public const string SampleMessage = "sampleMessage";
    public const string SampleFile = "sampleFile";
    public const string SampleCode = "sampleCode";
    public const string SortBy = "sortBy";
    public const string SortDirection = "sortDirection";
    public const string ResponseData = "Data";
}

public sealed class DiagnosticRow
{
    private readonly JsonObject _original;

    private DiagnosticRow(JsonObject original)
    {
        _original = ApplyAnalyzerLimitMetadata(original);
        Severity = GetString(_original, DiagnosticJsonNames.Severity);
        Code = GetString(_original, DiagnosticJsonNames.Code);
        Project = GetString(_original, DiagnosticJsonNames.Project);
        File = GetString(_original, DiagnosticJsonNames.File);
        Path = GetString(_original, DiagnosticJsonNames.Path);
        Line = GetNullableInt(_original, DiagnosticJsonNames.Line);
        Column = GetNullableInt(_original, DiagnosticJsonNames.Column);
        Message = GetString(_original, DiagnosticJsonNames.Message);
        Source = GetString(_original, DiagnosticJsonNames.Source);
        Tool = GetString(_original, DiagnosticJsonNames.Tool);
    }

    private static JsonObject ApplyAnalyzerLimitMetadata(JsonObject row)
    {
        string code = GetString(row, DiagnosticJsonNames.Code);
        if (!IsVisualStudioAnalyzerLimitedCode(code))
        {
            return row;
        }

        row["analyzerLimited"] = true;
        row["codeFamily"] = "analyzer";
        row[DiagnosticJsonNames.Source] = string.IsNullOrWhiteSpace(GetString(row, DiagnosticJsonNames.Source))
            ? "visual-studio-analyzer"
            : GetString(row, DiagnosticJsonNames.Source);
        row[DiagnosticJsonNames.Tool] = string.IsNullOrWhiteSpace(GetString(row, DiagnosticJsonNames.Tool))
            ? "visual-studio-analyzer"
            : GetString(row, DiagnosticJsonNames.Tool);
        row["authority"] = "visual-studio-analyzer-limited";
        row["guidance"] = MergeAnalyzerText(
            GetString(row, "guidance"),
            "Visual Studio analyzer-limited row: VCR001/VCR003 can lag or misread modern C++ constructs such as '= delete' and explicit specializations. Verify against current source and build output before editing.");
        row["suggestedAction"] = MergeAnalyzerText(
            GetString(row, "suggestedAction"),
            "Treat as advisory analyzer output; refresh diagnostics or build before making semantic changes.");
        row["llmFixPrompt"] = MergeAnalyzerText(
            GetString(row, "llmFixPrompt"),
            "Do not assume this VCR row is authoritative. Confirm the construct in source/build output first, especially for '= delete' and explicit specialization patterns.");
        return row;
    }

    private static bool IsVisualStudioAnalyzerLimitedCode(string code)
        => string.Equals(code, "VCR001", StringComparison.OrdinalIgnoreCase)
            || string.Equals(code, "VCR003", StringComparison.OrdinalIgnoreCase);

    private static string MergeAnalyzerText(string existing, string addition)
    {
        if (string.IsNullOrWhiteSpace(existing))
        {
            return addition;
        }

        return existing.Contains(addition, StringComparison.Ordinal)
            ? existing
            : string.Concat(existing, " ", addition);
    }

    public string Severity { get; }

    public string Code { get; }

    public string Project { get; }

    public string File { get; }

    public string Path { get; }

    public int? Line { get; }

    public int? Column { get; }

    public string Message { get; }

    public string Source { get; }

    public string Tool { get; }

    public string DisplayPath => !string.IsNullOrWhiteSpace(File) ? File : Path;

    public static DiagnosticRow FromJson(JsonObject row)
        => new(row.DeepClone().AsObject());

    public JsonObject ToJson()
        => _original.DeepClone().AsObject();

    public bool Matches(DiagnosticQueryOptions options)
    {
        return MatchesSeverity(options.Severity)
            && MatchesPrefix(Code, options.Code)
            && Contains(Project, options.Project)
            && MatchesPath(options.Path)
            && MatchesPath(options.File)
            && Contains(Message, options.Text);
    }

    public string GetSortText(string propertyName)
    {
        return propertyName switch
        {
            DiagnosticJsonNames.Severity => Severity,
            DiagnosticJsonNames.Code => Code,
            DiagnosticJsonNames.Project => Project,
            DiagnosticJsonNames.File or DiagnosticJsonNames.Path => DisplayPath,
            DiagnosticJsonNames.Message => Message,
            DiagnosticJsonNames.Source => Source,
            DiagnosticJsonNames.Tool => Tool,
            _ => string.Empty,
        };
    }

    public int GetSortNumber(string propertyName)
    {
        return propertyName switch
        {
            DiagnosticJsonNames.Line => Line ?? int.MaxValue,
            DiagnosticJsonNames.Column => Column ?? int.MaxValue,
            _ => int.MaxValue,
        };
    }

    private bool MatchesSeverity(string? severity)
    {
        if (string.IsNullOrWhiteSpace(severity) || string.Equals(severity, "all", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return string.Equals(Severity, NormalizeSeverity(severity!), StringComparison.OrdinalIgnoreCase);
    }

    private bool MatchesPath(string? path)
        => ContainsPath(DisplayPath, path) || ContainsPath(Path, path) || ContainsPath(File, path);

    private static bool MatchesPrefix(string value, string? prefix)
        => string.IsNullOrWhiteSpace(prefix) || value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);

    private static bool Contains(string value, string? text)
        => string.IsNullOrWhiteSpace(text) || value.Contains(text, StringComparison.OrdinalIgnoreCase);

    // Path comparison normalises separators so callers can pass either '\' or '/'
    // regardless of OS or how the path filter was typed.
    private static bool ContainsPath(string value, string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return true;
        return value.Replace('\\', '/').Contains(text!.Replace('\\', '/'), StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeSeverity(string severity)
    {
        return severity.Trim().ToLowerInvariant() switch
        {
            "error" or "errors" => "Error",
            "warning" or "warnings" => "Warning",
            "message" or "messages" or "info" or "information" => "Message",
            _ => severity.Trim(),
        };
    }

    private static string GetString(JsonObject row, string propertyName)
    {
        JsonNode? node = row[propertyName];
        if (node is JsonValue valueNode && valueNode.TryGetValue(out string? value))
        {
            return value ?? string.Empty;
        }

        return node?.ToString() ?? string.Empty;
    }

    private static int? GetNullableInt(JsonObject row, string propertyName)
    {
        if (row[propertyName] is not JsonValue valueNode)
        {
            return null;
        }

        if (valueNode.TryGetValue(out int value))
        {
            return value;
        }

        return int.TryParse(valueNode.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value)
            ? value
            : null;
    }
}

public sealed class DiagnosticCollection
{
    public DiagnosticCollection(
        IEnumerable<DiagnosticRow> rows,
        int? sourceCount = null,
        int? sourceTotalCount = null,
        bool sourceTruncated = false)
    {
        Rows = [.. rows];
        SourceCount = sourceCount ?? Rows.Count;
        SourceTotalCount = sourceTotalCount ?? SourceCount;
        SourceTruncated = sourceTruncated;
    }

    public IReadOnlyList<DiagnosticRow> Rows { get; }

    public int SourceCount { get; }

    public int SourceTotalCount { get; }

    public bool SourceTruncated { get; }

    public static DiagnosticCollection Empty { get; } = new([]);

    public static DiagnosticCollection FromJsonObject(JsonObject source)
    {
        IReadOnlyList<DiagnosticRow> rows = source[DiagnosticJsonNames.Rows] is JsonArray rowArray
            ? [.. rowArray.OfType<JsonObject>().Select(DiagnosticRow.FromJson)]
            : [];

        int sourceCount = GetOptionalInt(source, DiagnosticJsonNames.Count) ?? rows.Count;
        int sourceTotalCount = GetOptionalInt(source, DiagnosticJsonNames.TotalCount) ?? sourceCount;
        bool sourceTruncated = GetOptionalBool(source, DiagnosticJsonNames.Truncated) == true || sourceTotalCount > rows.Count;
        return new DiagnosticCollection(rows, sourceCount, sourceTotalCount, sourceTruncated);
    }

    public DiagnosticCollection WithSeverity(string? severity)
        => string.IsNullOrWhiteSpace(severity) || string.Equals(severity, "all", StringComparison.OrdinalIgnoreCase)
            ? this
            : WithFilteredRows(row => row.Matches(new DiagnosticQueryOptions(0, 0, null, false, severity, null, null, null, null, null)));

    public DiagnosticCollection WithCodePrefix(string? code)
        => string.IsNullOrWhiteSpace(code)
            ? this
            : WithFilteredRows(row => row.Code.StartsWith(code, StringComparison.OrdinalIgnoreCase));

    public DiagnosticCollection WithProject(string? project)
        => string.IsNullOrWhiteSpace(project)
            ? this
            : WithFilteredRows(row => row.Project.Contains(project, StringComparison.OrdinalIgnoreCase));

    public DiagnosticCollection WithPath(string? path)
        => string.IsNullOrWhiteSpace(path)
            ? this
            : WithFilteredRows(row => row.DisplayPath.Contains(path, StringComparison.OrdinalIgnoreCase)
                || row.Path.Contains(path, StringComparison.OrdinalIgnoreCase)
                || row.File.Contains(path, StringComparison.OrdinalIgnoreCase));

    public DiagnosticCollection WithText(string? text)
        => string.IsNullOrWhiteSpace(text)
            ? this
            : WithFilteredRows(row => row.Message.Contains(text, StringComparison.OrdinalIgnoreCase));

    public DiagnosticCollection ApplyFilters(DiagnosticQueryOptions options)
    {
        if (!options.HasContentFilters)
        {
            return this;
        }

        return WithFilteredRows(row => row.Matches(options));
    }

    public DiagnosticCollection Sort(DiagnosticQueryOptions options)
    {
        if (string.IsNullOrWhiteSpace(options.SortBy))
        {
            return this;
        }

        IEnumerable<DiagnosticRow> sortedRows = options.SortBy is DiagnosticJsonNames.Line or DiagnosticJsonNames.Column
            ? SortByNumber(options)
            : SortByText(options);

        return new DiagnosticCollection(sortedRows, SourceCount, SourceTotalCount, SourceTruncated);
    }

    public DiagnosticPage Page(DiagnosticQueryOptions options)
    {
        int filteredCount = EffectiveFilteredCount();
        int chunkCount = CalculateChunkCount(filteredCount, options.ChunkSize);
        int chunkStart = CalculateChunkStart(options.ChunkIndex, options.ChunkSize, filteredCount);
        int availableStart = Math.Min(chunkStart, Rows.Count);
        int availableEnd = options.ChunkSize == 0
            ? Rows.Count
            : Math.Min(availableStart + options.ChunkSize, Rows.Count);
        bool chunkOutOfRange = options.ChunkSize > 0 && options.ChunkIndex >= chunkCount && filteredCount > 0;

        IReadOnlyList<DiagnosticRow> pageRows = chunkOutOfRange
            ? []
            : [.. Rows.Skip(availableStart).Take(availableEnd - availableStart)];

        return new DiagnosticPage(
            pageRows,
            filteredCount,
            options.ChunkIndex,
            options.ChunkSize,
            chunkCount,
            chunkOutOfRange ? filteredCount : Math.Min(chunkStart, filteredCount),
            chunkOutOfRange ? filteredCount : Math.Min(chunkStart + pageRows.Count, filteredCount),
            options.ChunkSize > 0 && options.ChunkIndex + 1 < chunkCount,
            SourceTruncated || chunkCount > 1 || chunkOutOfRange,
            chunkOutOfRange);
    }

    public JsonArray GroupBy(DiagnosticQueryOptions options)
    {
        string? normalized = DiagnosticQueryOptions.NormalizeGroupBy(options.GroupBy);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return [];
        }

        string normalizedGroupBy = normalized!;
        List<(IGrouping<string, DiagnosticRow> Group, int Count)> buckets =
            [.. Rows
                .GroupBy(row => GetGroupKey(row, normalizedGroupBy), StringComparer.OrdinalIgnoreCase)
                .Select(g => (Group: g, Count: g.Count()))];

        if (options.GroupMinCount.HasValue || options.GroupMaxCount.HasValue)
        {
            buckets = [.. buckets.Where(b =>
                (!options.GroupMinCount.HasValue || b.Count >= options.GroupMinCount.Value) &&
                (!options.GroupMaxCount.HasValue || b.Count <= options.GroupMaxCount.Value))];
        }

        string? groupSortBy = options.GroupSortBy;
        bool sortDescending = options.GroupSortDescending ?? groupSortBy != "key";
        IEnumerable<(IGrouping<string, DiagnosticRow> Group, int Count)> ordered;
        if (groupSortBy == "key")
        {
            ordered = sortDescending
                ? buckets.OrderByDescending(b => b.Group.Key, StringComparer.OrdinalIgnoreCase)
                : buckets.OrderBy(b => b.Group.Key, StringComparer.OrdinalIgnoreCase);
        }
        else
        {
            ordered = sortDescending
                ? buckets.OrderByDescending(b => b.Count).ThenBy(b => b.Group.Key, StringComparer.OrdinalIgnoreCase)
                : buckets.OrderBy(b => b.Count).ThenBy(b => b.Group.Key, StringComparer.OrdinalIgnoreCase);
        }

        JsonArray groups = [];
        foreach ((IGrouping<string, DiagnosticRow> group, int count) in ordered)
        {
            DiagnosticRow sample = group.First();
            groups.Add(new JsonObject
            {
                [DiagnosticJsonNames.Key] = group.Key,
                [DiagnosticJsonNames.GroupBy] = normalizedGroupBy,
                [DiagnosticJsonNames.Count] = count,
                [DiagnosticJsonNames.SampleMessage] = sample.Message,
                [DiagnosticJsonNames.SampleFile] = sample.DisplayPath,
                [DiagnosticJsonNames.SampleCode] = sample.Code,
            });
        }

        return groups;
    }

    public JsonObject ToJsonObject(DiagnosticQueryOptions options, JsonObject? template = null)
    {
        JsonObject target = template?.DeepClone().AsObject() ?? [];
        WriteTo(target, options);
        return target;
    }

    public void WriteTo(JsonObject target, DiagnosticQueryOptions options)
    {
        DiagnosticCollection filtered = ApplyFilters(options);
        DiagnosticCollection sorted = filtered.Sort(options);
        DiagnosticPage page = sorted.Page(options);
        JsonArray rows = [];
        foreach (DiagnosticRow row in page.Rows)
        {
            rows.Add(row.ToJson());
        }

        target[DiagnosticJsonNames.Rows] = rows;
        target[DiagnosticJsonNames.Count] = rows.Count;
        target[DiagnosticJsonNames.TotalCount] = page.FilteredCount;
        target[DiagnosticJsonNames.FilteredCount] = page.FilteredCount;
        target[DiagnosticJsonNames.ChunkIndex] = page.ChunkIndex;
        target[DiagnosticJsonNames.ChunkSize] = page.ChunkSize;
        target[DiagnosticJsonNames.ChunkCount] = page.ChunkCount;
        target[DiagnosticJsonNames.ChunkStart] = page.ChunkStart;
        target[DiagnosticJsonNames.ChunkEnd] = page.ChunkEnd;
        target[DiagnosticJsonNames.HasMoreChunks] = page.HasMoreChunks;
        target[DiagnosticJsonNames.Truncated] = page.Truncated;

        if (page.ChunkOutOfRange)
        {
            target[DiagnosticJsonNames.ChunkOutOfRange] = true;
        }
        else
        {
            target.Remove(DiagnosticJsonNames.ChunkOutOfRange);
        }

        if (!string.IsNullOrWhiteSpace(options.SortBy))
        {
            target[DiagnosticJsonNames.SortBy] = options.SortBy;
            target[DiagnosticJsonNames.SortDirection] = options.SortDescending ? "desc" : "asc";
        }

        if (!string.IsNullOrWhiteSpace(options.GroupBy))
        {
            // Suppress individual rows when grouping — the compact group summary (key, count,
            // sampleMessage, sampleFile, sampleCode) is all a model needs to orient itself.
            // To drill into a specific group's rows, call again with a path/code/project filter
            // and no group_by.
            //
            // When this collection only holds a truncated row subset (the source payload was
            // already row-chunked), regrouping here would shrink every group count to the
            // current chunk. The upstream producer groups over the full filtered set, so
            // prefer its group summary when one is present.
            JsonArray allGroups = sorted.SourceTruncated
                && target[DiagnosticJsonNames.Groups] is JsonArray upstreamGroups
                && upstreamGroups.Count > 0
                ? (JsonArray)upstreamGroups.DeepClone()
                : sorted.GroupBy(options);
            int groupCount = allGroups.Count;

            // Paginate the groups themselves by chunk_size / chunk_index.
            // The row-level pagination written above is based on individual rows, so override
            // every pagination field here so the caller sees group-level pagination.
            int effectiveChunkSize = options.ChunkSize;
            int groupChunkCount = groupCount == 0 ? 0
                : effectiveChunkSize == 0 ? 1
                : (int)Math.Ceiling(groupCount / (double)effectiveChunkSize);
            int groupChunkStart = effectiveChunkSize == 0 ? 0
                : Math.Min(options.ChunkIndex * effectiveChunkSize, groupCount);
            int groupChunkEnd = effectiveChunkSize == 0 ? groupCount
                : Math.Min(groupChunkStart + effectiveChunkSize, groupCount);
            bool groupHasMoreChunks = effectiveChunkSize > 0 && options.ChunkIndex + 1 < groupChunkCount;

            JsonArray pagedGroups = [];
            foreach (JsonNode? node in allGroups.Skip(groupChunkStart).Take(groupChunkEnd - groupChunkStart))
            {
                pagedGroups.Add(node?.DeepClone());
            }

            target[DiagnosticJsonNames.Groups] = pagedGroups;
            target[DiagnosticJsonNames.Rows] = new JsonArray();
            // In grouped mode the generic count/pagination fields below refer to GROUP rows,
            // not diagnostics. State both populations explicitly so 14 groups is never misread
            // as 14 warnings (severityCounts still carries the per-severity diagnostic totals).
            target["countScope"] = "groups";
            target["groupCount"] = groupCount;
            target["diagnosticCount"] = EffectiveFilteredCount();
            target[DiagnosticJsonNames.Count] = groupCount;
            target[DiagnosticJsonNames.TotalCount] = groupCount;
            target[DiagnosticJsonNames.FilteredCount] = groupCount;
            target[DiagnosticJsonNames.ChunkIndex] = options.ChunkIndex;
            target[DiagnosticJsonNames.ChunkSize] = effectiveChunkSize;
            target[DiagnosticJsonNames.ChunkCount] = groupChunkCount;
            target[DiagnosticJsonNames.ChunkStart] = groupChunkStart;
            target[DiagnosticJsonNames.ChunkEnd] = groupChunkEnd;
            target[DiagnosticJsonNames.HasMoreChunks] = groupHasMoreChunks;
            target[DiagnosticJsonNames.Truncated] = groupHasMoreChunks;
            target.Remove(DiagnosticJsonNames.ChunkOutOfRange);
        }
    }

    public JsonObject ToUnpagedJsonObject(JsonObject? template = null)
    {
        JsonObject target = template?.DeepClone().AsObject() ?? [];
        JsonArray rows = [];
        foreach (DiagnosticRow row in Rows)
        {
            rows.Add(row.ToJson());
        }

        int filteredCount = EffectiveFilteredCount();
        target[DiagnosticJsonNames.Rows] = rows;
        target[DiagnosticJsonNames.Count] = rows.Count;
        target[DiagnosticJsonNames.TotalCount] = filteredCount;
        target[DiagnosticJsonNames.FilteredCount] = filteredCount;
        target[DiagnosticJsonNames.ChunkIndex] = 0;
        target[DiagnosticJsonNames.ChunkSize] = 0;
        target[DiagnosticJsonNames.ChunkCount] = filteredCount == 0 ? 0 : 1;
        target[DiagnosticJsonNames.ChunkStart] = 0;
        target[DiagnosticJsonNames.ChunkEnd] = rows.Count;
        target[DiagnosticJsonNames.HasMoreChunks] = false;
        target[DiagnosticJsonNames.Truncated] = SourceTruncated;
        target.Remove(DiagnosticJsonNames.ChunkOutOfRange);
        return target;
    }

    private DiagnosticCollection WithFilteredRows(Func<DiagnosticRow, bool> predicate)
    {
        IReadOnlyList<DiagnosticRow> filteredRows = [.. Rows.Where(predicate)];
        bool sourceTruncated = SourceTruncated && SourceTotalCount > Rows.Count;
        return new DiagnosticCollection(
            filteredRows,
            filteredRows.Count,
            filteredRows.Count,
            sourceTruncated);
    }

    private IEnumerable<DiagnosticRow> SortByNumber(DiagnosticQueryOptions options)
    {
        return options.SortDescending
            ? Rows.OrderByDescending(row => row.GetSortNumber(options.SortBy!))
                .ThenBy(row => row.GetSortText(DiagnosticJsonNames.File), StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.GetSortText(DiagnosticJsonNames.Message), StringComparer.OrdinalIgnoreCase)
            : Rows.OrderBy(row => row.GetSortNumber(options.SortBy!))
                .ThenBy(row => row.GetSortText(DiagnosticJsonNames.File), StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.GetSortText(DiagnosticJsonNames.Message), StringComparer.OrdinalIgnoreCase);
    }

    private IEnumerable<DiagnosticRow> SortByText(DiagnosticQueryOptions options)
    {
        return options.SortDescending
            ? Rows.OrderByDescending(row => row.GetSortText(options.SortBy!), StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.GetSortText(DiagnosticJsonNames.File), StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.GetSortNumber(DiagnosticJsonNames.Line))
            : Rows.OrderBy(row => row.GetSortText(options.SortBy!), StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.GetSortText(DiagnosticJsonNames.File), StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.GetSortNumber(DiagnosticJsonNames.Line));
    }

    private int EffectiveFilteredCount()
        => Math.Max(Rows.Count, Math.Max(SourceCount, SourceTotalCount));

    private static string GetGroupKey(DiagnosticRow row, string groupBy)
    {
        string key = groupBy switch
        {
            DiagnosticJsonNames.Severity => row.Severity,
            DiagnosticJsonNames.Code => row.Code,
            DiagnosticJsonNames.Project => row.Project,
            DiagnosticJsonNames.File or DiagnosticJsonNames.Path => row.DisplayPath,
            DiagnosticJsonNames.Source => row.Source,
            DiagnosticJsonNames.Tool => row.Tool,
            _ => string.Empty,
        };

        return string.IsNullOrWhiteSpace(key) ? "(blank)" : key;
    }

    private static int CalculateChunkCount(int filteredCount, int chunkSize)
    {
        if (filteredCount == 0)
        {
            return 0;
        }

        return chunkSize == 0
            ? 1
            : (int)Math.Ceiling(filteredCount / (double)chunkSize);
    }

    private static int CalculateChunkStart(int chunkIndex, int chunkSize, int filteredCount)
    {
        if (chunkSize == 0 || filteredCount == 0)
        {
            return 0;
        }

        try
        {
            return Math.Min(checked(chunkIndex * chunkSize), filteredCount);
        }
        catch (OverflowException)
        {
            return filteredCount;
        }
    }

    private static int? GetOptionalInt(JsonObject source, string propertyName)
    {
        if (source[propertyName] is not JsonValue valueNode)
        {
            return null;
        }

        if (valueNode.TryGetValue(out int value))
        {
            return value;
        }

        return int.TryParse(valueNode.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value)
            ? value
            : null;
    }

    private static bool? GetOptionalBool(JsonObject source, string propertyName)
    {
        if (source[propertyName] is not JsonValue valueNode)
        {
            return null;
        }

        if (valueNode.TryGetValue(out bool value))
        {
            return value;
        }

        return bool.TryParse(valueNode.ToString(), out value)
            ? value
            : null;
    }
}

public sealed class DiagnosticQueryOptions(
    int chunkSize,
    int chunkIndex,
    string? sortBy,
    bool sortDescending,
    string? severity,
    string? code,
    string? project,
    string? path,
    string? text,
    string? groupBy,
    string? file = null,
    string? groupSortBy = null,
    bool? groupSortDescending = null,
    int? groupMinCount = null,
    int? groupMaxCount = null)
{
    public int ChunkSize { get; } = chunkSize;

    public int ChunkIndex { get; } = chunkIndex;

    public string? SortBy { get; } = sortBy;

    public bool SortDescending { get; } = sortDescending;

    public string? Severity { get; } = severity;

    public string? Code { get; } = code;

    public string? Project { get; } = project;

    public string? Path { get; } = path;

    public string? Text { get; } = text;

    public string? GroupBy { get; } = groupBy;

    public string? File { get; } = file;

    public string? GroupSortBy { get; } = groupSortBy;

    public bool? GroupSortDescending { get; } = groupSortDescending;

    public int? GroupMinCount { get; } = groupMinCount;

    public int? GroupMaxCount { get; } = groupMaxCount;

    public bool HasContentFilters
        => !string.IsNullOrWhiteSpace(Severity)
            || !string.IsNullOrWhiteSpace(Code)
            || !string.IsNullOrWhiteSpace(Project)
            || !string.IsNullOrWhiteSpace(Path)
            || !string.IsNullOrWhiteSpace(File)
            || !string.IsNullOrWhiteSpace(Text);

    public static DiagnosticQueryOptions FromJsonObject(JsonObject? args, int defaultChunkSize)
    {
        int chunkSize = GetOptionalNonNegativeInt(args, "chunk_size")
            ?? GetOptionalNonNegativeInt(args, "chunk-size")
            ?? GetOptionalNonNegativeInt(args, "max")
            ?? defaultChunkSize;
        int chunkIndex = GetOptionalNonNegativeInt(args, "chunk_index") ?? 0;
        string? sortBy = NormalizeSortField(GetOptionalString(args, "sort_by"));
        bool sortDescending = IsDescendingSort(GetOptionalString(args, "sort_direction"));
        string? groupBy = NormalizeGroupBy(GetOptionalString(args, "group_by") ?? GetOptionalString(args, "group-by"));
        string? file = NullIfWhiteSpace(GetOptionalString(args, "file"));
        string? groupSortBy = NormalizeGroupSortBy(GetOptionalString(args, "group_sort_by") ?? GetOptionalString(args, "group-sort-by"));
        bool? groupSortDescending = ParseGroupSortDescending(GetOptionalString(args, "group_sort_direction") ?? GetOptionalString(args, "group-sort-direction"));
        int? groupMinCount = GetOptionalNonNegativeInt(args, "group_min_count") ?? GetOptionalNonNegativeInt(args, "group-min-count");
        int? groupMaxCount = GetOptionalNonNegativeInt(args, "group_max_count") ?? GetOptionalNonNegativeInt(args, "group-max-count");

        return new DiagnosticQueryOptions(
            chunkSize,
            chunkIndex,
            sortBy,
            sortDescending,
            NullIfWhiteSpace(GetOptionalString(args, "severity")),
            NullIfWhiteSpace(GetOptionalString(args, "code")),
            NullIfWhiteSpace(GetOptionalString(args, "project")),
            NullIfWhiteSpace(GetOptionalString(args, "path")),
            NullIfWhiteSpace(GetOptionalString(args, "text")),
            groupBy,
            file,
            groupSortBy,
            groupSortDescending,
            groupMinCount,
            groupMaxCount);
    }

    public static string? NormalizeSortField(string? sortBy)
    {
        string? normalized = sortBy?.Trim().ToLowerInvariant();
        return normalized switch
        {
            DiagnosticJsonNames.Severity => DiagnosticJsonNames.Severity,
            DiagnosticJsonNames.Code => DiagnosticJsonNames.Code,
            DiagnosticJsonNames.Project => DiagnosticJsonNames.Project,
            DiagnosticJsonNames.File or DiagnosticJsonNames.Path => DiagnosticJsonNames.File,
            DiagnosticJsonNames.Line => DiagnosticJsonNames.Line,
            DiagnosticJsonNames.Column => DiagnosticJsonNames.Column,
            DiagnosticJsonNames.Message => DiagnosticJsonNames.Message,
            DiagnosticJsonNames.Source => DiagnosticJsonNames.Source,
            DiagnosticJsonNames.Tool => DiagnosticJsonNames.Tool,
            _ => null,
        };
    }

    public static string? NormalizeGroupBy(string? groupBy)
    {
        string? normalized = groupBy?.Trim().ToLowerInvariant();
        return normalized switch
        {
            DiagnosticJsonNames.Severity => DiagnosticJsonNames.Severity,
            DiagnosticJsonNames.Code => DiagnosticJsonNames.Code,
            DiagnosticJsonNames.Project => DiagnosticJsonNames.Project,
            DiagnosticJsonNames.File or DiagnosticJsonNames.Path => DiagnosticJsonNames.File,
            DiagnosticJsonNames.Source => DiagnosticJsonNames.Source,
            DiagnosticJsonNames.Tool => DiagnosticJsonNames.Tool,
            _ => null,
        };
    }

    private static int? GetOptionalNonNegativeInt(JsonObject? args, string name)
    {
        if (args?[name] is not JsonNode node)
        {
            return null;
        }

        if (node is JsonValue valueNode && valueNode.TryGetValue(out int value))
        {
            return Math.Max(0, value);
        }

        return int.TryParse(node.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value)
            ? Math.Max(0, value)
            : null;
    }

    private static string? GetOptionalString(JsonObject? args, string name)
    {
        if (args?[name] is not JsonNode node)
        {
            return null;
        }

        if (node is JsonValue valueNode && valueNode.TryGetValue(out string? value))
        {
            return value;
        }

        return node.ToString();
    }

    private static bool IsDescendingSort(string? sortDirection)
        => string.Equals(sortDirection, "desc", StringComparison.OrdinalIgnoreCase)
            || string.Equals(sortDirection, "descending", StringComparison.OrdinalIgnoreCase);

    private static string? NormalizeGroupSortBy(string? value)
    {
        return value?.Trim().ToLowerInvariant() switch
        {
            "count" => "count",
            "key" or "name" => "key",
            _ => null,
        };
    }

    private static bool? ParseGroupSortDescending(string? sortDirection)
    {
        if (string.IsNullOrWhiteSpace(sortDirection))
        {
            return null;
        }

        return IsDescendingSort(sortDirection);
    }

    private static string? NullIfWhiteSpace(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value;
}

public readonly struct DiagnosticPage(
    IReadOnlyList<DiagnosticRow> rows,
    int filteredCount,
    int chunkIndex,
    int chunkSize,
    int chunkCount,
    int chunkStart,
    int chunkEnd,
    bool hasMoreChunks,
    bool truncated,
    bool chunkOutOfRange)
{
    public IReadOnlyList<DiagnosticRow> Rows { get; } = rows;

    public int FilteredCount { get; } = filteredCount;

    public int ChunkIndex { get; } = chunkIndex;

    public int ChunkSize { get; } = chunkSize;

    public int ChunkCount { get; } = chunkCount;

    public int ChunkStart { get; } = chunkStart;

    public int ChunkEnd { get; } = chunkEnd;

    public bool HasMoreChunks { get; } = hasMoreChunks;

    public bool Truncated { get; } = truncated;

    public bool ChunkOutOfRange { get; } = chunkOutOfRange;
}

public sealed class DiagnosticBucket
{
    private readonly JsonObject _responseTemplate;
    private readonly JsonObject _dataTemplate;

    private DiagnosticBucket(JsonObject responseTemplate, JsonObject dataTemplate, DiagnosticCollection diagnostics)
    {
        _responseTemplate = responseTemplate;
        _dataTemplate = dataTemplate;
        Diagnostics = diagnostics;
    }

    public DiagnosticCollection Diagnostics { get; }

    public static DiagnosticBucket FromResponse(JsonObject response)
    {
        JsonObject responseTemplate = response.DeepClone().AsObject();
        JsonObject dataTemplate = responseTemplate[DiagnosticJsonNames.ResponseData] is JsonObject data
            ? data.DeepClone().AsObject()
            : [];
        DiagnosticCollection diagnostics = DiagnosticCollection.FromJsonObject(dataTemplate);
        dataTemplate.Remove(DiagnosticJsonNames.Rows);
        responseTemplate.Remove(DiagnosticJsonNames.ResponseData);
        return new DiagnosticBucket(responseTemplate, dataTemplate, diagnostics);
    }

    public JsonObject ToResponseJson()
    {
        JsonObject response = _responseTemplate.DeepClone().AsObject();
        response[DiagnosticJsonNames.ResponseData] = Diagnostics.ToUnpagedJsonObject(_dataTemplate);
        return response;
    }

    public JsonObject ToDataJson()
        => Diagnostics.ToUnpagedJsonObject(_dataTemplate);

    /// <summary>
    /// Returns a compact summary with counts only — no rows — suitable for embedding in cache metadata.
    /// </summary>
    public JsonObject ToCountSummaryJson()
        => _dataTemplate.DeepClone().AsObject();
}
