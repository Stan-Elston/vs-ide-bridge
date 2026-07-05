using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json.Nodes;

namespace VsIdeBridge.Tooling.Search;

public static class SearchJsonNames
{
    public const string Matches = "matches";
    public const string Files = "files";
    public const string Results = "results";
    public const string Symbols = "symbols";
    public const string Contexts = "contexts";
    public const string Solutions = "solutions";
    public const string Count = "count";
    public const string TotalCount = "totalCount";
    public const string TotalMatchCount = "totalMatchCount";
    public const string FilteredCount = "filteredCount";
    public const string Truncated = "truncated";
    public const string ChunkOutOfRange = "chunkOutOfRange";
    public const string ChunkIndex = "chunkIndex";
    public const string ChunkSize = "chunkSize";
    public const string ChunkCount = "chunkCount";
    public const string ChunkStart = "chunkStart";
    public const string ChunkEnd = "chunkEnd";
    public const string HasMoreChunks = "hasMoreChunks";
    public const string SortBy = "sortBy";
    public const string SortDirection = "sortDirection";
    public const string Path = "path";
    public const string File = "file";
    public const string Name = "name";
    public const string FullName = "fullName";
    public const string Project = "project";
    public const string Kind = "kind";
    public const string Source = "source";
    public const string Text = "text";
    public const string Preview = "preview";
    public const string Message = "message";
    public const string Signature = "signature";
    public const string Line = "line";
    public const string Column = "column";
    public const string Score = "score";
    public const string GroupBy = "groupBy";
    public const string Groups = "groups";
    public const string Key = "key";
    public const string SamplePath = "samplePath";
    public const string SampleName = "sampleName";
    public const string SampleText = "sampleText";
}

public sealed class SearchResultRow
{
    private readonly JsonNode _original;

    private SearchResultRow(JsonNode original, int index)
    {
        _original = original;
        Index = index;

        if (original is JsonObject obj)
        {
            Path = FirstString(obj, "path", "file", "fullPath");
            Name = FirstString(obj, "name", "displayName");
            FullName = FirstString(obj, "fullName", "name", "displayName");
            if (string.IsNullOrWhiteSpace(Name))
            {
                Name = FileName(Path);
            }

            Project = FirstString(obj, "project", "projectName", "projectUniqueName");
            // Text-source symbol hits carry "inferredKind" instead of "kind"; without the
            // fallback the kind filter/sort/group sees an empty string and drops every row.
            Kind = FirstString(obj, "kind", "inferredKind");
            Source = FirstString(obj, "source", "tool");
            Message = FirstString(obj, "message");
            Signature = FirstString(obj, SearchJsonNames.Signature);
            Text = FirstString(obj, "text", "message", "name", "fullName", SearchJsonNames.Signature, "preview", "displayName", "context");
            Preview = FirstString(obj, "preview", "context", "text", "message");
            Line = FirstInt(obj, "line", "startLine");
            Column = FirstInt(obj, "column", "startColumn");
            Score = FirstDouble(obj, "score", "scoreHint");
            return;
        }

        string text = NodeToString(original);
        Path = text;
        Name = FileName(text);
        FullName = text;
        Text = text;
        Preview = text;
    }

    public int Index { get; }

    public string Path { get; } = string.Empty;

    public string Name { get; } = string.Empty;

    public string FullName { get; } = string.Empty;

    public string Project { get; } = string.Empty;

    public string Kind { get; } = string.Empty;

    public string Source { get; } = string.Empty;

    public string Text { get; } = string.Empty;

    public string Preview { get; } = string.Empty;

    public string Message { get; } = string.Empty;

    public string Signature { get; } = string.Empty;

    public int? Line { get; }

    public int? Column { get; }

    public double? Score { get; }

    public static SearchResultRow FromJson(JsonNode? node, int index)
        => new((node?.DeepClone() ?? JsonValue.Create(string.Empty))!, index);

    public JsonNode ToJson()
        => _original.DeepClone();

    public bool Matches(SearchQueryOptions options)
        => ContainsPath(Path, options.Path)
            && ContainsPath(Project, options.Project)
            && Contains(Kind, options.Kind)
            && Contains(Source, options.Source)
            && ContainsAny(options.Text, Text, Preview, Message, Signature, Name, FullName, Path, Project, Kind, Source);

    public string GetSortText(string? sortBy)
    {
        return sortBy switch
        {
            "path" => Path,
            "file" => Path,
            "name" => Name,
            "fullname" => FullName,
            "project" => Project,
            "kind" => Kind,
            "source" => Source,
            "preview" => Preview,
            "message" => Message,
            SearchJsonNames.Signature => Signature,
            "text" => Text,
            _ => Text,
        };
    }

    public double GetSortNumber(string? sortBy)
    {
        return sortBy switch
        {
            "line" => Line ?? double.MaxValue,
            "column" => Column ?? double.MaxValue,
            "score" => Score ?? double.MinValue,
            "index" => Index,
            _ => double.MaxValue,
        };
    }

    private static bool Contains(string value, string? filter)
        => string.IsNullOrWhiteSpace(filter) || value.Contains(filter, StringComparison.OrdinalIgnoreCase);

    private static bool ContainsPath(string value, string? filter)
    {
        if (filter is null || string.IsNullOrWhiteSpace(filter))
        {
            return true;
        }

        return NormalizePathSeparators(value).Contains(NormalizePathSeparators(filter), StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizePathSeparators(string value)
        => value.Replace('\\', '/');

    private static bool ContainsAny(string? filter, params string[] values)
        => string.IsNullOrWhiteSpace(filter)
            || values.Any(value => !string.IsNullOrWhiteSpace(value)
                && value.Contains(filter, StringComparison.OrdinalIgnoreCase));

    private static string FirstString(JsonObject obj, params string[] names)
    {
        foreach (string name in names)
        {
            string value = NodeToString(obj[name]);
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return string.Empty;
    }

    private static int? FirstInt(JsonObject obj, params string[] names)
    {
        foreach (string name in names)
        {
            int? value = NodeToInt(obj[name]);
            if (value.HasValue)
            {
                return value.Value;
            }
        }

        return null;
    }

    private static double? FirstDouble(JsonObject obj, params string[] names)
    {
        foreach (string name in names)
        {
            double? value = NodeToDouble(obj[name]);
            if (value.HasValue)
            {
                return value.Value;
            }
        }

        return null;
    }

    private static string NodeToString(JsonNode? node)
    {
        if (node is null)
        {
            return string.Empty;
        }

        if (node is JsonValue value)
        {
            if (value.TryGetValue(out string? stringValue))
            {
                return stringValue ?? string.Empty;
            }

            if (value.TryGetValue(out int intValue))
            {
                return intValue.ToString(CultureInfo.InvariantCulture);
            }

            if (value.TryGetValue(out long longValue))
            {
                return longValue.ToString(CultureInfo.InvariantCulture);
            }

            if (value.TryGetValue(out double doubleValue))
            {
                return doubleValue.ToString(CultureInfo.InvariantCulture);
            }

            if (value.TryGetValue(out bool boolValue))
            {
                return boolValue.ToString(CultureInfo.InvariantCulture);
            }
        }

        return node.ToString();
    }

    private static int? NodeToInt(JsonNode? node)
    {
        if (node is not JsonValue value)
        {
            return null;
        }

        if (value.TryGetValue(out int intValue))
        {
            return intValue;
        }

        if (value.TryGetValue(out long longValue))
        {
            return longValue is >= int.MinValue and <= int.MaxValue ? (int)longValue : null;
        }

        string text = NodeToString(value);
        return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? parsed
            : null;
    }

    private static double? NodeToDouble(JsonNode? node)
    {
        if (node is not JsonValue value)
        {
            return null;
        }

        if (value.TryGetValue(out double doubleValue))
        {
            return doubleValue;
        }

        if (value.TryGetValue(out int intValue))
        {
            return intValue;
        }

        if (value.TryGetValue(out long longValue))
        {
            return longValue;
        }

        string text = NodeToString(value);
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? parsed
            : null;
    }

    private static string FileName(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        string normalized = path.Replace('\\', '/');
        string name = System.IO.Path.GetFileName(normalized);
        return string.IsNullOrWhiteSpace(name) ? normalized : name;
    }
}

public sealed class SearchResultCollection
{
    private SearchResultCollection(
        IReadOnlyList<SearchResultRow> rows,
        int sourceCount,
        int sourceTotalCount,
        bool sourceTruncated)
    {
        Rows = rows;
        SourceCount = sourceCount;
        SourceTotalCount = sourceTotalCount;
        SourceTruncated = sourceTruncated;
    }

    public IReadOnlyList<SearchResultRow> Rows { get; }

    public int SourceCount { get; }

    public int SourceTotalCount { get; }

    public bool SourceTruncated { get; }

    public static bool TryFromJsonObject(JsonObject source, string itemProperty, out SearchResultCollection collection)
    {
        if (source[itemProperty] is not JsonArray array)
        {
            collection = Empty;
            return false;
        }

        collection = FromJsonArray(array, source);
        return true;
    }

    public static SearchResultCollection FromJsonObject(JsonObject source, string itemProperty)
        => TryFromJsonObject(source, itemProperty, out SearchResultCollection collection)
            ? collection
            : Empty;

    public static SearchResultCollection FromJsonArray(JsonArray array, JsonObject? metadata = null)
    {
        List<SearchResultRow> rows = [];
        for (int index = 0; index < array.Count; index++)
        {
            rows.Add(SearchResultRow.FromJson(array[index], index));
        }

        int sourceCount = GetOptionalInt(metadata, SearchJsonNames.Count) ?? rows.Count;
        int sourceTotalCount = GetOptionalInt(metadata, SearchJsonNames.TotalCount)
            ?? GetOptionalInt(metadata, SearchJsonNames.TotalMatchCount)
            ?? sourceCount;
        bool sourceTruncated = GetOptionalBool(metadata, SearchJsonNames.Truncated)
            ?? sourceTotalCount > sourceCount;
        return new SearchResultCollection(rows, sourceCount, Math.Max(sourceTotalCount, sourceCount), sourceTruncated);
    }

    public JsonObject ToJsonObject(SearchQueryOptions options, JsonObject template, string itemProperty)
    {
        JsonObject result = (JsonObject)template.DeepClone();
        WriteTo(result, itemProperty, options);
        return result;
    }

    public void WriteTo(JsonObject target, string itemProperty, SearchQueryOptions options)
    {
        SearchResultCollection filtered = ApplyFilters(options).Sort(options.SortBy, options.SortDescending);
        SearchPage page = filtered.Page(options);
        JsonArray pageRows = [];
        foreach (SearchResultRow row in page.Rows)
        {
            pageRows.Add(row.ToJson());
        }

        target[itemProperty] = pageRows;
        target[SearchJsonNames.Count] = page.Rows.Count;
        target[SearchJsonNames.TotalCount] = SourceTotalCount;
        target[SearchJsonNames.FilteredCount] = filtered.Rows.Count;
        target[SearchJsonNames.ChunkIndex] = options.ChunkIndex;
        target[SearchJsonNames.ChunkSize] = options.ChunkSize;
        target[SearchJsonNames.ChunkCount] = page.ChunkCount;
        target[SearchJsonNames.ChunkStart] = page.ChunkStart;
        target[SearchJsonNames.ChunkEnd] = page.ChunkEnd;
        target[SearchJsonNames.HasMoreChunks] = page.HasMoreChunks;
        target[SearchJsonNames.Truncated] = SourceTruncated || page.ChunkCount > 1 || page.ChunkOutOfRange;
        target[SearchJsonNames.ChunkOutOfRange] = page.ChunkOutOfRange;

        if (!string.IsNullOrWhiteSpace(options.SortBy))
        {
            target[SearchJsonNames.SortBy] = options.SortBy;
            target[SearchJsonNames.SortDirection] = options.SortDescending ? "desc" : "asc";
        }

        JsonArray groups = filtered.GroupBy(options.GroupBy);
        if (groups.Count > 0)
        {
            target[SearchJsonNames.Groups] = groups;
        }
        else
        {
            target.Remove(SearchJsonNames.Groups);
        }
    }

    public SearchResultCollection ApplyFilters(SearchQueryOptions options)
        => WithFilteredRows(row => row.Matches(options));

    public SearchResultCollection WithPath(string? path)
        => string.IsNullOrWhiteSpace(path) ? this : WithFilteredRows(row => row.Matches(new SearchQueryOptions(0, 0, null, false, path, null, null, null, null)));

    public SearchResultCollection WithProject(string? project)
        => string.IsNullOrWhiteSpace(project) ? this : WithFilteredRows(row => row.Matches(new SearchQueryOptions(0, 0, null, false, null, project, null, null, null)));

    public SearchResultCollection WithKind(string? kind)
        => string.IsNullOrWhiteSpace(kind) ? this : WithFilteredRows(row => row.Matches(new SearchQueryOptions(0, 0, null, false, null, null, kind, null, null)));

    public SearchResultCollection WithSource(string? source)
        => string.IsNullOrWhiteSpace(source) ? this : WithFilteredRows(row => row.Matches(new SearchQueryOptions(0, 0, null, false, null, null, null, source, null)));

    public SearchResultCollection WithText(string? text)
        => string.IsNullOrWhiteSpace(text) ? this : WithFilteredRows(row => row.Matches(new SearchQueryOptions(0, 0, null, false, null, null, null, null, text)));

    public JsonArray GroupBy(string? groupBy)
    {
        string? normalized = SearchQueryOptions.NormalizeGroupBy(groupBy);
        JsonArray groups = [];
        if (normalized is null)
        {
            return groups;
        }

        string normalizedGroupBy = normalized;
        foreach (IGrouping<string, SearchResultRow> group in Rows.GroupBy(row => GroupKey(row, normalizedGroupBy), StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase))
        {
            SearchResultRow sample = group.First();
            groups.Add(new JsonObject
            {
                [SearchJsonNames.GroupBy] = normalized,
                [SearchJsonNames.Key] = group.Key,
                [SearchJsonNames.Count] = group.Count(),
                [SearchJsonNames.SamplePath] = sample.Path,
                [SearchJsonNames.SampleName] = sample.Name,
                [SearchJsonNames.SampleText] = sample.Preview,
            });
        }

        return groups;
    }

    public SearchResultCollection Sort(string? sortBy, bool descending)
    {
        string? normalized = SearchQueryOptions.NormalizeSortField(sortBy);
        if (normalized is null)
        {
            return this;
        }

        string normalizedSort = normalized;
        IEnumerable<SearchResultRow> sorted = normalizedSort is "line" or "column" or "score" or "index"
            ? SortByNumber(normalizedSort, descending)
            : SortByText(normalizedSort, descending);
        return new SearchResultCollection([.. sorted], SourceCount, SourceTotalCount, SourceTruncated);
    }

    public SearchPage Page(SearchQueryOptions options)
    {
        int filteredCount = Rows.Count;
        int chunkCount = CalculateChunkCount(filteredCount, options.ChunkSize);
        int chunkStart = CalculateChunkStart(filteredCount, options.ChunkSize, options.ChunkIndex);
        bool chunkOutOfRange = options.ChunkSize > 0 && filteredCount > 0 && chunkStart >= filteredCount;
        IReadOnlyList<SearchResultRow> pageRows = chunkOutOfRange
            ? []
            : options.ChunkSize == 0
                ? Rows
                : [.. Rows.Skip(chunkStart).Take(Math.Min(options.ChunkSize, Math.Max(0, filteredCount - chunkStart)))];

        return new SearchPage(
            pageRows,
            chunkCount,
            chunkOutOfRange ? filteredCount : Math.Min(chunkStart, filteredCount),
            chunkOutOfRange ? filteredCount : Math.Min(chunkStart + pageRows.Count, filteredCount),
            options.ChunkSize > 0 && options.ChunkIndex + 1 < chunkCount,
            chunkOutOfRange);
    }

    private SearchResultCollection WithFilteredRows(Func<SearchResultRow, bool> predicate)
        => new([.. Rows.Where(predicate)], SourceCount, SourceTotalCount, SourceTruncated);

    private IEnumerable<SearchResultRow> SortByNumber(string sortBy, bool descending)
        => descending
            ? Rows.OrderByDescending(row => row.GetSortNumber(sortBy)).ThenBy(row => row.Index)
            : Rows.OrderBy(row => row.GetSortNumber(sortBy)).ThenBy(row => row.Index);

    private IEnumerable<SearchResultRow> SortByText(string sortBy, bool descending)
        => descending
            ? Rows.OrderByDescending(row => row.GetSortText(sortBy), StringComparer.OrdinalIgnoreCase).ThenBy(row => row.Index)
            : Rows.OrderBy(row => row.GetSortText(sortBy), StringComparer.OrdinalIgnoreCase).ThenBy(row => row.Index);

    private static string GroupKey(SearchResultRow row, string groupBy)
    {
        string key = groupBy switch
        {
            "path" => row.Path,
            "name" => row.Name,
            "project" => row.Project,
            "kind" => row.Kind,
            "source" => row.Source,
            _ => string.Empty,
        };

        return string.IsNullOrWhiteSpace(key) ? "(blank)" : key;
    }

    private static int CalculateChunkCount(int count, int chunkSize)
    {
        if (count == 0)
        {
            return 0;
        }

        return chunkSize == 0
            ? 1
            : (int)Math.Ceiling(count / (double)chunkSize);
    }

    private static int CalculateChunkStart(int count, int chunkSize, int chunkIndex)
    {
        if (count == 0 || chunkSize == 0)
        {
            return 0;
        }

        try
        {
            long start = checked((long)chunkIndex * chunkSize);
            return start > count ? count : (int)start;
        }
        catch (OverflowException)
        {
            return count;
        }
    }

    private static int? GetOptionalInt(JsonObject? source, string propertyName)
    {
        if (source?[propertyName] is not JsonValue valueNode)
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

    private static bool? GetOptionalBool(JsonObject? source, string propertyName)
    {
        if (source?[propertyName] is not JsonValue valueNode)
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

    private static SearchResultCollection Empty { get; } = new([], 0, 0, false);
}

public sealed class SearchQueryOptions(
    int chunkSize,
    int chunkIndex,
    string? sortBy,
    bool sortDescending,
    string? path,
    string? project,
    string? kind,
    string? source,
    string? text,
    string? groupBy = null)
{
    public int ChunkSize { get; } = chunkSize;

    public int ChunkIndex { get; } = chunkIndex;

    public string? SortBy { get; } = NormalizeSortField(sortBy);

    public bool SortDescending { get; } = sortDescending;

    public string? Path { get; } = NullIfWhiteSpace(path);

    public string? Project { get; } = NullIfWhiteSpace(project);

    public string? Kind { get; } = NullIfWhiteSpace(kind);

    public string? Source { get; } = NullIfWhiteSpace(source);

    public string? Text { get; } = NullIfWhiteSpace(text);

    public string? GroupBy { get; } = NormalizeGroupBy(groupBy);

    public static SearchQueryOptions FromJsonObject(JsonObject? source, int defaultChunkSize, bool includePathFilter = true)
    {
        return new SearchQueryOptions(
            GetOptionalNonNegativeInt(source, "chunk_size") ?? defaultChunkSize,
            GetOptionalNonNegativeInt(source, "chunk_index") ?? 0,
            GetOptionalString(source, "sort_by"),
            IsDescendingSort(GetOptionalString(source, "sort_direction")),
            includePathFilter ? GetOptionalString(source, SearchJsonNames.Path) : null,
            GetOptionalString(source, SearchJsonNames.Project),
            GetOptionalString(source, SearchJsonNames.Kind),
            GetOptionalString(source, SearchJsonNames.Source),
            GetOptionalString(source, SearchJsonNames.Text),
            GetOptionalString(source, "group_by") ?? GetOptionalString(source, "group-by"));
    }

    public static string? NormalizeSortField(string? sortBy)
    {
        string? normalized = NullIfWhiteSpace(sortBy)?.ToLowerInvariant();
        return normalized switch
        {
            "file" => "path",
            "path" => "path",
            "name" => "name",
            "project" => "project",
            "kind" => "kind",
            "source" => "source",
            "line" => "line",
            "column" => "column",
            "score" => "score",
            "scorehint" => "score",
            "text" => "text",
            "preview" => "preview",
            "message" => "message",
            SearchJsonNames.Signature => SearchJsonNames.Signature,
            "fullname" => "fullname",
            "index" => "index",
            _ => null,
        };
    }

    public static string? NormalizeGroupBy(string? groupBy)
    {
        string? normalized = NullIfWhiteSpace(groupBy)?.ToLowerInvariant();
        return normalized switch
        {
            "file" => "path",
            "path" => "path",
            "name" => "name",
            "project" => "project",
            "kind" => "kind",
            "source" => "source",
            _ => null,
        };
    }

    private static int? GetOptionalNonNegativeInt(JsonObject? source, string propertyName)
    {
        if (source?[propertyName] is not JsonValue valueNode)
        {
            return null;
        }

        if (valueNode.TryGetValue(out int value))
        {
            return Math.Max(0, value);
        }

        return int.TryParse(valueNode.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value)
            ? Math.Max(0, value)
            : null;
    }

    private static string? GetOptionalString(JsonObject? source, string propertyName)
    {
        if (source?[propertyName] is not JsonValue valueNode)
        {
            return null;
        }

        if (valueNode.TryGetValue(out string? value))
        {
            return NullIfWhiteSpace(value);
        }

        return NullIfWhiteSpace(valueNode.ToString());
    }

    private static bool IsDescendingSort(string? sortDirection)
        => string.Equals(sortDirection, "desc", StringComparison.OrdinalIgnoreCase)
            || string.Equals(sortDirection, "descending", StringComparison.OrdinalIgnoreCase);

    private static string? NullIfWhiteSpace(string? value)
    {
        return string.IsNullOrWhiteSpace(value)
            ? null
            : value!.Trim();
    }
}

public sealed class SearchPage(
    IReadOnlyList<SearchResultRow> rows,
    int chunkCount,
    int chunkStart,
    int chunkEnd,
    bool hasMoreChunks,
    bool chunkOutOfRange)
{
    public IReadOnlyList<SearchResultRow> Rows { get; } = rows;

    public int ChunkCount { get; } = chunkCount;

    public int ChunkStart { get; } = chunkStart;

    public int ChunkEnd { get; } = chunkEnd;

    public bool HasMoreChunks { get; } = hasMoreChunks;

    public bool ChunkOutOfRange { get; } = chunkOutOfRange;
}
