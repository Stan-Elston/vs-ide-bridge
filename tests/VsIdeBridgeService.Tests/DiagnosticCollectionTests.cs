using System.Text.Json.Nodes;
using VsIdeBridge.Diagnostics;
using static VsIdeBridgeService.Tests.DiagnosticCollectionTestPaging;
using static VsIdeBridgeService.Tests.DiagnosticCollectionTestRows;
using Xunit;

namespace VsIdeBridgeService.Tests;

internal static class DiagnosticCollectionTestPaging
{
    public const int DefaultChunkSize = 10;
    public const int AllRowsChunkSize = 0;
    public const int SingleRowChunkSize = 1;
    public const int FirstChunkIndex = 0;
    public const int SecondChunkIndex = 1;
    public const int OutOfRangeChunkIndex = 99;
    public const int DiagnosticRowCount = 3;
    public const int WarningSeverityCount = 2;
    public const int SecondChunkEnd = 2;
}

internal static class DiagnosticCollectionTestRows
{
    public const int CacheWarningLine = 25;
    public const int CacheWarningColumn = 7;
    public const int CompilerErrorLine = 10;
    public const int CompilerErrorColumn = 4;
    public const int AnalyzerWarningLine = 1;
    public const int AnalyzerWarningColumn = 2;
}

public sealed class DiagnosticCollectionTests
{
    [Fact]
    public void FiltersBySeverityCodeProjectPathAndText()
    {
        DiagnosticCollection collection = DiagnosticCollection.FromJsonObject(CreateBucket());
        DiagnosticQueryOptions options = new(
            DefaultChunkSize,
            FirstChunkIndex,
            null,
            false,
            "warning",
            "BP10",
            "Service",
            "Diagnostics",
            "leak",
            null);

        JsonObject result = collection.ToJsonObject(options, CreateBucket());
        JsonArray rows = Assert.IsType<JsonArray>(result[DiagnosticJsonNames.Rows]);
        JsonObject row = Assert.IsType<JsonObject>(rows.Single());

        Assert.Equal(SingleRowChunkSize, result[DiagnosticJsonNames.Count]!.GetValue<int>());
        Assert.Equal(SingleRowChunkSize, result[DiagnosticJsonNames.FilteredCount]!.GetValue<int>());
        Assert.Equal("BP1044", row[DiagnosticJsonNames.Code]!.GetValue<string>());
        Assert.Equal("keep-me", row["guidance"]!.GetValue<string>());
    }

    [Theory]
    [InlineData("code", "BP1044", "CS1001")]
    [InlineData("project", "VsIdeBridge", "VsIdeBridgeService")]
    [InlineData("file", "src/Analyzer.cs", "src/Program.cs")]
    [InlineData("path", "src/Analyzer.cs", "src/Program.cs")]
    [InlineData("line", "1", "25")]
    [InlineData("column", "2", "7")]
    [InlineData("message", "Compiler failure", "Potential leak in diagnostics cache")]
    [InlineData("source", "Analyzer", "Compiler")]
    [InlineData("tool", "bridge", "roslyn")]
    public void SortsAscendingAndDescending(string sortBy, string ascendingFirst, string descendingFirst)
    {
        DiagnosticCollection collection = DiagnosticCollection.FromJsonObject(CreateBucket());

        JsonObject ascending = collection.ToJsonObject(CreateSortOptions(sortBy, sortDescending: false), CreateBucket());
        JsonObject descending = collection.ToJsonObject(CreateSortOptions(sortBy, sortDescending: true), CreateBucket());

        Assert.Equal(ascendingFirst, GetSortValue(ascending, sortBy));
        Assert.Equal(descendingFirst, GetSortValue(descending, sortBy));
        Assert.Equal(sortBy is "path" ? "file" : sortBy, ascending[DiagnosticJsonNames.SortBy]!.GetValue<string>());
        Assert.Equal("desc", descending[DiagnosticJsonNames.SortDirection]!.GetValue<string>());
    }

    [Fact]
    public void PagesRowsAndKeepsChunkMetadataStable()
    {
        DiagnosticCollection collection = DiagnosticCollection.FromJsonObject(CreateBucket());
        JsonObject firstPage = collection.ToJsonObject(new(SingleRowChunkSize, FirstChunkIndex, "code", false, null, null, null, null, null, null), CreateBucket());
        JsonObject secondPage = collection.ToJsonObject(new(SingleRowChunkSize, SecondChunkIndex, "code", false, null, null, null, null, null, null), CreateBucket());

        Assert.Equal(SingleRowChunkSize, firstPage[DiagnosticJsonNames.Count]!.GetValue<int>());
        Assert.Equal(DiagnosticRowCount, firstPage[DiagnosticJsonNames.FilteredCount]!.GetValue<int>());
        Assert.Equal(DiagnosticRowCount, firstPage[DiagnosticJsonNames.ChunkCount]!.GetValue<int>());
        Assert.Equal(FirstChunkIndex, firstPage[DiagnosticJsonNames.ChunkStart]!.GetValue<int>());
        Assert.Equal(SingleRowChunkSize, firstPage[DiagnosticJsonNames.ChunkEnd]!.GetValue<int>());
        Assert.True(firstPage[DiagnosticJsonNames.HasMoreChunks]!.GetValue<bool>());

        Assert.Equal(SecondChunkIndex, secondPage[DiagnosticJsonNames.ChunkStart]!.GetValue<int>());
        Assert.Equal(SecondChunkEnd, secondPage[DiagnosticJsonNames.ChunkEnd]!.GetValue<int>());
    }

    [Fact]
    public void FromJsonObjectFiltersByFileAndChunkSize()
    {
        DiagnosticCollection collection = DiagnosticCollection.FromJsonObject(CreateBucket());
        JsonObject args = new()
        {
            ["file"] = "Analyzer.cs",
            ["chunk_size"] = SingleRowChunkSize,
        };

        JsonObject result = collection.ToJsonObject(DiagnosticQueryOptions.FromJsonObject(args, DefaultChunkSize), CreateBucket());
        JsonArray rows = Assert.IsType<JsonArray>(result[DiagnosticJsonNames.Rows]);
        JsonObject row = Assert.IsType<JsonObject>(rows.Single());

        Assert.Equal(SingleRowChunkSize, result[DiagnosticJsonNames.Count]!.GetValue<int>());
        Assert.Equal(SingleRowChunkSize, result[DiagnosticJsonNames.ChunkSize]!.GetValue<int>());
        Assert.Equal("CA2000", row[DiagnosticJsonNames.Code]!.GetValue<string>());
    }

    [Fact]
    public void VisualStudioAnalyzerRowsAreMarkedAnalyzerLimited()
    {
        JsonObject bucket = new()
        {
            [DiagnosticJsonNames.Rows] = new JsonArray
            {
                new JsonObject
                {
                    [DiagnosticJsonNames.Severity] = "Message",
                    [DiagnosticJsonNames.Code] = "VCR003",
                    [DiagnosticJsonNames.File] = "src/Widget.cpp",
                    [DiagnosticJsonNames.Message] = "Function definition missing.",
                },
            },
        };

        JsonObject result = DiagnosticCollection.FromJsonObject(bucket)
            .ToJsonObject(new(DefaultChunkSize, FirstChunkIndex, null, false, null, null, null, null, null, null), bucket);
        JsonObject row = Assert.IsType<JsonObject>(Assert.IsType<JsonArray>(result[DiagnosticJsonNames.Rows]).Single());

        Assert.True(row["analyzerLimited"]!.GetValue<bool>());
        Assert.Equal("visual-studio-analyzer-limited", row["authority"]!.GetValue<string>());
        Assert.Contains("modern C++", row["guidance"]!.GetValue<string>());
    }

    [Fact]
    public void FilteringTruncatedSourceReportsKnownMatchesOnly()
    {
        JsonObject bucket = CreateBucket();
        bucket[DiagnosticJsonNames.TotalCount] = 100;
        bucket[DiagnosticJsonNames.Truncated] = true;
        DiagnosticCollection collection = DiagnosticCollection.FromJsonObject(bucket);
        DiagnosticQueryOptions options = new(
            DefaultChunkSize,
            FirstChunkIndex,
            null,
            false,
            null,
            "BP1044",
            null,
            null,
            null,
            null);

        JsonObject result = collection.ToJsonObject(options, bucket);

        Assert.Equal(SingleRowChunkSize, result[DiagnosticJsonNames.Count]!.GetValue<int>());
        Assert.Equal(SingleRowChunkSize, result[DiagnosticJsonNames.FilteredCount]!.GetValue<int>());
        Assert.True(result[DiagnosticJsonNames.Truncated]!.GetValue<bool>());
    }

    [Fact]
    public void ChunkSizeZeroReturnsAllRowsAsOneChunk()
    {
        DiagnosticCollection collection = DiagnosticCollection.FromJsonObject(CreateBucket());
        JsonObject result = collection.ToJsonObject(new(AllRowsChunkSize, FirstChunkIndex, "code", false, null, null, null, null, null, null), CreateBucket());

        Assert.Equal(DiagnosticRowCount, result[DiagnosticJsonNames.Count]!.GetValue<int>());
        Assert.Equal(DiagnosticRowCount, result[DiagnosticJsonNames.FilteredCount]!.GetValue<int>());
        Assert.Equal(SingleRowChunkSize, result[DiagnosticJsonNames.ChunkCount]!.GetValue<int>());
        Assert.Equal(AllRowsChunkSize, result[DiagnosticJsonNames.ChunkSize]!.GetValue<int>());
        Assert.False(result[DiagnosticJsonNames.HasMoreChunks]!.GetValue<bool>());
    }

    [Fact]
    public void OutOfRangeChunkReturnsEmptyRowsAndFlag()
    {
        DiagnosticCollection collection = DiagnosticCollection.FromJsonObject(CreateBucket());
        JsonObject result = collection.ToJsonObject(new(SingleRowChunkSize, OutOfRangeChunkIndex, "code", false, null, null, null, null, null, null), CreateBucket());

        JsonArray rows = Assert.IsType<JsonArray>(result[DiagnosticJsonNames.Rows]);
        Assert.Empty(rows);
        Assert.Equal(DiagnosticRowCount, result[DiagnosticJsonNames.ChunkStart]!.GetValue<int>());
        Assert.Equal(DiagnosticRowCount, result[DiagnosticJsonNames.ChunkEnd]!.GetValue<int>());
        Assert.True(result[DiagnosticJsonNames.ChunkOutOfRange]!.GetValue<bool>());
        Assert.True(result[DiagnosticJsonNames.Truncated]!.GetValue<bool>());
    }

    [Fact]
    public void EmptyRowsSerializeWithZeroCounts()
    {
        JsonObject emptyBucket = new()
        {
            [DiagnosticJsonNames.Count] = 0,
            [DiagnosticJsonNames.TotalCount] = 0,
            [DiagnosticJsonNames.Rows] = new JsonArray(),
        };

        JsonObject result = DiagnosticCollection.FromJsonObject(emptyBucket)
            .ToJsonObject(new(DefaultChunkSize, FirstChunkIndex, null, false, null, null, null, null, null, null), emptyBucket);

        Assert.Equal(AllRowsChunkSize, result[DiagnosticJsonNames.Count]!.GetValue<int>());
        Assert.Equal(AllRowsChunkSize, result[DiagnosticJsonNames.FilteredCount]!.GetValue<int>());
        Assert.Equal(AllRowsChunkSize, result[DiagnosticJsonNames.ChunkCount]!.GetValue<int>());
        Assert.False(result[DiagnosticJsonNames.HasMoreChunks]!.GetValue<bool>());
    }

    [Fact]
    public void SerializationKeepsToolResponseFieldsAndGroups()
    {
        DiagnosticCollection collection = DiagnosticCollection.FromJsonObject(CreateBucket());
        JsonObject result = collection.ToJsonObject(new(WarningSeverityCount, FirstChunkIndex, "code", false, null, null, null, null, null, "code"), CreateBucket());

        Assert.NotNull(result[DiagnosticJsonNames.Rows]);
        Assert.NotNull(result[DiagnosticJsonNames.Count]);
        Assert.NotNull(result[DiagnosticJsonNames.TotalCount]);
        Assert.NotNull(result[DiagnosticJsonNames.FilteredCount]);
        Assert.NotNull(result[DiagnosticJsonNames.ChunkIndex]);
        Assert.NotNull(result[DiagnosticJsonNames.ChunkSize]);
        Assert.NotNull(result[DiagnosticJsonNames.ChunkCount]);
        Assert.NotNull(result[DiagnosticJsonNames.ChunkStart]);
        Assert.NotNull(result[DiagnosticJsonNames.ChunkEnd]);
        Assert.NotNull(result[DiagnosticJsonNames.HasMoreChunks]);
        Assert.NotNull(result[DiagnosticJsonNames.Truncated]);
        Assert.Equal("code", result[DiagnosticJsonNames.SortBy]!.GetValue<string>());
        Assert.Equal("asc", result[DiagnosticJsonNames.SortDirection]!.GetValue<string>());

        JsonArray groups = Assert.IsType<JsonArray>(result[DiagnosticJsonNames.Groups]);
        Assert.Contains(groups.OfType<JsonObject>(), group => group[DiagnosticJsonNames.Key]?.GetValue<string>() == "BP1044");
    }

    [Fact]
    public void GroupedReadWithFiniteChunkSizeGroupsAllRows()
    {
        DiagnosticCollection collection = DiagnosticCollection.FromJsonObject(CreateBucket());
        JsonObject result = collection.ToJsonObject(new(SingleRowChunkSize, FirstChunkIndex, null, false, null, null, null, null, null, "code"), CreateBucket());

        // Grouping must cover every filtered row; a finite chunk_size only paginates the groups.
        Assert.Equal(DiagnosticRowCount, result["groupCount"]!.GetValue<int>());
        Assert.Equal(DiagnosticRowCount, result["diagnosticCount"]!.GetValue<int>());
        Assert.Equal(DiagnosticRowCount, result[DiagnosticJsonNames.ChunkCount]!.GetValue<int>());
        JsonArray groups = Assert.IsType<JsonArray>(result[DiagnosticJsonNames.Groups]);
        Assert.Single(groups);
    }

    [Fact]
    public void GroupedWriteFromTruncatedPayloadKeepsUpstreamGroups()
    {
        // Simulates a payload whose rows array was already row-chunked (1 of 10 rows) while
        // its groups were computed upstream over the full filtered set. Regrouping the chunk
        // would shrink every count to the current chunk; the upstream summary must win.
        JsonObject truncatedBucket = CreateBucket();
        JsonArray rows = (JsonArray)truncatedBucket[DiagnosticJsonNames.Rows]!;
        while (rows.Count > 1)
        {
            rows.RemoveAt(rows.Count - 1);
        }

        truncatedBucket[DiagnosticJsonNames.Count] = 1;
        truncatedBucket[DiagnosticJsonNames.TotalCount] = 10;
        truncatedBucket[DiagnosticJsonNames.Groups] = new JsonArray
        {
            new JsonObject { [DiagnosticJsonNames.Key] = "BP1002", [DiagnosticJsonNames.Count] = 7 },
            new JsonObject { [DiagnosticJsonNames.Key] = "C4244", [DiagnosticJsonNames.Count] = 3 },
        };

        JsonObject result = DiagnosticCollection.FromJsonObject(truncatedBucket)
            .ToJsonObject(new(DefaultChunkSize, FirstChunkIndex, null, false, null, null, null, null, null, "code"), truncatedBucket);

        JsonArray groups = Assert.IsType<JsonArray>(result[DiagnosticJsonNames.Groups]);
        Assert.Equal(2, groups.Count);
        Assert.Equal("BP1002", groups[0]![DiagnosticJsonNames.Key]!.GetValue<string>());
        Assert.Equal(7, groups[0]![DiagnosticJsonNames.Count]!.GetValue<int>());
        Assert.Equal(2, result["groupCount"]!.GetValue<int>());
    }

    [Fact]
    public void GroupedBridgeCommandArgsRequestFullRowSet()
    {
        JsonObject groupedArgs = new()
        {
            ["group_by"] = "code",
            ["chunk_size"] = 20,
            ["chunk_index"] = 1,
        };
        string grouped = ToolCatalog.BuildDiagnosticRowsCommandArgs(groupedArgs, "warning");
        Assert.Contains("--chunk-size", grouped);
        Assert.DoesNotContain("20", grouped);
        Assert.DoesNotContain("--chunk-index", grouped);

        JsonObject ungroupedArgs = new()
        {
            ["chunk_size"] = 20,
            ["chunk_index"] = 1,
        };
        string ungrouped = ToolCatalog.BuildDiagnosticRowsCommandArgs(ungroupedArgs, "warning");
        Assert.Contains("--chunk-size", ungrouped);
        Assert.Contains("20", ungrouped);
        Assert.Contains("--chunk-index", ungrouped);
    }

    private static string GetSortValue(JsonObject result, string sortBy)
    {
        JsonArray rows = Assert.IsType<JsonArray>(result[DiagnosticJsonNames.Rows]);
        JsonObject first = Assert.IsType<JsonObject>(rows.First());
        return sortBy switch
        {
            "path" => first[DiagnosticJsonNames.File]!.GetValue<string>(),
            "line" => first[DiagnosticJsonNames.Line]!.GetValue<int>().ToString(),
            "column" => first[DiagnosticJsonNames.Column]!.GetValue<int>().ToString(),
            _ => first[sortBy]!.GetValue<string>(),
        };
    }

    private static DiagnosticQueryOptions CreateSortOptions(string sortBy, bool sortDescending)
    {
        JsonObject args = new()
        {
            ["chunk_size"] = DefaultChunkSize,
            ["chunk_index"] = FirstChunkIndex,
            ["sort_by"] = sortBy,
            ["sort_direction"] = sortDescending ? "desc" : "asc",
        };

        return DiagnosticQueryOptions.FromJsonObject(args, DefaultChunkSize);
    }

    private static JsonObject CreateBucket()
    {
        return new JsonObject
        {
            [DiagnosticJsonNames.Count] = DiagnosticRowCount,
            [DiagnosticJsonNames.TotalCount] = DiagnosticRowCount,
            ["severityCounts"] = new JsonObject
            {
                ["Error"] = 1,
                ["Warning"] = WarningSeverityCount,
                ["Message"] = 0,
            },
            [DiagnosticJsonNames.Rows] = new JsonArray
            {
                new JsonObject
                {
                    [DiagnosticJsonNames.Severity] = "Warning",
                    [DiagnosticJsonNames.Code] = "BP1044",
                    [DiagnosticJsonNames.Project] = "VsIdeBridgeService",
                    [DiagnosticJsonNames.File] = "src/Diagnostics/Cache.cs",
                    [DiagnosticJsonNames.Path] = "src/Diagnostics/Cache.cs",
                    [DiagnosticJsonNames.Line] = CacheWarningLine,
                    [DiagnosticJsonNames.Column] = CacheWarningColumn,
                    [DiagnosticJsonNames.Message] = "Potential leak in diagnostics cache",
                    [DiagnosticJsonNames.Source] = "Bridge",
                    [DiagnosticJsonNames.Tool] = "bridge",
                    ["guidance"] = "keep-me",
                },
                new JsonObject
                {
                    [DiagnosticJsonNames.Severity] = "Error",
                    [DiagnosticJsonNames.Code] = "CS1001",
                    [DiagnosticJsonNames.Project] = "VsIdeBridge",
                    [DiagnosticJsonNames.File] = "src/Program.cs",
                    [DiagnosticJsonNames.Path] = "src/Program.cs",
                    [DiagnosticJsonNames.Line] = CompilerErrorLine,
                    [DiagnosticJsonNames.Column] = CompilerErrorColumn,
                    [DiagnosticJsonNames.Message] = "Compiler failure",
                    [DiagnosticJsonNames.Source] = "Compiler",
                    [DiagnosticJsonNames.Tool] = "compiler",
                },
                new JsonObject
                {
                    [DiagnosticJsonNames.Severity] = "Warning",
                    [DiagnosticJsonNames.Code] = "CA2000",
                    [DiagnosticJsonNames.Project] = "VsIdeBridgeService",
                    [DiagnosticJsonNames.File] = "src/Analyzer.cs",
                    [DiagnosticJsonNames.Path] = "src/Analyzer.cs",
                    [DiagnosticJsonNames.Line] = AnalyzerWarningLine,
                    [DiagnosticJsonNames.Column] = AnalyzerWarningColumn,
                    [DiagnosticJsonNames.Message] = "Dispose analyzer warning",
                    [DiagnosticJsonNames.Source] = "Analyzer",
                    [DiagnosticJsonNames.Tool] = "roslyn",
                },
            },
        };
    }
}
