using System.Text.Json.Nodes;
using VsIdeBridgeService.SystemTools;
using Xunit;

namespace VsIdeBridgeService.Tests;

public sealed class MemoryToolTests : IDisposable
{
    private readonly string _codexHome;
    private readonly string? _previousCodexHome;

    public MemoryToolTests()
    {
        _previousCodexHome = Environment.GetEnvironmentVariable("CODEX_HOME");
        _codexHome = Path.Combine(Path.GetTempPath(), "VsIdeBridgeMemoryTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(_codexHome, "memories"));
        Environment.SetEnvironmentVariable("CODEX_HOME", _codexHome);
    }

    [Fact]
    public async Task SearchAsyncMatchesRankedMultiTermQueries()
    {
        string memories = Path.Combine(_codexHome, "memories");
        File.WriteAllText(Path.Combine(memories, "memory_summary.md"), string.Join(Environment.NewLine,
        [
            "A harmless line about tabs.",
            "SuperSlicer warning cleanup through the bridge: warnings, BP1045, BP1014, and BP1015.",
            "Bridge memory lookup inside the bound solution: memory_search, memory_read, list_tool_categories.",
        ]));
        File.WriteAllText(Path.Combine(memories, "MEMORY.md"), "Bridge only without the warning code.");

        JsonObject structured = await SearchAsync("SuperSlicer warning cleanup bridge BP1045");
        JsonArray results = Assert.IsType<JsonArray>(structured["results"]);
        JsonObject first = Assert.IsType<JsonObject>(results[0]);

        Assert.True(structured["count"]!.GetValue<int>() >= 1);
        Assert.True(structured["matchCount"]!.GetValue<int>() >= 1);
        Assert.Equal("memory_summary.md", first["path"]!.GetValue<string>());
        Assert.Equal(2, first["line"]!.GetValue<int>());
        Assert.Equal("terms", first["matchType"]!.GetValue<string>());
        Assert.Equal(5, first["matchedTerms"]!.GetValue<int>());
    }

    [Fact]
    public async Task SearchAsyncRanksExactPhraseAheadOfTermMatches()
    {
        string memories = Path.Combine(_codexHome, "memories");
        File.WriteAllText(Path.Combine(memories, "memory_summary.md"), string.Join(Environment.NewLine,
        [
            "bridge memory phrase without the exact leading term",
            "exact bridge memory phrase",
        ]));

        JsonObject structured = await SearchAsync("exact bridge memory phrase");
        JsonArray results = Assert.IsType<JsonArray>(structured["results"]);
        JsonObject first = Assert.IsType<JsonObject>(results[0]);

        Assert.Equal(2, first["line"]!.GetValue<int>());
        Assert.Equal("exact", first["matchType"]!.GetValue<string>());
    }

    public void Dispose()
    {
        Environment.SetEnvironmentVariable("CODEX_HOME", _previousCodexHome);
        if (Directory.Exists(_codexHome))
        {
            Directory.Delete(_codexHome, recursive: true);
        }
    }

    private static async Task<JsonObject> SearchAsync(string query)
    {
        JsonObject result = Assert.IsType<JsonObject>(await MemoryTool.SearchAsync(null, new JsonObject
        {
            ["query"] = query,
            ["max_results"] = 5,
            ["context_lines"] = 0,
            ["include_rollouts"] = false,
        }, new BridgeConnection([])));

        Assert.False(result["isError"]!.GetValue<bool>());
        return Assert.IsType<JsonObject>(result["structuredContent"]);
    }
}
