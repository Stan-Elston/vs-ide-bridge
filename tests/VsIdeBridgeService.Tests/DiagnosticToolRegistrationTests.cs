using System.Text.Json.Nodes;
using VsIdeBridge.Shared;
using Xunit;

namespace VsIdeBridgeService.Tests;

public sealed class DiagnosticToolRegistrationTests
{
    private static readonly string[] SharedDiagnosticParameters =
    [
        "severity",
        "wait_for_intellisense",
        "quick",
        "refresh",
        "max",
        "chunk_size",
        "chunk_index",
        "sort_by",
        "sort_direction",
        "code",
        "project",
        "path",
        "file",
        "text",
        "group_by",
    ];

    [Theory]
    [InlineData("errors", "Optional diagnostic code prefix filter.")]
    [InlineData("warnings", "Optional warning code prefix filter.")]
    [InlineData("messages", "Optional message code prefix filter.")]
    public void DiagnosticRowToolsExposeSharedFilterSchema(string toolName, string codeDescription)
    {
        ToolDefinition definition = GetDefinition(toolName);
        JsonObject properties = GetSchemaProperties(definition);

        foreach (string parameter in SharedDiagnosticParameters)
        {
            Assert.True(properties.ContainsKey(parameter), $"{toolName} should expose '{parameter}'.");
        }

        Assert.Equal(codeDescription, properties["code"]!["description"]!.GetValue<string>());
        Assert.Equal("Rows per returned chunk (default 10, or max when set). Set 0 to return all filtered rows.",
            properties["chunk_size"]!["description"]!.GetValue<string>());
    }

    [Fact]
    public void WarningsToolUsesCatalogDefinitionMetadata()
    {
        ToolDefinition definition = GetDefinition("warnings");

        Assert.True(definition.ReadOnly);
        Assert.Equal("diagnostics", definition.Category);
        Assert.Equal("warnings", definition.BridgeCommand);
        Assert.Equal("Error List Warnings", definition.Title);
        Assert.Contains("list_warnings", definition.Aliases);
        Assert.Contains("diagnostic_warnings", definition.Aliases);
        Assert.NotNull(definition.SearchHints);
    }

    [Fact]
    public void DiagnosticTimeoutProfilesCoverAllErrorListCommands()
    {
        Assert.Equal(BridgeConnection.ToolTimeoutProfile.Interactive, BridgeConnectionArgs.SelectTimeoutProfile("errors"));
        Assert.Equal(BridgeConnection.ToolTimeoutProfile.Interactive, BridgeConnectionArgs.SelectTimeoutProfile("warnings"));
        Assert.Equal(BridgeConnection.ToolTimeoutProfile.Interactive, BridgeConnectionArgs.SelectTimeoutProfile("messages"));
        Assert.Equal(BridgeConnection.ToolTimeoutProfile.Interactive, BridgeConnectionArgs.SelectTimeoutProfile("diagnostics-snapshot"));
    }

    [Fact]
    public void QuickDiagnosticsFallbackAllowsPlainPaging()
    {
        JsonObject args = new()
        {
            ["chunk_size"] = 100,
            ["chunk_index"] = 1,
        };

        Assert.True(ToolCatalog.CanUseQuickDiagnosticsFallback(args));
    }

    [Theory]
    [InlineData("severity", "warning")]
    [InlineData("code", "BP1044")]
    [InlineData("project", "VsIdeBridgeService")]
    [InlineData("path", "src")]
    [InlineData("file", "DiagnosticsTools.cs")]
    [InlineData("text", "suppression")]
    [InlineData("sort_by", "file")]
    [InlineData("group_by", "code")]
    [InlineData("group_sort_by", "count")]
    [InlineData("group_sort_direction", "desc")]
    public void QuickDiagnosticsFallbackRejectsServiceSideQueries(string parameter, string value)
    {
        JsonObject args = new() { [parameter] = value };

        Assert.False(ToolCatalog.CanUseQuickDiagnosticsFallback(args));
    }

    [Fact]
    public void QuickDiagnosticsFallbackRejectsGroupCountFilters()
    {
        JsonObject args = new() { ["group_min_count"] = 2 };

        Assert.False(ToolCatalog.CanUseQuickDiagnosticsFallback(args));
    }

    [Theory]
    [InlineData(true, false, false, true)]
    [InlineData(true, true, false, false)]
    [InlineData(false, false, false, false)]
    public void SolutionBuildCourtesyWaitOnlyAppliesToDefaultWaits(
        bool waitForCompletion,
        bool waitForCompletionExplicit,
        bool includeProject,
        bool expected)
    {
        Assert.Equal(expected, ToolCatalog.ShouldUseCourtesyWaitForBuild(
            waitForCompletion,
            waitForCompletionExplicit,
            includeProject,
            args: null));
    }

    [Fact]
    public void ProjectBuildDoesNotUseCourtesyWait()
    {
        JsonObject args = new() { ["project"] = "VsIdeBridgeService" };

        Assert.False(ToolCatalog.ShouldUseCourtesyWaitForBuild(
            waitForCompletion: true,
            waitForCompletionExplicit: false,
            includeProject: true,
            args));
    }

    [Fact]
    public void BuildToolsDescribeExplicitWaitAsTenMinutes()
    {
        JsonObject properties = GetSchemaProperties(GetDefinition("rebuild"));
        string description = properties["wait_for_completion"]!["description"]!.GetValue<string>();

        Assert.Contains("wait up to 10 minutes", description);
    }

    [Fact]
    public void DiagnosticsCompactionArgsRemoveHandleFileFiltersOnlyForSecondPass()
    {
        JsonObject args = new()
        {
            ["file"] = "w:35",
            ["path"] = "src/libslic3r",
            ["code"] = "C4244",
            ["chunk_size"] = 0,
        };

        JsonObject compactArgs = Assert.IsType<JsonObject>(ToolCatalog.CreateDiagnosticsCompactionArgs(args));

        Assert.False(compactArgs.ContainsKey("file"));
        Assert.Equal("src/libslic3r", compactArgs["path"]!.GetValue<string>());
        Assert.Equal("C4244", compactArgs["code"]!.GetValue<string>());
        Assert.Equal("w:35", args["file"]!.GetValue<string>());
    }

    private static ToolDefinition GetDefinition(string toolName)
    {
        Assert.True(ToolCatalog.CreateRegistry().TryGetDefinition(toolName, out ToolDefinition? definition));
        return definition;
    }

    private static JsonObject GetSchemaProperties(ToolDefinition definition)
    {
        JsonObject schema = definition.ParameterSchema;
        return Assert.IsType<JsonObject>(schema["properties"]);
    }
}
