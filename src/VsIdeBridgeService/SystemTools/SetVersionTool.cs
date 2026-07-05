using System.IO;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace VsIdeBridgeService.SystemTools;

internal static partial class SetVersionTool
{
    public static Task<JsonNode> ExecuteAsync(JsonNode? id, JsonObject? args, BridgeConnection bridge)
    {
        string version = args?["version"]?.GetValue<string>() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(version))
        {
            throw new McpRequestException(id, McpErrorCodes.InvalidParams, "Missing required argument 'version'.");
        }

        string solutionDirectory = ServiceToolPaths.ResolveSolutionDirectory(bridge);
        JsonArray updatedFiles = [];
        JsonArray skippedFiles = [];

        UpdateFile(
            Path.Combine(solutionDirectory, "Directory.Build.props"),
            text => VersionTagRegex().Replace(text, $"<Version>{version}</Version>"),
            text => VersionTagRegex().IsMatch(text),
            "Directory.Build.props",
            updatedFiles, skippedFiles);

        UpdateFile(
            Path.Combine(solutionDirectory, "src", "VsIdeBridge", "source.extension.vsixmanifest"),
            text => VsixVersionRegex().Replace(text, version),
            text => VsixVersionRegex().IsMatch(text),
            "src/VsIdeBridge/source.extension.vsixmanifest",
            updatedFiles, skippedFiles);

        UpdateFile(
            Path.Combine(solutionDirectory, "installer", "inno", "vs-ide-bridge.iss"),
            text => IssVersionRegex().Replace(text, version),
            text => IssVersionRegex().IsMatch(text),
            "installer/inno/vs-ide-bridge.iss",
            updatedFiles, skippedFiles);

        bool hasSkipped = skippedFiles.Count > 0;
        JsonObject payload = new()
        {
            ["success"] = true,
            ["version"] = version,
            ["solutionDirectory"] = solutionDirectory,
            ["updated_files"] = updatedFiles,
            ["file_count"] = updatedFiles.Count,
        };
        if (hasSkipped)
        {
            payload["skipped_files"] = skippedFiles;
            payload["warning"] = $"Version pattern not found in {skippedFiles.Count} file(s) - those files were NOT updated.";
        }

        string successText = hasSkipped
            ? $"Updated {updatedFiles.Count} file(s) to {version}. WARNING: version pattern not found in {skippedFiles.Count} file(s) - check skipped_files."
            : $"Updated {updatedFiles.Count} version file(s) to {version}.";
        return Task.FromResult(ToolResultFormatter.StructuredToolResult(payload, args, successText: successText));
    }

    private static void UpdateFile(
        string filePath,
        Func<string, string> transform,
        Func<string, bool> hasMatch,
        string resultPath,
        JsonArray updatedFiles,
        JsonArray skippedFiles)
    {
        if (!File.Exists(filePath))
        {
            return;
        }

        string text = File.ReadAllText(filePath);
        if (!hasMatch(text))
        {
            skippedFiles.Add(JsonValue.Create(resultPath));
            return;
        }

        string next = transform(text);
        if (!string.Equals(next, text, StringComparison.Ordinal))
        {
            File.WriteAllText(filePath, next);
        }

        updatedFiles.Add(JsonValue.Create(resultPath));
    }

    [GeneratedRegex("<Version>[^<]*</Version>")]
    private static partial Regex VersionTagRegex();

    [GeneratedRegex("""(?<=<Identity[^>]+Version=")[^"]*(?=")""")]
    private static partial Regex VsixVersionRegex();

    [GeneratedRegex("""(?<=#define MyAppVersion ")[^"]*(?=")""")]
    private static partial Regex IssVersionRegex();
}
