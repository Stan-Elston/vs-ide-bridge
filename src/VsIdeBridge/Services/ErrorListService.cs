using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed class ErrorListQuery
{
    public string? Severity { get; set; }

    public string? Code { get; set; }

    public string? Project { get; set; }

    public string? Path { get; set; }

    public string? Text { get; set; }

    public string? GroupBy { get; set; }

    public int? Max { get; set; }

    public JObject ToJson()
    {
        return new JObject
        {
            ["severity"] = Severity ?? string.Empty,
            ["code"] = Code ?? string.Empty,
            ["project"] = Project ?? string.Empty,
            ["path"] = Path ?? string.Empty,
            ["text"] = Text ?? string.Empty,
            ["groupBy"] = GroupBy ?? string.Empty,
            ["max"] = Max.HasValue ? Max.Value : JValue.CreateNull(),
        };
    }
}

internal sealed class ErrorListService
{
    private const int StableSampleCount = 3;
    private const int PopulationPollIntervalMilliseconds = 2000;

    private static readonly string[] BuildOutputPaneNames = { "Build", "Build Order" };
    private static readonly Regex ExplicitCodePattern = new(
        @"\b(?:LINK|LNK|MSB|VCR|E|C)\d+\b|\blnt-[a-z0-9-]+\b|\bInt-make\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex MsBuildDiagnosticPattern = new(
        @"^(?<file>[A-Za-z]:\\.*?|\S.*?)(?:\((?<line>\d+)(?:,(?<column>\d+))?\))?\s*:\s*(?<severity>warning|error)\s+(?<code>[A-Za-z]+[A-Za-z0-9-]*)\s*:\s*(?<message>.+?)(?:\s+\[(?<project>[^\]]+)\])?$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex StructuredOutputPattern = new(
        @"^(?<project>.+?)\s*>\s*(?<file>[A-Za-z]:\\.*?|\S.*?)(?:\((?<line>\d+)(?:,(?<column>\d+))?\))?\s*:\s*(?<severity>warning|error)\s+(?<code>[A-Za-z]+[A-Za-z0-9-]*)\s*:\s*(?<message>.+)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private readonly ReadinessService _readinessService;

    public ErrorListService(ReadinessService readinessService)
    {
        _readinessService = readinessService;
    }

    public async Task<JObject> GetErrorListAsync(
        IdeCommandContext context,
        bool waitForIntellisense,
        int timeoutMilliseconds,
        bool quickSnapshot = false,
        ErrorListQuery? query = null)
    {
        if (waitForIntellisense)
        {
            await _readinessService.WaitForReadyAsync(context, timeoutMilliseconds).ConfigureAwait(true);
        }

        IReadOnlyList<JObject> rows;
        if (quickSnapshot)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);
            EnsureErrorListWindow(context.Dte);
            try
            {
                rows = ReadRows(context.Dte);
            }
            catch (InvalidOperationException)
            {
                rows = Array.Empty<JObject>();
            }

            if (rows.Count == 0)
            {
                rows = ReadBuildOutputRows(context.Dte);
            }
        }
        else
        {
            rows = await WaitForRowsAsync(context, timeoutMilliseconds).ConfigureAwait(true);
            if (rows.Count == 0)
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);
                rows = ReadBuildOutputRows(context.Dte);
            }
        }
        var filteredRows = ApplyQuery(rows, query).ToArray();
        var severityCounts = CreateSeverityCounts();
        foreach (var row in filteredRows)
        {
            severityCounts[(string)row["severity"]!] += 1;
        }

        var totalSeverityCounts = CreateSeverityCounts();
        foreach (var row in rows)
        {
            totalSeverityCounts[(string)row["severity"]!] += 1;
        }

        return new JObject
        {
            ["count"] = filteredRows.Length,
            ["totalCount"] = rows.Count,
            ["severityCounts"] = JObject.FromObject(severityCounts),
            ["totalSeverityCounts"] = JObject.FromObject(totalSeverityCounts),
            ["hasErrors"] = severityCounts["Error"] > 0,
            ["hasWarnings"] = severityCounts["Warning"] > 0,
            ["filter"] = query?.ToJson() ?? new JObject(),
            ["rows"] = new JArray(filteredRows),
            ["groups"] = BuildGroups(filteredRows, query?.GroupBy),
        };
    }

    private static Dictionary<string, int> CreateSeverityCounts()
    {
        return new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["Error"] = 0,
            ["Warning"] = 0,
            ["Message"] = 0,
        };
    }

    private async Task<IReadOnlyList<JObject>> WaitForRowsAsync(IdeCommandContext context, int timeoutMilliseconds)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);
        EnsureErrorListWindow(context.Dte);

        var timeout = timeoutMilliseconds > 0 ? timeoutMilliseconds : 90000;
        var deadline = DateTimeOffset.UtcNow.AddMilliseconds(timeout);
        var lastRows = Array.Empty<JObject>();
        int? lastCount = null;
        var stableSamples = 0;

        while (true)
        {
            context.CancellationToken.ThrowIfCancellationRequested();

            IReadOnlyList<JObject>? rows = null;
            try
            {
                rows = ReadRows(context.Dte);
            }
            catch (InvalidOperationException)
            {
                EnsureErrorListWindow(context.Dte);
            }

            if (rows is not null)
            {
                if (rows.Count != lastCount)
                {
                    lastCount = rows.Count;
                    stableSamples = 1;
                }
                else
                {
                    stableSamples++;
                }

                lastRows = rows.ToArray();
                // A clean solution should return promptly once the Error List is stable,
                // instead of waiting out the full timeout for a non-zero row count.
                if (stableSamples >= StableSampleCount)
                {
                    return rows;
                }
            }

            if (DateTimeOffset.UtcNow >= deadline)
            {
                return lastRows;
            }

            await Task.Delay(PopulationPollIntervalMilliseconds, context.CancellationToken).ConfigureAwait(false);
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);
        }
    }

    private static IReadOnlyList<JObject> ReadRows(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var window = TryGetErrorListWindow(dte);
        if (window?.Object is not ErrorList errorList)
        {
            throw new InvalidOperationException("Error List window is not available.");
        }

        var items = errorList.ErrorItems;
        var rows = new List<JObject>(items.Count);
        for (var i = 1; i <= items.Count; i++)
        {
            var item = items.Item(i);
            var severity = MapSeverity(item.ErrorLevel);
            var description = item.Description ?? string.Empty;
            var code = InferCode(description, item.Project ?? string.Empty, item.FileName ?? string.Empty, item.Line);
            rows.Add(new JObject
            {
                ["severity"] = severity,
                ["code"] = code,
                ["codeFamily"] = InferCodeFamily(code),
                ["tool"] = InferTool(code, description),
                ["message"] = description,
                ["project"] = item.Project ?? string.Empty,
                ["file"] = item.FileName ?? string.Empty,
                ["line"] = item.Line,
                ["column"] = item.Column,
                ["symbols"] = new JArray(ExtractSymbols(description)),
            });
        }

        return rows;
    }

    private static IReadOnlyList<JObject> ReadBuildOutputRows(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var pane = TryGetBuildOutputPane(dte);
        if (pane is null)
        {
            return Array.Empty<JObject>();
        }

        var text = TryReadBuildOutputText(dte, pane);
        if (string.IsNullOrWhiteSpace(text))
        {
            return Array.Empty<JObject>();
        }

        var rows = new List<JObject>();
        using var reader = new StringReader(text);
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            if (TryParseBuildOutputLine(line, out var row))
            {
                rows.Add(row);
            }
        }

        return rows;
    }

    private static string TryReadBuildOutputText(DTE2 dte, OutputWindowPane pane)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        ActivateBuildOutputPane(dte, pane);

        for (var attempt = 0; attempt < 2; attempt++)
        {
            try
            {
                return pane.TextDocument is TextDocument textDocument
                    ? ReadTextDocument(textDocument)
                    : string.Empty;
            }
            catch (COMException)
            {
                if (attempt == 0)
                {
                    ActivateBuildOutputPane(dte, pane);
                    continue;
                }

                return string.Empty;
            }
        }

        return string.Empty;
    }

    private static void EnsureErrorListWindow(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (TryGetErrorListWindow(dte)?.Object is ErrorList)
        {
            return;
        }

        try
        {
            dte.ExecuteCommand("View.ErrorList", string.Empty);
        }
        catch
        {
        }
    }

    private static void ActivateBuildOutputPane(DTE2 dte, OutputWindowPane pane)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            dte.ExecuteCommand("View.Output", string.Empty);
        }
        catch
        {
        }

        try
        {
            pane.Activate();
        }
        catch
        {
        }
    }

    private static OutputWindowPane? TryGetBuildOutputPane(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        foreach (OutputWindowPane pane in dte.ToolWindows.OutputWindow.OutputWindowPanes)
        {
            var paneName = pane.Name;
            foreach (var candidateName in BuildOutputPaneNames)
            {
                if (string.Equals(paneName, candidateName, StringComparison.OrdinalIgnoreCase))
                {
                    return pane;
                }
            }
        }

        return null;
    }

    private static string ReadTextDocument(TextDocument textDocument)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var start = textDocument.StartPoint.CreateEditPoint();
        return start.GetText(textDocument.EndPoint);
    }

    private static bool TryParseBuildOutputLine(string line, out JObject row)
    {
        row = null!;
        if (string.IsNullOrWhiteSpace(line))
        {
            return false;
        }

        var match = StructuredOutputPattern.Match(line);
        if (!match.Success)
        {
            match = MsBuildDiagnosticPattern.Match(line);
        }

        if (!match.Success)
        {
            return false;
        }

        var severity = NormalizeParsedSeverity(match.Groups["severity"].Value);
        var description = match.Groups["message"].Value.Trim();
        var project = match.Groups["project"].Value.Trim();
        var file = NormalizeFilePath(match.Groups["file"].Value.Trim());
        var lineNumber = ParseOptionalInt(match.Groups["line"].Value);
        var columnNumber = ParseOptionalInt(match.Groups["column"].Value);
        var code = NormalizeCode(match.Groups["code"].Value, description, project, file, lineNumber);

        row = new JObject
        {
            ["severity"] = severity,
            ["code"] = code,
            ["codeFamily"] = InferCodeFamily(code),
            ["tool"] = InferTool(code, description),
            ["message"] = description,
            ["project"] = project,
            ["file"] = file,
            ["line"] = lineNumber,
            ["column"] = columnNumber,
            ["symbols"] = new JArray(ExtractSymbols(description)),
            ["source"] = "build-output",
        };
        return true;
    }

    private static string NormalizeParsedSeverity(string severity)
    {
        return severity.Equals("error", StringComparison.OrdinalIgnoreCase) ? "Error" : "Warning";
    }

    private static string NormalizeFilePath(string value)
    {
        return string.IsNullOrWhiteSpace(value) ? string.Empty : PathNormalization.NormalizeFilePath(value);
    }

    private static int ParseOptionalInt(string value)
    {
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ? parsed : 0;
    }

    private static string NormalizeCode(string explicitCode, string description, string project, string fileName, int line)
    {
        return !string.IsNullOrWhiteSpace(explicitCode)
            ? explicitCode
            : InferCode(description, project, fileName, line);
    }

    private static Window? TryGetErrorListWindow(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        return dte.Windows
            .Cast<Window>()
            .FirstOrDefault(IsErrorListWindow);
    }

    private static bool IsErrorListWindow(Window candidate)
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        return string.Equals(candidate.Caption, "Error List", StringComparison.OrdinalIgnoreCase);
    }

    private static string MapSeverity(vsBuildErrorLevel level)
    {
        return level switch
        {
            vsBuildErrorLevel.vsBuildErrorLevelHigh => "Error",
            vsBuildErrorLevel.vsBuildErrorLevelMedium => "Warning",
            _ => "Message",
        };
    }

    private static string InferCode(string description, string project, string fileName, int line)
    {
        var explicitCode = ExtractExplicitCode(description);
        if (!string.IsNullOrWhiteSpace(explicitCode))
        {
            return explicitCode;
        }

        if (description.IndexOf("identifier \"", StringComparison.OrdinalIgnoreCase) >= 0 &&
            description.IndexOf("\" is undefined", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return "E0020";
        }

        if (description.IndexOf("can be made static", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return "VCR003";
        }

        if (description.IndexOf("can be made const", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return "VCR001";
        }

        if (description.IndexOf("Return value ignored", StringComparison.OrdinalIgnoreCase) >= 0 &&
            description.IndexOf("UnregisterWaitEx", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return "C6031";
        }

        if (description.IndexOf("PCH warning:", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return "Int-make";
        }

        if (description.IndexOf("doesn't deduce references", StringComparison.OrdinalIgnoreCase) >= 0 &&
            description.IndexOf("possibly unintended copy", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return "lnt-accidental-copy";
        }

        if (description.IndexOf("cannot open file '", StringComparison.OrdinalIgnoreCase) >= 0 &&
            IsLinkerContext(project, fileName, line))
        {
            return "LNK1104";
        }

        return string.Empty;
    }

    private static string ExtractExplicitCode(string description)
    {
        var match = ExplicitCodePattern.Match(description);
        return match.Success ? NormalizeCode(match.Value) : string.Empty;
    }

    private static string NormalizeCode(string code)
    {
        if (code.StartsWith("LINK", StringComparison.OrdinalIgnoreCase) &&
            code.Length > 4 &&
            int.TryParse(code.Substring(4), NumberStyles.None, CultureInfo.InvariantCulture, out _))
        {
            return "LNK" + code.Substring(4);
        }

        if (code.StartsWith("lnt-", StringComparison.OrdinalIgnoreCase))
        {
            return code.ToLowerInvariant();
        }

        return code.ToUpperInvariant();
    }

    private static string InferCodeFamily(string code)
    {
        if (string.IsNullOrWhiteSpace(code))
        {
            return string.Empty;
        }

        if (code.StartsWith("LNK", StringComparison.OrdinalIgnoreCase))
        {
            return "linker";
        }

        if (code.StartsWith("MSB", StringComparison.OrdinalIgnoreCase))
        {
            return "msbuild";
        }

        if (code.StartsWith("VCR", StringComparison.OrdinalIgnoreCase))
        {
            return "analyzer";
        }

        if (code.StartsWith("lnt-", StringComparison.OrdinalIgnoreCase))
        {
            return "linter";
        }

        if (code.StartsWith("C", StringComparison.OrdinalIgnoreCase) ||
            code.StartsWith("E", StringComparison.OrdinalIgnoreCase))
        {
            return "compiler";
        }

        return "other";
    }

    private static string InferTool(string code, string description)
    {
        var family = InferCodeFamily(code);
        if (!string.IsNullOrWhiteSpace(family))
        {
            return family;
        }

        if (description.IndexOf("IntelliSense", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return "intellisense";
        }

        return "diagnostic";
    }

    private static IReadOnlyList<string> ExtractSymbols(string description)
    {
        var symbols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (Match match in Regex.Matches(description, "\"([^\"]+)\"|'([^']+)'"))
        {
            var value = match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value;
            if (LooksLikeSymbol(value))
            {
                symbols.Add(value);
            }
        }

        foreach (Match match in Regex.Matches(description, @"\b[A-Za-z_~][A-Za-z0-9_:<>~]*\b"))
        {
            var value = match.Value;
            if (LooksLikeSymbol(value))
            {
                symbols.Add(value);
            }
        }

        return symbols.Take(8).ToArray();
    }

    private static bool LooksLikeSymbol(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length < 2)
        {
            return false;
        }

        return value.Contains("::", StringComparison.Ordinal) ||
            value.Contains("_", StringComparison.Ordinal) ||
            char.IsUpper(value[0]) ||
            value.StartsWith("C", StringComparison.OrdinalIgnoreCase) ||
            value.StartsWith("E", StringComparison.OrdinalIgnoreCase);
    }

    private static IReadOnlyList<JObject> ApplyQuery(IReadOnlyList<JObject> rows, ErrorListQuery? query)
    {
        if (query is null)
        {
            return rows;
        }

        IEnumerable<JObject> filtered = rows;

        if (!string.IsNullOrWhiteSpace(query.Severity) &&
            !string.Equals(query.Severity, "all", StringComparison.OrdinalIgnoreCase))
        {
            filtered = filtered.Where(row => string.Equals(
                (string?)row["severity"],
                NormalizeSeverity(query.Severity),
                StringComparison.OrdinalIgnoreCase));
        }

        if (!string.IsNullOrWhiteSpace(query.Code))
        {
            filtered = filtered.Where(row => ((string?)row["code"] ?? string.Empty)
                .StartsWith(query.Code, StringComparison.OrdinalIgnoreCase));
        }

        if (!string.IsNullOrWhiteSpace(query.Project))
        {
            filtered = filtered.Where(row => ((string?)row["project"] ?? string.Empty)
                .IndexOf(query.Project, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        if (!string.IsNullOrWhiteSpace(query.Path))
        {
            filtered = filtered.Where(row => ((string?)row["file"] ?? string.Empty)
                .IndexOf(query.Path, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        if (!string.IsNullOrWhiteSpace(query.Text))
        {
            filtered = filtered.Where(row => ((string?)row["message"] ?? string.Empty)
                .IndexOf(query.Text, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        if (query.Max.HasValue && query.Max.Value > 0)
        {
            filtered = filtered.Take(query.Max.Value);
        }

        return filtered.ToArray();
    }

    private static string NormalizeSeverity(string? severity)
    {
        return severity?.ToLowerInvariant() switch
        {
            "error" => "Error",
            "warning" => "Warning",
            "message" => "Message",
            _ => severity ?? string.Empty,
        };
    }

    private static JArray BuildGroups(IReadOnlyList<JObject> rows, string? groupBy)
    {
        if (string.IsNullOrWhiteSpace(groupBy))
        {
            return new JArray();
        }

        var groupKey = groupBy!;
        Func<JObject, string> keySelector = groupKey.ToLowerInvariant() switch
        {
            "code" => row => (string?)row["code"] ?? string.Empty,
            "file" => row => (string?)row["file"] ?? string.Empty,
            "project" => row => (string?)row["project"] ?? string.Empty,
            "tool" => row => (string?)row["tool"] ?? string.Empty,
            _ => row => string.Empty,
        };

        if (groupKey is not ("code" or "file" or "project" or "tool"))
        {
            return new JArray();
        }

        return new JArray(rows
            .GroupBy(keySelector, StringComparer.OrdinalIgnoreCase)
            .Where(group => !string.IsNullOrWhiteSpace(group.Key))
            .OrderByDescending(group => group.Count())
            .ThenBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
            .Select(group => new JObject
            {
                ["key"] = group.Key,
                ["groupBy"] = groupKey,
                ["count"] = group.Count(),
                ["sampleMessage"] = (string?)group.First()["message"] ?? string.Empty,
                ["sampleFile"] = (string?)group.First()["file"] ?? string.Empty,
                ["sampleCode"] = (string?)group.First()["code"] ?? string.Empty,
            }));
    }

    private static bool IsLinkerContext(string project, string fileName, int line)
    {
        var normalizedFile = (fileName ?? string.Empty).Replace('/', '\\');
        if (normalizedFile.EndsWith("\\LINK", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return string.IsNullOrWhiteSpace(project) && string.IsNullOrWhiteSpace(fileName) && line <= 0;
    }
}
