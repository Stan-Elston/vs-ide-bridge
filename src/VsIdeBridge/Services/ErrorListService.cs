using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
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
            ["max"] = (JToken?)Max ?? JValue.CreateNull(),
        };
    }
}

internal sealed class ErrorListService(ReadinessService readinessService)
{
    private const int StableSampleCount = 3;
    private const int PopulationPollIntervalMilliseconds = 2000;
    private const int DefaultWaitTimeoutMilliseconds = 90_000;
    private const int MaxBestPracticeFiles = 64;
    private const int MaxBestPracticeFindingsPerFile = 25;
    private const int RepeatedStringThreshold = 3;
    private const int RepeatedNumberThreshold = 4;
    private const int MaxSuppressionFindingsPerFile = 5;

    private static readonly string[] BuildOutputPaneNames = ["Build", "Build Order"];
    private static readonly Regex ExplicitCodePattern = new(
        @"\b(?:LINK|LNK|MSB|VCR|E|C)\d+\b|\blnt-[a-z0-9-]+\b|\bInt-make\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex MsBuildDiagnosticPattern = new(
        @"^(?<file>[A-Za-z]:\\.*?|\S.*?)(?:\((?<line>\d+)(?:,(?<column>\d+))?\))?\s*:\s*(?<severity>warning|error)\s+(?<code>[A-Za-z]+[A-Za-z0-9-]*)\s*:\s*(?<message>.+?)(?:\s+\[(?<project>[^\]]+)\])?$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex StructuredOutputPattern = new(
        @"^(?<project>.+?)\s*>\s*(?<file>[A-Za-z]:\\.*?|\S.*?)(?:\((?<line>\d+)(?:,(?<column>\d+))?\))?\s*:\s*(?<severity>warning|error)\s+(?<code>[A-Za-z]+[A-Za-z0-9-]*)\s*:\s*(?<message>.+)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex StringLiteralPattern = new("\"([^\"\\r\\n]{4,})\"", RegexOptions.Compiled);
    private static readonly Regex NumberLiteralPattern = new(@"(?<![A-Za-z0-9_\.])(?<value>-?\d+(?:\.\d+)?)\b", RegexOptions.Compiled);
    private static readonly Regex SuspiciousRoundDownPattern = new(@"Math\s*\.\s*(?<op>Floor|Truncate)\s*\(", RegexOptions.Compiled);
    private static readonly Regex SuppressionIntentPattern = new(@"\b(?:fix|silence|suppress|workaround|appease)\b.{0,30}\b(?:error|warning|lint|analyzer)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private readonly ReadinessService _readinessService = readinessService;

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
                rows = [];
            }

            if (rows.Count == 0)
            {
                rows = ReadBuildOutputRows(context.Dte);
            }
        }
        else
        {
            rows = await WaitForRowsAsync(context, timeoutMilliseconds, intellisenseReady: waitForIntellisense).ConfigureAwait(true);
            if (rows.Count == 0)
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);
                rows = ReadBuildOutputRows(context.Dte);
            }
        }

        var bestPracticeRows = await Task.Run(() => AnalyzeBestPracticeFindings(rows), context.CancellationToken).ConfigureAwait(false);
        if (bestPracticeRows.Count > 0)
        {
            rows = [.. rows, .. bestPracticeRows];
        }

        var filteredRows = ApplyQuery(rows, query).ToArray();
        var severityCounts = CreateSeverityCounts();
        foreach (var row in filteredRows)
        {
            severityCounts[(string)row["severity"]!]++;
        }

        var totalSeverityCounts = CreateSeverityCounts();
        foreach (var row in rows)
        {
            totalSeverityCounts[(string)row["severity"]!]++;
        }

        return new JObject
        {
            ["count"] = filteredRows.Length,
            ["totalCount"] = rows.Count,
            ["severityCounts"] = JObject.FromObject(severityCounts),
            ["totalSeverityCounts"] = JObject.FromObject(totalSeverityCounts),
            ["hasErrors"] = severityCounts["Error"] > 0,
            ["hasWarnings"] = severityCounts["Warning"] > 0,
            ["filter"] = query?.ToJson() ?? [],
            ["rows"] = new JArray(filteredRows),
            ["groups"] = BuildGroups(filteredRows, query?.GroupBy),
        };
    }

    private static IReadOnlyList<JObject> AnalyzeBestPracticeFindings(IReadOnlyList<JObject> rows)
    {
        var findings = new List<JObject>();
        var files = rows
            .Select(row => row["file"]?.ToString())
            .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(MaxBestPracticeFiles)
            .ToArray();

        foreach (var file in files)
        {
            var content = SafeReadFile(file!);
            var perFileFindings = 0;
            if (string.IsNullOrWhiteSpace(content))
            {
                continue;
            }

            foreach (var finding in FindRepeatedStringLiterals(file!, content)
                .Concat(FindMagicNumbers(file!, content))
                .Concat(FindSuspiciousRoundDown(file!, content)))
            {
                findings.Add(finding);
                perFileFindings++;
                if (perFileFindings >= MaxBestPracticeFindingsPerFile)
                {
                    break;
                }
            }
        }

        return [.. findings
            .GroupBy(row => $"{(string?)row["code"]}|{(string?)row["file"]}|{(int?)row["line"]}", StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())];
    }

    private static string SafeReadFile(string filePath)
    {
        try
        {
            return File.ReadAllText(filePath);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static IEnumerable<JObject> FindRepeatedStringLiterals(string file, string content)
    {
        var occurrences = StringLiteralPattern.Matches(content)
            .Cast<Match>()
            .GroupBy(match => match.Groups[1].Value)
            .Where(group => group.Count() >= RepeatedStringThreshold)
            .OrderByDescending(group => group.Count())
            .ThenBy(group => group.Key, StringComparer.Ordinal)
            .Take(MaxBestPracticeFindingsPerFile);

        foreach (var repeated in occurrences)
        {
            yield return CreateBestPracticeRow(
                code: "BP1001",
                message: $"String literal '{repeated.Key}' is repeated {repeated.Count()} times. Extract a constant.",
                file: file,
                line: GetLineNumber(content, repeated.First().Index),
                symbol: repeated.Key);
        }
    }

    private static IEnumerable<JObject> FindMagicNumbers(string file, string content)
    {
        var matches = NumberLiteralPattern.Matches(content)
            .Cast<Match>()
            .Select(match =>
            {
                var value = match.Groups["value"].Value;
                return new { match, value };
            })
            .Where(item => item.value is not "0" and not "1" and not "-1")
            .GroupBy(item => item.value)
            .Where(group => group.Count() >= RepeatedNumberThreshold)
            .OrderByDescending(group => group.Count())
            .ThenBy(group => group.Key, StringComparer.Ordinal)
            .Take(MaxBestPracticeFindingsPerFile);

        foreach (var repeated in matches)
        {
            yield return CreateBestPracticeRow(
                code: "BP1002",
                message: $"Numeric literal '{repeated.Key}' appears {repeated.Count()} times. Replace magic numbers with named constants.",
                file: file,
                line: GetLineNumber(content, repeated.First().match.Index),
                symbol: repeated.Key);
        }
    }

    private static IEnumerable<JObject> FindSuspiciousRoundDown(string file, string content)
    {
        var lines = content.Split(["\r\n", "\n"], StringSplitOptions.None);
        var findingCount = 0;
        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            var roundDownMatch = SuspiciousRoundDownPattern.Match(line);
            if (roundDownMatch.Success && SuppressionIntentPattern.IsMatch(line))
            {
                yield return CreateBestPracticeRow(
                    code: "BP1003",
                    message: "Possible error suppression via rounding down detected. Investigate root cause instead of forcing Math.Floor/Math.Truncate.",
                    file: file,
                    line: i + 1,
                    symbol: $"Math.{roundDownMatch.Groups["op"].Value}");
                findingCount++;
                if (findingCount >= MaxSuppressionFindingsPerFile)
                {
                    yield break;
                }
            }
        }
    }

    private static JObject CreateBestPracticeRow(string code, string message, string file, int line, string symbol)
    {
        return new JObject
        {
            ["severity"] = "Error",
            ["code"] = code,
            ["codeFamily"] = "best-practice",
            ["tool"] = "best-practice",
            ["message"] = message,
            ["project"] = string.Empty,
            ["file"] = file,
            ["line"] = line,
            ["column"] = 1,
            ["symbols"] = new JArray(symbol),
            ["source"] = "best-practice",
        };
    }

    private static int GetLineNumber(string content, int index)
    {
        var line = 1;
        for (var i = 0; i < index && i < content.Length; i++)
        {
            if (content[i] == '\n')
            {
                line++;
            }
        }

        return line;
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

    private async Task<IReadOnlyList<JObject>> WaitForRowsAsync(IdeCommandContext context, int timeoutMilliseconds, bool intellisenseReady = false)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);
        EnsureErrorListWindow(context.Dte);

        var timeout = timeoutMilliseconds > 0 ? timeoutMilliseconds : DefaultWaitTimeoutMilliseconds;
        var deadline = DateTimeOffset.UtcNow.AddMilliseconds(timeout);
        var lastRows = Array.Empty<JObject>();
        int? lastCount = null;
        var stableSamples = 0;
        // When IntelliSense has already confirmed ready, one stable read is sufficient.
        var requiredStableSamples = intellisenseReady ? 1 : StableSampleCount;

        while (DateTimeOffset.UtcNow < deadline)
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

                lastRows = [.. rows];
                // A clean solution should return promptly once the Error List is stable,
                // instead of waiting out the full timeout for a non-zero row count.
                if (stableSamples >= requiredStableSamples)
                {
                    return rows;
                }
            }

            await Task.Delay(PopulationPollIntervalMilliseconds, context.CancellationToken).ConfigureAwait(false);
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);
        }

        return lastRows;
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
            return [];
        }

        var text = TryReadBuildOutputText(dte, pane);
        if (string.IsNullOrWhiteSpace(text))
        {
            return [];
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
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        try
        {
            return PathNormalization.NormalizeFilePath(value);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            // Build output can include non-path tokens in the file column; keep the raw value
            // so diagnostics still flow without failing the entire error-list request.
            return value;
        }
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

        return [.. symbols.Take(8)];
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

        if (query.Max > 0)
        {
            filtered = filtered.Take(query.Max.Value);
        }

        return [.. filtered];
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
            return [];
        }

        var groupKey = groupBy!;
        Func<JObject, string> keySelector = groupKey.ToLowerInvariant() switch
        {
            "code" => row => (string?)row["code"] ?? string.Empty,
            "file" => row => (string?)row["file"] ?? string.Empty,
            "project" => row => (string?)row["project"] ?? string.Empty,
            "tool" => row => (string?)row["tool"] ?? string.Empty,
            _ => static _ => string.Empty,
        };

        if (groupKey is not ("code" or "file" or "project" or "tool"))
        {
            return [];
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
