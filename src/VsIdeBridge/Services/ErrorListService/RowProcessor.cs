using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using static VsIdeBridge.Diagnostics.ErrorListConstants;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed partial class ErrorListService
{
    private static string InferCodeFamily(string code)
    {
        if (string.IsNullOrWhiteSpace(code))
        {
            return string.Empty;
        }

        if (code.StartsWith("BP", StringComparison.OrdinalIgnoreCase))
        {
            return "best-practice";
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
        string family = InferCodeFamily(code);
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

    private static JObject ApplyDiagnosticAnnotations(JObject row)
    {
        string code = GetRowString(row, CodeKey);
        if (!IsVisualStudioAnalyzerLimitedCode(code))
        {
            return row;
        }

        row["analyzerLimited"] = true;
        row[CodeFamilyKey] = "analyzer";
        row[SourceKey] = string.IsNullOrWhiteSpace(GetRowString(row, SourceKey))
            ? "visual-studio-analyzer"
            : GetRowString(row, SourceKey);
        row[ToolKey] = string.IsNullOrWhiteSpace(GetRowString(row, ToolKey))
            ? "visual-studio-analyzer"
            : GetRowString(row, ToolKey);
        row[AuthorityKey] = "visual-studio-analyzer-limited";
        row[GuidanceKey] = MergeDiagnosticText(
            GetRowString(row, GuidanceKey),
            "Visual Studio analyzer-limited row: VCR001/VCR003 can lag or misread modern C++ constructs such as '= delete' and explicit specializations. Verify against current source and build output before editing.");
        row[SuggestedActionKey] = MergeDiagnosticText(
            GetRowString(row, SuggestedActionKey),
            "Treat as advisory analyzer output; refresh diagnostics or build before making semantic changes.");
        row[LlmFixPromptKey] = MergeDiagnosticText(
            GetRowString(row, LlmFixPromptKey),
            "Do not assume this VCR row is authoritative. Confirm the construct in source/build output first, especially for '= delete' and explicit specialization patterns.");
        return row;
    }

    private static bool IsVisualStudioAnalyzerLimitedCode(string code)
        => string.Equals(code, "VCR001", StringComparison.OrdinalIgnoreCase)
            || string.Equals(code, "VCR003", StringComparison.OrdinalIgnoreCase);

    private static string MergeDiagnosticText(string existing, string addition)
    {
        if (string.IsNullOrWhiteSpace(existing))
        {
            return addition;
        }

        return existing.Contains(addition, StringComparison.Ordinal)
            ? existing
            : string.Concat(existing, " ", addition);
    }

    private static IReadOnlyList<string> ExtractSymbols(string description)
    {
        HashSet<string> symbols = [];

        foreach (Match match in Regex.Matches(description, "\"(?<doubleQuoted>[^\"]+)\"|'(?<singleQuoted>[^']+)'"))
        {
            string value = match.Groups["doubleQuoted"].Success
                ? match.Groups["doubleQuoted"].Value
                : match.Groups["singleQuoted"].Value;
            if (LooksLikeSymbol(value))
            {
                symbols.Add(value);
            }
        }

        foreach (Match match in Regex.Matches(description, @"\b[A-Za-z_~][A-Za-z0-9_:<>~]*\b"))
        {
            string value = match.Value;
            if (LooksLikeSymbol(value))
            {
                symbols.Add(value);
            }
        }

        return [.. symbols.Take(MaxSymbolsPerDiagnostic)];
    }

    private static bool LooksLikeSymbol(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length < MinimumSymbolLength)
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
                GetRowString(row, SeverityKey),
                NormalizeSeverity(query.Severity),
                StringComparison.OrdinalIgnoreCase));
        }

        if (!string.IsNullOrWhiteSpace(query.Code))
        {
            filtered = filtered.Where(row => GetRowString(row, CodeKey)
                .StartsWith(query.Code, StringComparison.OrdinalIgnoreCase));
        }

        if (!string.IsNullOrWhiteSpace(query.Project))
        {
            filtered = filtered.Where(row => GetRowString(row, ProjectKey)
                .IndexOf(query.Project, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        if (!string.IsNullOrWhiteSpace(query.Path))
        {
            filtered = filtered.Where(row => NormalizePath(GetRowString(row, FileKey))
                .IndexOf(NormalizePath(query.Path!), StringComparison.OrdinalIgnoreCase) >= 0);
        }

        string fileFilter = (query.File ?? string.Empty).Trim();
        if (fileFilter.Length > 0)
        {
            filtered = filtered.Where(row => MatchesFileFilter(row, fileFilter));
        }

        if (!string.IsNullOrWhiteSpace(query.Text))
        {
            filtered = filtered.Where(row => GetRowString(row, MessageKey)
                .IndexOf(query.Text, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        return [.. filtered];
    }

    /// <summary>
    /// Sorts <paramref name="rows"/> by the field named in <paramref name="sortBy"/>.
    /// Returns the original sequence unchanged when <paramref name="sortBy"/> is null/empty/none.
    /// </summary>
    private static IEnumerable<JObject> ApplySort(IEnumerable<JObject> rows, string? sortBy, string? sortDirection)
    {
        if (string.IsNullOrWhiteSpace(sortBy) ||
            string.Equals(sortBy, "none", StringComparison.OrdinalIgnoreCase))
        {
            return rows;
        }

        bool descending = string.Equals(sortDirection, "desc",       StringComparison.OrdinalIgnoreCase)
                       || string.Equals(sortDirection, "descending", StringComparison.OrdinalIgnoreCase);

        string field = sortBy!.Trim().ToLowerInvariant();

        // Line sorts numerically; everything else sorts as a case-insensitive string.
        if (field == "line")
        {
            return descending
                ? rows.OrderByDescending(row => GetNullableRowInt(row, LineKey) ?? 0)
                : rows.OrderBy(row => GetNullableRowInt(row, LineKey) ?? 0);
        }

        string key = field switch
        {
            "severity"       => SeverityKey,
            "file" or "path" => FileKey,
            "code"           => CodeKey,
            "message"        => MessageKey,
            "project"        => ProjectKey,
            _                => SeverityKey,   // graceful fallback
        };

        return descending
            ? rows.OrderByDescending(row => GetRowString(row, key), StringComparer.OrdinalIgnoreCase)
            : rows.OrderBy(row => GetRowString(row, key), StringComparer.OrdinalIgnoreCase);
    }

    private static bool MatchesFileFilter(JObject row, string file)
    {
        string value = GetRowString(row, FileKey);
        return NormalizePath(value).IndexOf(NormalizePath(file), StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static string NormalizePath(string path) => path.Replace('\\', '/');

    private delegate bool TableValueReader(string keyName, out object content);

    private static string CreateFindingIdentity(JObject row)
    {
        return string.Join(
            "|",
            GetRowString(row, CodeKey),
            GetRowString(row, FileKey),
            (GetNullableRowInt(row, LineKey) ?? 0).ToString(CultureInfo.InvariantCulture));
    }

    private static string CreateDiagnosticIdentity(JObject row)
    {
        string file = GetRowString(row, FileKey);
        int line = GetNullableRowInt(row, LineKey) ?? 0;
        int column = GetNullableRowInt(row, ColumnKey) ?? 0;
        NormalizeBuildOutputLocation(ref file, ref line, ref column);

        return string.Join(
            "|",
            GetRowString(row, SeverityKey),
            GetRowString(row, CodeKey),
            file,
            line.ToString(CultureInfo.InvariantCulture),
            GetRowString(row, MessageKey));
    }

    private static JObject SelectPreferredDiagnosticRow(IGrouping<string, JObject> group)
    {
        return group
            .OrderByDescending(GetDiagnosticLocationQuality)
            .First();
    }

    private static int GetDiagnosticLocationQuality(JObject row)
    {
        string file = GetRowString(row, FileKey);
        int line = GetNullableRowInt(row, LineKey) ?? 0;
        int column = GetNullableRowInt(row, ColumnKey) ?? 0;
        NormalizeBuildOutputLocation(ref file, ref line, ref column);

        int score = 0;
        if (!string.IsNullOrWhiteSpace(file))
        {
            score += 1;
        }

        if (line > 0)
        {
            score += DiagnosticLineQualityScore;
        }

        if (column > 0)
        {
            score += DiagnosticColumnQualityScore;
        }

        if (!string.IsNullOrWhiteSpace(GetRowString(row, HelpUriKey)))
        {
            score += 1;
        }

        if (!string.Equals(GetRowString(row, SourceKey), "build-output", StringComparison.OrdinalIgnoreCase))
        {
            score += 1;
        }

        if (file.IndexOf("\\src\\", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            score += 1;
        }

        if (file.IndexOf("\\build\\", StringComparison.OrdinalIgnoreCase) >= 0 ||
            file.IndexOf("\\bin\\", StringComparison.OrdinalIgnoreCase) >= 0 ||
            file.IndexOf("\\obj\\", StringComparison.OrdinalIgnoreCase) >= 0 ||
            file.IndexOf("\\out\\", StringComparison.OrdinalIgnoreCase) >= 0 ||
            file.IndexOf("\\output\\", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            score -= 1;
        }

        return score;
    }

    private static string GetRowString(JObject row, string key)
    {
        return (string?)row[key] ?? string.Empty;
    }

    private static int? GetNullableRowInt(JObject row, string key)
    {
        return (int?)row[key];
    }

    private static void LogNonCriticalException(Exception ex)
    {
        BridgeActivityLog.LogWarning("ErrorListService", ex.GetType().Name, ex);
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

        string groupKey = groupBy!;
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

        IEnumerable<JObject> groups = rows
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
            });

        return [.. groups];
    }

    private static bool IsLinkerContext(string project, string fileName, string line)
    {
        string normalizedFile = (fileName ?? string.Empty).Replace('/', '\\');
        if (normalizedFile.EndsWith("\\LINK", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return string.IsNullOrWhiteSpace(project) && string.IsNullOrWhiteSpace(fileName) && int.TryParse(line, out int lineNumber) && lineNumber <= 0;
    }

    private static Window? TryGetErrorListWindow(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        try
        {
            return dte.Windows.Item("{DD1DDD20-D0F8-11ce-8C69-00AA004AC40}");
        }
        catch
        {
            return null;
        }
    }

    private static string MapSeverity(vsBuildErrorLevel errorLevel)
    {
        return errorLevel switch
        {
            vsBuildErrorLevel.vsBuildErrorLevelHigh => "Error",
            vsBuildErrorLevel.vsBuildErrorLevelMedium => "Warning",
            vsBuildErrorLevel.vsBuildErrorLevelLow => "Message",
            _ => "Error",
        };
    }

    private static string InferCode(string description)
    {
        string explicitCode = ExtractExplicitCode(description);
        if (!string.IsNullOrWhiteSpace(explicitCode))
        {
            return explicitCode;
        }
        // Additional inference logic could go here
        return string.Empty;
    }
}
