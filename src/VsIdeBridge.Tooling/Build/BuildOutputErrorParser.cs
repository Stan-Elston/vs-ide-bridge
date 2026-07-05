using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace VsIdeBridge.Tooling.Build;

/// <summary>
/// One compiler/linker/MSBuild diagnostic extracted from raw Build output text.
/// <see cref="File"/> holds the source path for located diagnostics, or the object/library/target
/// origin for linker-style diagnostics that carry no source location. When MSVC attributes a
/// warning to an external header, <see cref="OriginFile"/> carries the project-code instantiation
/// site parsed from the following Build-pane note/include-stack lines.
/// Constructor-initialized class (not a record): the net472 target of this project has no
/// IsExternalInit shim, so init accessors do not compile there.
/// </summary>
public sealed class BuildOutputError(
    string code,
    string message,
    string? file,
    int? line,
    int? column,
    string rawLine,
    string severity = "Error",
    string? originFile = null,
    int? originLine = null,
    int? originColumn = null,
    string? originRawLine = null)
{
    public string Severity { get; } = NormalizeSeverityValue(severity);
    public string Code { get; } = code;
    public string Message { get; } = message;
    public string? File { get; } = file;
    public int? Line { get; } = line;
    public int? Column { get; } = column;
    public string RawLine { get; } = rawLine;
    public string? OriginFile { get; } = originFile;
    public int? OriginLine { get; } = originLine;
    public int? OriginColumn { get; } = originColumn;
    public string? OriginRawLine { get; } = originRawLine;
    public bool HasProjectOrigin => !string.IsNullOrWhiteSpace(OriginFile);

    public BuildOutputError WithProjectOrigin(string originFile, int? originLine, int? originColumn, string originRawLine)
        => new(
            Code,
            Message,
            File,
            Line,
            Column,
            RawLine,
            Severity,
            originFile,
            originLine,
            originColumn,
            originRawLine);

    public string CreateDedupeKey()
        => $"{Severity}|{Code}|{File}|{Line}|{Column}|{Message}|{OriginFile}|{OriginLine}";

    private static string NormalizeSeverityValue(string value)
        => string.Equals(value, "Warning", StringComparison.OrdinalIgnoreCase) ? "Warning" : "Error";
}

/// <summary>
/// Extracts structured diagnostic rows from raw Visual Studio Build output. The Error List does
/// not always receive these rows (linker failures in particular never populate it), so build
/// commands parse the Build pane as a fallback after a failed build and as context for warnings
/// attributed to external headers.
/// </summary>
public static partial class BuildOutputErrorParser
{
    public const int MaxRows = 50;

    private static readonly string[] ProjectSourceExtensions =
    [
        ".c", ".cc", ".cpp", ".cxx", ".h", ".hh", ".hpp", ".hxx", ".ixx", ".inl", ".cs",
    ];

    // "1>path\file.cpp(123,45): warning C4244: message" — compiler/MSBuild diagnostics with a location.
    private const string LocatedDiagnosticPattern =
        @"^\s*(?:\d+>)?\s*(?<file>[^><]+?)\((?<line>\d+)(?:,(?<column>\d+))?\)\s*:\s*(?<severity>fatal error|error|warning)\s+(?<code>[A-Za-z]+[A-Za-z0-9-]*)\s*:\s*(?<message>.+)$";

    // "1>lib.obj : error LNK2001: message" / "Target.dll : fatal error LNK1120: ..." — no location.
    private const string UnlocatedDiagnosticPattern =
        @"^\s*(?:\d+>)?\s*(?<file>[^><]+?)\s*:\s*(?<severity>fatal error|error|warning)\s+(?<code>[A-Za-z]+[A-Za-z0-9-]*)\s*:\s*(?<message>.+)$";

    private const string ProjectOriginLocationPattern =
        @"^\s*(?:\d+>)?\s*(?<file>(?:[A-Za-z]:\\|\\\\|/|\.\.?[\\/])[^()\r\n]+?)\((?<line>\d+)(?:,(?<column>\d+))?\)\s*:\s*(?:(?:note|message)\s*:\s*)?(?<message>.*)$";

    private const string CompilingSourceFilePattern =
        "\\bcompiling source file\\s+['\"](?<file>[^'\"]+)['\"]";

#if NET7_0_OR_GREATER
    [GeneratedRegex(LocatedDiagnosticPattern, RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex LocatedDiagnostic();

    [GeneratedRegex(UnlocatedDiagnosticPattern, RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex UnlocatedDiagnostic();

    [GeneratedRegex(ProjectOriginLocationPattern, RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex ProjectOriginLocation();

    [GeneratedRegex(CompilingSourceFilePattern, RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex CompilingSourceFile();
#else
    private static readonly Regex LocatedDiagnosticCompiled =
        new(LocatedDiagnosticPattern, RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex UnlocatedDiagnosticCompiled =
        new(UnlocatedDiagnosticPattern, RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex ProjectOriginLocationCompiled =
        new(ProjectOriginLocationPattern, RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex CompilingSourceFileCompiled =
        new(CompilingSourceFilePattern, RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static Regex LocatedDiagnostic() => LocatedDiagnosticCompiled;

    private static Regex UnlocatedDiagnostic() => UnlocatedDiagnosticCompiled;

    private static Regex ProjectOriginLocation() => ProjectOriginLocationCompiled;

    private static Regex CompilingSourceFile() => CompilingSourceFileCompiled;
#endif

    public static IReadOnlyList<BuildOutputError> ParseErrors(string? buildOutput)
    {
        List<BuildOutputError> rows = [];
        foreach (BuildOutputError row in ParseDiagnostics(buildOutput))
        {
            if (string.Equals(row.Severity, "Error", StringComparison.OrdinalIgnoreCase))
            {
                rows.Add(row);
            }
        }

        return rows;
    }

    public static IReadOnlyList<BuildOutputError> ParseDiagnostics(string? buildOutput)
    {
        List<BuildOutputError> rows = [];
        if (string.IsNullOrWhiteSpace(buildOutput))
        {
            return rows;
        }

        string[] lines = buildOutput!.Split('\n');
        HashSet<string> seen = new(StringComparer.Ordinal);
        for (int index = 0; index < lines.Length; index++)
        {
            string line = lines[index].TrimEnd('\r');
            if (!LooksLikeDiagnosticLine(line))
            {
                continue;
            }

            BuildOutputError? row = TryParseLine(line);
            if (row is null || IsMsvcContinuationDiagnostic(row))
            {
                continue;
            }

            row = AttachProjectOrigin(row, lines, index + 1);
            if (!seen.Add(row.CreateDedupeKey()))
            {
                continue;
            }

            rows.Add(row);
            if (rows.Count >= MaxRows)
            {
                break;
            }
        }

        return rows;
    }

    private static bool LooksLikeDiagnosticLine(string line)
        => ContainsOrdinalIgnoreCase(line, "error")
            || ContainsOrdinalIgnoreCase(line, "warning");

    private static BuildOutputError? TryParseLine(string line)
    {
        Match located = LocatedDiagnostic().Match(line);
        if (located.Success)
        {
            return CreateDiagnostic(located, line);
        }

        Match unlocated = UnlocatedDiagnostic().Match(line);
        return unlocated.Success ? CreateDiagnostic(unlocated, line) : null;
    }

    private static BuildOutputError CreateDiagnostic(Match match, string rawLine)
    {
        return new BuildOutputError(
            match.Groups["code"].Value,
            match.Groups["message"].Value.Trim(),
            match.Groups["file"].Value.Trim(),
            TryParseInt(match.Groups["line"].Value),
            match.Groups["column"].Success ? TryParseInt(match.Groups["column"].Value) : null,
            rawLine,
            NormalizeParsedSeverity(match.Groups["severity"].Value));
    }

    private static string NormalizeParsedSeverity(string value)
        => ContainsOrdinalIgnoreCase(value, "warning") ? "Warning" : "Error";

    private static bool IsMsvcContinuationDiagnostic(BuildOutputError row)
    {
        if (string.IsNullOrWhiteSpace(row.Code) || row.Code[0] != 'C')
        {
            return false;
        }

        string message = row.Message.Trim();
        return message.Equals("with", StringComparison.OrdinalIgnoreCase)
            || message.Equals("[", StringComparison.Ordinal)
            || message.Equals("]", StringComparison.Ordinal)
            || (message.Length > 0
                && message[0] == '_'
                && message.IndexOf('=') > 0
                && message.IndexOf(' ') < 0
                && message.IndexOf(':') < 0);
    }

    private static bool ContainsOrdinalIgnoreCase(string value, string search)
    {
#if NET7_0_OR_GREATER
        return value.Contains(search, StringComparison.OrdinalIgnoreCase);
#else
        return value.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0;
#endif
    }

    private static int? TryParseInt(string value)
        => int.TryParse(value, out int parsed) ? parsed : null;

    private static BuildOutputError AttachProjectOrigin(BuildOutputError row, IReadOnlyList<string> lines, int startIndex)
    {
        if (!IsExternalBuildPath(row.File))
        {
            return row;
        }

        BuildOutputOrigin? origin = TryFindProjectOrigin(lines, startIndex);
        return origin is null
            ? row
            : row.WithProjectOrigin(origin.File, origin.Line, origin.Column, origin.RawLine);
    }

    private static BuildOutputOrigin? TryFindProjectOrigin(IReadOnlyList<string> lines, int startIndex)
    {
        int scanEnd = Math.Min(lines.Count, startIndex + 80);
        for (int index = startIndex; index < scanEnd; index++)
        {
            string line = lines[index].TrimEnd('\r');
            BuildOutputError? nextDiagnostic = TryParseLine(line);
            if (nextDiagnostic is not null)
            {
                if (!IsMsvcContinuationDiagnostic(nextDiagnostic))
                {
                    return null;
                }

                continue;
            }

            if (TryParseOriginLocation(line, out BuildOutputOrigin? originFromLocation) && originFromLocation is not null)
            {
                return originFromLocation;
            }

            if (TryParseCompilingSourceFile(line, out BuildOutputOrigin? originFromCompilingSource) && originFromCompilingSource is not null)
            {
                return originFromCompilingSource;
            }
        }

        return null;
    }

    private static bool TryParseOriginLocation(string line, out BuildOutputOrigin? origin)
    {
        origin = null;
        Match match = ProjectOriginLocation().Match(line);
        if (!match.Success)
        {
            return false;
        }

        string file = match.Groups["file"].Value.Trim();
        if (!IsProjectOriginPath(file))
        {
            return false;
        }

        origin = new BuildOutputOrigin(
            file,
            TryParseInt(match.Groups["line"].Value),
            match.Groups["column"].Success ? TryParseInt(match.Groups["column"].Value) : null,
            line);
        return true;
    }

    private static bool TryParseCompilingSourceFile(string line, out BuildOutputOrigin? origin)
    {
        origin = null;
        Match match = CompilingSourceFile().Match(line);
        if (!match.Success)
        {
            return false;
        }

        string file = match.Groups["file"].Value.Trim();
        if (!IsProjectOriginPath(file))
        {
            return false;
        }

        origin = new BuildOutputOrigin(file, line: null, column: null, line);
        return true;
    }

    private static bool IsProjectOriginPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        if (IsExternalBuildPath(path))
        {
            return false;
        }

        string normalized = NormalizePathForComparison(path!);
        if (normalized.Contains("/src/", StringComparison.Ordinal))
        {
            return true;
        }

        foreach (string extension in ProjectSourceExtensions)
        {
            if (normalized.EndsWith(extension, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsExternalBuildPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        string normalized = NormalizePathForComparison(path!);
        return normalized.Contains("/program files/", StringComparison.Ordinal)
            || normalized.Contains("/program files (x86)/", StringComparison.Ordinal)
            || normalized.Contains("/microsoft visual studio/", StringComparison.Ordinal)
            || normalized.Contains("/windows kits/", StringComparison.Ordinal)
            || normalized.Contains("/vc/tools/msvc/", StringComparison.Ordinal)
            || normalized.Contains("/ucrt/", StringComparison.Ordinal);
    }

    private static string NormalizePathForComparison(string path)
        => path.Trim().Replace('\\', '/').ToLowerInvariant();

    private sealed class BuildOutputOrigin(string file, int? line, int? column, string rawLine)
    {
        public string File { get; } = file;
        public int? Line { get; } = line;
        public int? Column { get; } = column;
        public string RawLine { get; } = rawLine;
    }
}
