using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using static VsIdeBridge.Diagnostics.BestPracticeAnalyzerHelpers;
using static VsIdeBridge.Diagnostics.ErrorListConstants;
using static VsIdeBridge.Diagnostics.ErrorListPatterns;

namespace VsIdeBridge.Diagnostics;

internal static partial class BestPracticeAnalyzer
{
    // Keep these comments ASCII-only so the analyzer does not flag its own file
    // while still matching the escaped mojibake signatures below.
    // Number of extra lines beyond the match line to scan for an opening brace
    // when distinguishing real method bodies from call expressions in FindLongMethods.
    private const int MethodBraceLookAheadLines = 2;

    private static readonly string[] MojibakeSignatures =
    [
        "\u00C3\u0192",       // Misread UTF-8 prefix sequence.
        "\u00C2\u0080",       // Latin-1 control-character sequence.
        "\u00EF\u00BF\u00BD", // UTF-8 replacement-character sequence.
        "\u00E2\u0082\u00AC", // Misread euro-sign sequence.
        "\u00C3\u0080",       // Another common double-decoding prefix.
        "\u00E2\u0080",       // Smart quote or dash prefix sequence.
        "\u00C3\u0081",       // Another common double-decoding prefix.
        "\u00E2\u0084\u00A2", // Misread trademark-sign sequence.
    ];

    // ── Public entry point ────────────────────────────────────────────────────

    public static IEnumerable<JObject> AnalyzeFile(string file, string content)
    {
        CodeLanguage language = GetLanguage(file);

        // Cross-language rules
        IEnumerable<JObject> findings = FindRepeatedStringLiterals(file, content)
            .Concat(FindMagicNumbers(file, content))
            .Concat(FindFileTooLong(file, content))
            .Concat(FindLongMethods(file, content, language))
            .Concat(FindPoorNaming(file, content, language))
            .Concat(FindDeepNesting(file, content, language))
            .Concat(FindCommentedOutCode(file, content, language))
            .Concat(FindMixedIndentation(file, content))
            .Concat(FindMojibake(file, content))
            .Concat(FindDiagnosticSuppressions(file, content))
            .Concat(FindLongLines(file, content))
            .Concat(FindTodoComments(file, content));

        // Language-specific rules
        if (language == CodeLanguage.CSharp)
        {
            findings = findings
                .Concat(FindImplicitVarUsage(file, content))
                .Concat(FindBroadCatchException(file, content))
                .Concat(FindFrameworkTypeAliases(file, content))
                .Concat(FindLongMainThreadScopes(file, content))
                .Concat(FindSuspiciousRoundDown(file, content))
                .Concat(FindEmptyCatchBlocks(file, content))
                .Concat(FindAsyncVoid(file, content))
                .Concat(FindGodClass(file, content))
                .Concat(FindPropertyBagClass(file, content))
                .Concat(FindDateTimeInLoop(file, content))
                .Concat(FindDynamicObjectOveruse(file, content))
                .Concat(FindUnnecessaryComments(file, content))
                .Concat(FindNamespaceFolderStructureIssues(file, content));
        }
        else if (language == CodeLanguage.Cpp)
        {
            findings = findings
                .Concat(FindRawDelete(file, content))
                .Concat(FindCStyleCasts(file, content))
                .Concat(FindRawNew(file, content))
                .Concat(FindMacroOveruse(file, content))
                .Concat(FindDeepInheritance(file, content))
                .Concat(FindMissingConst(file, content));
            if (IsHeaderFile(file))
            {
                findings = findings.Concat(FindUsingNamespaceInHeader(file, content));
            }
        }
        else if (language == CodeLanguage.Python)
        {
            findings = findings
                .Concat(FindBareExcept(file, content))
                .Concat(FindMutableDefaultArgs(file, content))
                .Concat(FindImportStar(file, content))
                .Concat(FindBooleanComparison(file, content))
                .Concat(FindNoneEqualityComparison(file, content));
        }
        else if (language == CodeLanguage.VisualBasic)
        {
            findings = findings
                .Concat(FindMissingOptionStrict(file, content))
                .Concat(FindVbMultipleStatementsPerLine(file, content))
                .Concat(FindVbExplicitLineContinuation(file, content));
        }
        else if (language == CodeLanguage.FSharp)
        {
            findings = findings
                .Concat(FindFSharpMutableState(file, content))
                .Concat(FindFSharpBlockComments(file, content));
        }
        else if (language == CodeLanguage.PowerShell)
        {
            findings = findings
                .Concat(FindWriteHostUsage(file, content))
                .Concat(FindMissingStrictMode(file, content))
                .Concat(FindPowerShellAliases(file, content));
        }

        return findings;
    }

    // ── Language detection ────────────────────────────────────────────────────

    public static CodeLanguage GetLanguage(string filePath)
    {
        string ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ext switch
        {
            ".cs" => CodeLanguage.CSharp,
            ".vb" => CodeLanguage.VisualBasic,
            ".fs" or ".fsi" or ".fsx" => CodeLanguage.FSharp,
            ".c" or ".cc" or ".cpp" or ".cxx" or ".h" or ".hh" or ".hpp" or ".hxx" => CodeLanguage.Cpp,
            ".py" => CodeLanguage.Python,
            ".ps1" or ".psm1" or ".psd1" => CodeLanguage.PowerShell,
            _ => CodeLanguage.Unknown,
        };
    }

    public static bool IsHeaderFile(string filePath)
    {
        string ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ext is ".h" or ".hh" or ".hpp" or ".hxx";
    }

    // ── BP1001: Repeated string literals ──────────────────────────────────────

    public static IEnumerable<JObject> FindRepeatedStringLiterals(string file, string content)
    {
        // Test files use inline string literals by convention — test inputs and expected values
        // are deliberately repeated across test methods and should not be extracted to constants.
        if (IsTestFile(file))
        {
            yield break;
        }

        CodeLanguage language = GetLanguage(file);
        IEnumerable<IGrouping<string, Match>> occurrences = StringLiteralPattern().Matches(content)
            .Cast<Match>()
            .Where(m => !IsInsideStringLiteral(content, m.Index))
            .Where(m => !IsInsideComment(content, m.Index, language))
            .Where(m => !ConstStringDeclPattern().IsMatch(GetLineAt(content, m.Index)))
            .GroupBy(match => match.Groups[1].Value)
            .Where(group => group.Count() >= RepeatedStringThreshold)
            .OrderByDescending(group => group.Count())
            .ThenBy(group => group.Key, StringComparer.Ordinal)
            .Take(MaxBestPracticeFindingsPerFile);

        foreach (IGrouping<string, Match> repeated in occurrences)
        {
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1001",
                message: $"String literal '{repeated.Key}' is repeated {repeated.Count()} times. Extract a constant.",
                file: file,
                line: GetLineNumber(content, repeated.First().Index),
                symbol: repeated.Key,
                helpUri: BP1001HelpUri);
        }
    }

    // ── BP1002: Magic numbers ─────────────────────────────────────────────────

    public static IEnumerable<JObject> FindMagicNumbers(string file, string content)
    {
        // Same rationale as BP1001: test methods use literal numbers as inputs/assertions.
        if (IsTestFile(file))
        {
            yield break;
        }

        CodeLanguage language = GetLanguage(file);
        IEnumerable<IGrouping<string, (Match Match, string Value)>> matches =
            NumberLiteralPattern().Matches(content)
            .Cast<Match>()
            .Select(match => (Match: match, match.Groups["value"].Value))
            .Where(item => item.Value is not "0" and not "1" and not "-1")
            .Where(item => !IsInsideStringLiteral(content, item.Match.Index))
            .Where(item => !IsInsideComment(content, item.Match.Index, language))
            .GroupBy(item => item.Value)
            .Where(group => group.Count() >= RepeatedNumberThreshold)
            .OrderByDescending(group => group.Count())
            .ThenBy(group => group.Key, StringComparer.Ordinal)
            .Take(MaxBestPracticeFindingsPerFile);

        foreach (IGrouping<string, (Match Match, string Value)> repeated in matches)
        {
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1002",
                message: $"Numeric literal '{repeated.Key}' appears {repeated.Count()} times. Use a named constant when it carries domain meaning, or add a short comment when the value is only local arithmetic.",
                file: file,
                line: GetLineNumber(content, repeated.First().Match.Index),
                symbol: repeated.Key,
                helpUri: BP1002HelpUri);
        }
    }

    // ── BP1003: Suspicious round-down cast (C#) ───────────────────────────────

    public static IEnumerable<JObject> FindSuspiciousRoundDown(string file, string content)
    {
        string[] lines = content.Split(["\r\n", "\n"], StringSplitOptions.None);
        int findingCount = 0;
        for (int i = 0; i < lines.Length; i++)
        {
            Match roundDownMatch = SuspiciousRoundDownPattern().Match(lines[i]);
            if (roundDownMatch.Success)
            {
                yield return DiagnosticRowFactory.CreateBestPracticeRow(
                    code: "BP1003",
                    message: $"(int)Math.{roundDownMatch.Groups["op"].Value}(...) casts a float floor/truncate result to int. This pattern usually means integer division (/) is what you want instead.",
                    file: file,
                    line: i + 1,
                    symbol: $"Math.{roundDownMatch.Groups["op"].Value}",
                    helpUri: BP1003HelpUri);
                findingCount++;
                if (findingCount >= MaxSuppressionFindingsPerFile)
                {
                    yield break;
                }
            }
        }
    }

    // ── BP1005: async void (C#) ───────────────────────────────────────────────

    public static IEnumerable<JObject> FindAsyncVoid(string file, string content)
    {
        MatchCollection matches = AsyncVoidPattern().Matches(content);
        int findingCount = 0;
        foreach (Match match in matches)
        {
            if (IsInsideStringLiteral(content, match.Index))
                continue;
            string methodName = match.Groups[1].Value;
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1005",
                message: $"'async void {methodName}' has unobservable exceptions. Use 'async Task' instead.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: methodName,
                helpUri: BP1005HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile)
            {
                yield break;
            }
        }
    }

    // ── BP1006: Raw delete (C++) ──────────────────────────────────────────────

    public static IEnumerable<JObject> FindRawDelete(string file, string content)
    {
        MatchCollection matches = RawDeletePattern().Matches(content);
        int findingCount = 0;
        foreach (Match match in matches)
        {
            string op = match.Groups[1].Success ? "delete[]" : "delete";
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1006",
                message: $"Raw '{op}' detected. Prefer smart pointers (std::unique_ptr, std::shared_ptr).",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: op,
                helpUri: BP1006HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile)
            {
                yield break;
            }
        }
    }

    // ── BP1007: using namespace in header (C++) ───────────────────────────────

    public static IEnumerable<JObject> FindUsingNamespaceInHeader(string file, string content)
    {
        MatchCollection matches = UsingNamespacePattern().Matches(content);
        int findingCount = 0;
        foreach (Match match in matches)
        {
            string ns = match.Groups[1].Value;
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1007",
                message: $"'using namespace {ns}' in header file pollutes the global namespace of every includer.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: ns,
                helpUri: BP1007HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile)
            {
                yield break;
            }
        }
    }

    // ── BP1008: C-style cast (C++) ────────────────────────────────────────────

    // FindCStyleCasts lives in BestPracticeAnalyzerStructureRules.cs with the other
    // comment/string-guarded rules (this file is at the BP1012 length limit).

    // ── BP1009: Bare except (Python) ──────────────────────────────────────────

    public static IEnumerable<JObject> FindBareExcept(string file, string content)
    {
        MatchCollection matches = BareExceptPattern().Matches(content);
        int findingCount = 0;
        foreach (Match match in matches)
        {
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1009",
                message: "Bare 'except:' catches all exceptions including SystemExit and KeyboardInterrupt. Catch specific exceptions.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: "except",
                helpUri: BP1009HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile)
            {
                yield break;
            }
        }
    }

    // ── BP1010: Mutable default argument (Python) ─────────────────────────────

    public static IEnumerable<JObject> FindMutableDefaultArgs(string file, string content)
    {
        MatchCollection matches = MutableDefaultArgPattern().Matches(content);
        int findingCount = 0;
        foreach (Match match in matches)
        {
            string funcName = match.Groups[1].Value;
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1010",
                message: $"Function '{funcName}' uses a mutable default argument. Use None and initialize inside the function.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: funcName,
                helpUri: BP1010HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile)
            {
                yield break;
            }
        }
    }

    // ── BP1011: from X import * (Python) ──────────────────────────────────────

    public static IEnumerable<JObject> FindImportStar(string file, string content)
    {
        MatchCollection matches = ImportStarPattern().Matches(content);
        int findingCount = 0;
        foreach (Match match in matches)
        {
            string module = match.Groups[1].Value;
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1011",
                message: $"'from {module} import *' pollutes the namespace. Import specific names.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: module,
                helpUri: BP1011HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile)
            {
                yield break;
            }
        }
    }

    // ── BP1012: File too long ─────────────────────────────────────────────────

    public static IEnumerable<JObject> FindFileTooLong(string file, string content)
    {
        int lineCount = content.Split('\n').Length;
        if (lineCount >= FileTooLongErrorThreshold)
        {
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1012",
                message: $"This file has grown to {lineCount} lines — well past the {FileTooLongErrorThreshold}-line limit. Split it into smaller focused types or extract a class library.",
                file: file,
                line: 1,
                symbol: Path.GetFileName(file),
                helpUri: BP1012HelpUri);
        }
        else if (lineCount >= FileTooLongWarningThreshold)
        {
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1012",
                message: $"This file is {lineCount} lines long, approaching the {FileTooLongWarningThreshold}-line limit. Consider splitting it into smaller focused types or a separate class library.",
                file: file,
                line: 1,
                symbol: Path.GetFileName(file),
                helpUri: BP1012HelpUri);
        }
    }

    // ── BP1013: Method/function too long ──────────────────────────────────────

    public static IEnumerable<JObject> FindLongMethods(string file, string content, CodeLanguage language)
    {
        string[] lines = content.Split('\n');
        Regex? pattern = language switch
        {
            CodeLanguage.CSharp => CSharpMethodSignaturePattern(),
            CodeLanguage.Python => PythonDefPattern(),
            CodeLanguage.Cpp => CppFunctionPattern(),
            _ => null,
        };
        if (pattern is null)
        {
            yield break;
        }

        int findingCount = 0;
        MatchCollection matches = pattern.Matches(content);
        foreach (Match match in matches)
        {
            string methodName = match.Groups[1].Value;
            int startLine = BestPracticeAnalyzerHelpers.GetLineNumber(content, match.Index);
            if (language == CodeLanguage.CSharp && BestPracticeAnalyzerHelpers.IsGeneratedRegexDeclaration(lines, startLine))
            {
                continue;
            }

            int methodLength;
            if (language == CodeLanguage.Python)
            {
                methodLength = BestPracticeAnalyzerHelpers.CountPythonFunctionLines(lines, startLine - 1);
            }
            else
            {
                if (!BestPracticeAnalyzerHelpers.HasNearbyOpeningBrace(lines, startLine, MethodBraceLookAheadLines))
                {
                    continue;
                }

                methodLength = BestPracticeAnalyzerHelpers.CountBracedBlockLines(lines, startLine - 1);
            }

            if (methodLength > MethodTooLongThreshold)
            {
                yield return DiagnosticRowFactory.CreateBestPracticeRow(
                    code: "BP1013",
                    message: $"Method '{methodName}' is {methodLength} lines long (threshold: {MethodTooLongThreshold}). Break into smaller methods.",
                    file: file,
                    line: startLine,
                    symbol: methodName,
                    helpUri: BP1013HelpUri);
                findingCount++;
                if (findingCount >= MaxSuppressionFindingsPerFile)
                {
                    yield break;
                }
            }
        }
    }

    // ── BP1014: Poor/vague naming ─────────────────────────────────────────────

    public static IEnumerable<JObject> FindPoorNaming(string file, string content, CodeLanguage language)
    {
        int findingCount = 0;
        if (language == CodeLanguage.CSharp)
        {
            foreach (Match match in SingleLetterVarPattern().Matches(content))
            {
                // Skip declaration-shaped snippets inside comments or C# string literals
                // (common in analyzer tests) so only real local variables are flagged.
                if (IsInsideComment(content, match.Index, language) || IsInsideStringLiteral(content, match.Index))
                {
                    continue;
                }

                string name = match.Groups["name"].Value;
                if (name is "i" or "j" or "k" or "m" or "n" or "x" or "y" or "z" or "e" or "s" or "_")
                {
                    continue;
                }
                yield return DiagnosticRowFactory.CreateBestPracticeRow(
                    code: "BP1014",
                    message: $"Single-letter variable '{name}' is unclear. Use a descriptive name.",
                    file: file,
                    line: GetLineNumber(content, match.Index),
                    symbol: name,
                    helpUri: BP1014HelpUri);
                findingCount++;
                if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
            }

            foreach (Match match in PoorCSharpNamingPattern().Matches(content))
            {
                if (IsInsideComment(content, match.Index, language) || IsInsideStringLiteral(content, match.Index))
                {
                    continue;
                }

                string name = match.Groups["name"].Value;
                yield return DiagnosticRowFactory.CreateBestPracticeRow(
                    code: "BP1014",
                    message: $"Vague variable name '{name}'. Use a name that describes the value's purpose.",
                    file: file,
                    line: GetLineNumber(content, match.Index),
                    symbol: name,
                    helpUri: BP1014HelpUri);
                findingCount++;
                if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
            }
        }
        else if (language == CodeLanguage.Python)
        {
            foreach (Match match in PythonSingleLetterAssignPattern().Matches(content))
            {
                string name = match.Groups["name"].Value;
                if (name is "i" or "j" or "k" or "m" or "n" or "x" or "y" or "z" or "e" or "s" or "_")
                {
                    continue;
                }
                yield return DiagnosticRowFactory.CreateBestPracticeRow(
                    code: "BP1014",
                    message: $"Single-letter variable '{name}' is unclear. Use a descriptive name.",
                    file: file,
                    line: GetLineNumber(content, match.Index),
                    symbol: name,
                    helpUri: BP1014HelpUri);
                findingCount++;
                if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
            }

            foreach (Match match in PythonPoorNamingPattern().Matches(content))
            {
                string name = match.Groups["name"].Value;
                yield return DiagnosticRowFactory.CreateBestPracticeRow(
                    code: "BP1014",
                    message: $"Vague variable name '{name}'. Use a name that describes the value's purpose.",
                    file: file,
                    line: GetLineNumber(content, match.Index),
                    symbol: name,
                    helpUri: BP1014HelpUri);
                findingCount++;
                if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
            }
        }
    }

    // ── BP1033: implicit var usage (C#) ───────────────────────────────────────

    public static IEnumerable<JObject> FindImplicitVarUsage(string file, string content)
    {
        int findingCount = 0;
        foreach (Match match in ImplicitVarPattern().Matches(content))
        {
            if (IsInsideStringLiteral(content, match.Index) || IsInsideLineComment(content, match.Index))
            {
                continue;
            }

            string name = match.Groups["name"].Value;
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1033",
                message: $"Implicitly typed local '{name}' uses 'var'. Prefer the explicit type unless the concrete type would be excessively noisy.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: name,
                helpUri: BP1033HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile)
            {
                yield break;
            }
        }
    }

    public static IEnumerable<JObject> FindBroadCatchException(string file, string content)
    {
        int findingCount = 0;
        foreach (Match match in BroadCatchPattern().Matches(content))
        {
            if (IsInsideStringLiteral(content, match.Index) || IsInsideLineComment(content, match.Index))
            {
                continue;
            }

            if (IsTopLevelProgramCatch(file, content, match.Index))
            {
                continue;
            }

            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1034",
                message: "Catching general Exception makes failures harder to reason about. Catch a narrower exception type or let the failure propagate.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: "Exception",
                helpUri: BP1034HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile)
            {
                yield break;
            }
        }
    }

    public static IEnumerable<JObject> FindFrameworkTypeAliases(string file, string content)
    {
        int findingCount = 0;
        foreach (Match match in FrameworkTypePattern().Matches(content))
        {
            if (IsInsideStringLiteral(content, match.Index) || IsInsideLineComment(content, match.Index))
            {
                continue;
            }

            string typeName = match.Groups["type"].Value;
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1035",
                message: $"Use the C# keyword form instead of 'System.{typeName}' for built-in types where possible.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: typeName,
                helpUri: BP1035HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile)
            {
                yield break;
            }
        }
    }

    public static IEnumerable<JObject> FindLongMainThreadScopes(string file, string content)
    {
        string[] lines = content.Split('\n');
        int findingCount = 0;

        foreach (Match match in CSharpMethodSignaturePattern().Matches(content))
        {
            string methodName = match.Groups[1].Value;
            int startLine = BestPracticeAnalyzerHelpers.GetLineNumber(content, match.Index);

            if (!BestPracticeAnalyzerHelpers.HasNearbyOpeningBrace(lines, startLine, MethodBraceLookAheadLines))
            {
                continue;
            }

            int methodLength = BestPracticeAnalyzerHelpers.CountBracedBlockLines(lines, startLine - 1);
            int methodEndExclusive = Math.Min(lines.Length, startLine - 1 + methodLength);
            int switchLineIndex = -1;
            int nonEmptyLinesBeforeSwitch = 0;

            for (int i = startLine - 1; i < methodEndExclusive; i++)
            {
                string trimmedLine = lines[i].Trim();
                if (trimmedLine.Length > 0)
                {
                    nonEmptyLinesBeforeSwitch++;
                }

#if NET5_0_OR_GREATER
                if (trimmedLine.Contains("SwitchToMainThreadAsync", StringComparison.Ordinal))
#else
                if (trimmedLine.IndexOf("SwitchToMainThreadAsync", StringComparison.Ordinal) >= 0)
#endif
                {
                    switchLineIndex = i;
                    break;
                }
            }

            if (switchLineIndex < 0 || nonEmptyLinesBeforeSwitch > MainThreadSwitchEarlyLineThreshold)
            {
                continue;
            }

            int nonEmptyLinesAfterSwitch = 0;
            for (int i = switchLineIndex + 1; i < methodEndExclusive; i++)
            {
                if (!string.IsNullOrWhiteSpace(lines[i]))
                {
                    nonEmptyLinesAfterSwitch++;
                }
            }

            if (nonEmptyLinesAfterSwitch < MainThreadScopeWarningThreshold)
            {
                continue;
            }

            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1043",
                message: $"Method '{methodName}' switches to the Visual Studio UI thread early and then keeps {nonEmptyLinesAfterSwitch} non-empty lines in that scope. Narrow the main-thread region.",
                file: file,
                line: switchLineIndex + 1,
                symbol: methodName,
                helpUri: BP1043HelpUri);

            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile)
            {
                yield break;
            }
        }
    }

    // ── BP1015: Excessive nesting depth ───────────────────────────────────────

    public static IEnumerable<JObject> FindDeepNesting(string file, string content, CodeLanguage language)
    {
        return language == CodeLanguage.Python
            ? FindPythonDeepNesting(file, content)
            : FindCSharpDeepNesting(file, content);
    }

    private static IEnumerable<JObject> FindPythonDeepNesting(string file, string content)
    {
        string[] lines = content.Split('\n');
        HashSet<int> reported = [];

        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i];
            if (string.IsNullOrWhiteSpace(line)) { continue; }

            int indentLevel = (line.Length - line.TrimStart().Length) / 4;
            if (indentLevel >= DeepNestingThreshold && reported.Add(i))
            {
                yield return DiagnosticRowFactory.CreateBestPracticeRow(
                    code: "BP1015",
                    message: $"Code is nested {indentLevel} levels deep (threshold: {DeepNestingThreshold}). " +
                             "Consider extracting a method or using guard clauses to reduce nesting.",
                    file: file,
                    line: i + 1,
                    symbol: $"Nesting depth {indentLevel}",
                    helpUri: BP1015HelpUri);

                if (reported.Count >= MaxFindingsPerFile) { yield break; }
            }
        }
    }

    private static IEnumerable<JObject> FindCSharpDeepNesting(string file, string content)
    {
        string fileName = Path.GetFileName(file);
        if (fileName.StartsWith("ErrorListPatterns", StringComparison.OrdinalIgnoreCase))
        {
            yield break;
        }

        HashSet<int> reported = [];
        int currentDepth = 0;
        string[] lines = content.Split('\n');

        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i].Trim();
            currentDepth += CountStructuralBraceDelta(line);
            if (currentDepth < 0) { currentDepth = 0; }

            if (currentDepth >= CSharpDeepNestingThreshold && reported.Add(i))
            {
                yield return DiagnosticRowFactory.CreateBestPracticeRow(
                    code: "BP1015",
                    message: $"Code is nested {currentDepth} levels deep (threshold: {CSharpDeepNestingThreshold}). " +
                             "Consider extracting a method, using early returns/guard clauses, or restructuring the logic.",
                    file: file,
                    line: i + 1,
                    symbol: $"Nesting depth {currentDepth}",
                    helpUri: BP1015HelpUri);

                if (reported.Count >= 5) { yield break; }
            }
        }
    }

    // ── BP1016: Commented-out code blocks ─────────────────────────────────────

    public static IEnumerable<JObject> FindCommentedOutCode(string file, string content, CodeLanguage language)
    {
        Regex? pattern = language switch
        {
            CodeLanguage.CSharp => CSharpCommentedCodePattern(),
            CodeLanguage.Python => PythonCommentedCodePattern(),
            CodeLanguage.Cpp => CppCommentedCodePattern(),
            _ => null,
        };
        if (pattern is null) { yield break; }

        string[] lines = content.Split('\n');
        int findingCount = 0;
        int consecutiveCount = 0;
        int blockStart = -1;

        for (int i = 0; i < lines.Length; i++)
        {
            if (pattern.IsMatch(lines[i]))
            {
                if (consecutiveCount == 0) { blockStart = i; }
                consecutiveCount++;
            }
            else
            {
                if (consecutiveCount >= CommentedOutCodeThreshold)
                {
                    yield return DiagnosticRowFactory.CreateBestPracticeRow(
                        code: "BP1016",
                        message: $"{consecutiveCount} consecutive lines of commented-out code. Remove dead code; use version control to recover it.",
                        file: file,
                        line: blockStart + 1,
                        symbol: "commented-code",
                        helpUri: BP1016HelpUri);
                    findingCount++;
                    if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
                }
                consecutiveCount = 0;
            }
        }

        if (consecutiveCount >= CommentedOutCodeThreshold)
        {
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1016",
                message: $"{consecutiveCount} consecutive lines of commented-out code at end of file. Remove dead code; use version control to recover it.",
                file: file,
                line: blockStart + 1,
                symbol: "commented-code",
                helpUri: BP1016HelpUri);
        }
    }

    // ── BP1017: Mixed indentation ─────────────────────────────────────────────

    public static IEnumerable<JObject> FindMixedIndentation(string file, string content)
    {
        bool hasTabs = TabIndentedLinePattern().IsMatch(content);
        bool hasSpaces = SpaceIndentedLinePattern().IsMatch(content);
        if (hasTabs && hasSpaces)
        {
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1017",
                message: "File mixes tabs and spaces for indentation. Pick one and be consistent.",
                file: file,
                line: 1,
                symbol: "indentation",
                helpUri: BP1017HelpUri);
        }
    }

    // ── BP1018: God class (C#) ────────────────────────────────────────────────

    public static IEnumerable<JObject> FindGodClass(string file, string content)
    {
        MatchCollection classMatches = CSharpClassDeclPattern().Matches(content);
        int findingCount = 0;

        foreach (Match classMatch in classMatches)
        {
            string className = classMatch.Groups[1].Value;
            // Partial classes spread methods across files; per-file counts are unreliable. Skip them.
#if NET5_0_OR_GREATER
            if (classMatch.Value.Contains("partial", StringComparison.Ordinal))
#else
            if (classMatch.Value.IndexOf("partial", StringComparison.Ordinal) >= 0)
#endif
                continue;
            int classStartLine = BestPracticeAnalyzerHelpers.GetLineNumber(content, classMatch.Index);
            string[] lines = content.Split('\n');
            string classBody = BestPracticeAnalyzerHelpers.ExtractBracedBlock(lines, classStartLine - 1);

            int methodCount = CSharpMethodSignaturePattern().Matches(classBody).Count;
            int fieldCount = CSharpFieldDeclPattern().Matches(classBody).Count;

            if (methodCount >= GodClassMethodThreshold)
            {
                yield return DiagnosticRowFactory.CreateBestPracticeRow(
                    code: "BP1018",
                    message: $"Class '{className}' has {methodCount} methods (threshold: {GodClassMethodThreshold}). Split responsibilities by extracting focused collaborators or helper services so this type stops owning unrelated workflows.",
                    file: file,
                    line: classStartLine,
                    symbol: className,
                    helpUri: BP1018HelpUri);
                findingCount++;
            }

            if (fieldCount >= GodClassFieldThreshold)
            {
                // Static classes hold only class-level state (constants, cached patterns, etc.).
                // High field counts in static classes are typical of catalog/registry designs -- skip.
#if NET5_0_OR_GREATER
                bool isStaticClass = classMatch.Value.Contains("static ", StringComparison.Ordinal);
#else
                bool isStaticClass = classMatch.Value.IndexOf("static ", StringComparison.Ordinal) >= 0;
#endif
                if (!isStaticClass)
                {
                    yield return DiagnosticRowFactory.CreateBestPracticeRow(
                        code: "BP1018",
                        message: $"Class '{className}' has {fieldCount} fields (threshold: {GodClassFieldThreshold}). Split state into smaller focused objects so this type stops carrying unrelated responsibilities.",
                        file: file,
                        line: classStartLine,
                        symbol: className,
                        helpUri: BP1018HelpUri);
                    findingCount++;
                }
            }

            if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
        }
    }

    

    



    public static IEnumerable<JObject> FindVbMultipleStatementsPerLine(string file, string content)
    {
        string[] lines = content.Split('\n');
        int findingCount = 0;
        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i];
            int commentIndex = line.IndexOf('\'');
#if NET5_0_OR_GREATER
            string code = commentIndex >= 0 ? line[..commentIndex] : line;
#else
            string code = commentIndex >= 0 ? line.Substring(0, commentIndex) : line;
#endif
            if (code.IndexOf(':') < 0)
            {
                continue;
            }

            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1037",
                message: "Visual Basic line contains multiple statements separated by ':'. Prefer one statement per line for readability.",
                file: file,
                line: i + 1,
                symbol: ":",
                helpUri: BP1037HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile)
            {
                yield break;
            }
        }
    }

    public static IEnumerable<JObject> FindVbExplicitLineContinuation(string file, string content)
    {
        int findingCount = 0;
        foreach (Match match in VbLineContinuationPattern().Matches(content))
        {
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1038",
                message: "Visual Basic file uses explicit line continuation '_'. Prefer implicit line continuation where the language already supports it.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: "_",
                helpUri: BP1038HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile)
            {
                yield break;
            }
        }
    }


}
