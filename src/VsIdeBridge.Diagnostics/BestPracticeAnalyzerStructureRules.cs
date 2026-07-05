using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using static VsIdeBridge.Diagnostics.BestPracticeAnalyzerHelpers;
using static VsIdeBridge.Diagnostics.ErrorListConstants;
using static VsIdeBridge.Diagnostics.ErrorListPatterns;

namespace VsIdeBridge.Diagnostics;

internal static partial class BestPracticeAnalyzer
{
    // ── BP1044: Diagnostic suppression ─────────────────────────────────────────

    public static IEnumerable<JObject> FindDiagnosticSuppressions(string file, string content)
    {
        int findingCount = 0;

        foreach (JObject finding in FindPragmaWarningSuppressions(file, content))
        {
            yield return finding;
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
        }

        foreach (JObject finding in FindNativeDiagnosticSuppressions(file, content))
        {
            yield return finding;
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
        }

        foreach (Match match in EditorConfigDiagnosticSuppressionPattern().Matches(content))
        {
            string diagnosticCode = match.Groups["code"].Value;
            string severity = match.Groups["severity"].Value;
            string symbol = $"dotnet_diagnostic.{diagnosticCode}.severity = {severity}";
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1044",
                message: $"Diagnostic suppression '{symbol}' hides analyzer output instead of fixing the root cause. Restore a visible severity and address the underlying diagnostic.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: symbol,
                helpUri: BP1044HelpUri);

            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
        }

        foreach (Match match in NoWarnElementPattern().Matches(content))
        {
            string codes = match.Groups["codes"].Value.Trim();
            string symbol = $"NoWarn={codes}";
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1044",
                message: $"Project-level suppression '{symbol}' hides diagnostics instead of fixing the root cause. Remove NoWarn entries and repair the underlying diagnostics.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: symbol,
                helpUri: BP1044HelpUri);

            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
        }

        foreach (Match match in NoWarnAttributePattern().Matches(content))
        {
            string codes = match.Groups["codes"].Value.Trim();
            string symbol = $"NoWarn={codes}";
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1044",
                message: $"Project-level suppression '{symbol}' hides diagnostics instead of fixing the root cause. Remove NoWarn entries and repair the underlying diagnostics.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: symbol,
                helpUri: BP1044HelpUri);

            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
        }

        foreach (Match match in SuppressMessagePattern().Matches(content))
        {
            string suppressionText = GetLineAt(content, match.Index).Trim();
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1044",
                message: $"Suppression attribute '{Truncate(suppressionText, 96)}' hides diagnostics instead of fixing the root cause. Remove the attribute and address the underlying diagnostic.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: "SuppressMessage",
                helpUri: BP1044HelpUri);

            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
        }

        foreach (Match match in RuleSetSuppressionPattern().Matches(content))
        {
            string diagnosticCode = match.Groups["code"].Value;
            string action = match.Groups["action"].Value;
            string symbol = $"Rule {diagnosticCode} Action={action}";
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1044",
                message: $"Ruleset suppression '{symbol}' hides diagnostics instead of fixing the root cause. Restore a visible action and address the underlying diagnostic.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: symbol,
                helpUri: BP1044HelpUri);

            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
        }
    }

    private static IEnumerable<JObject> FindPragmaWarningSuppressions(string file, string content)
    {
        foreach (Match match in PragmaWarningDisablePattern().Matches(content))
        {
            string pragmaText = GetLineAt(content, match.Index).Trim();
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1044",
                message: $"In-source warning suppression '{pragmaText}' hides diagnostics instead of fixing the root cause. Remove the pragma and address the underlying warning.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: pragmaText,
                helpUri: BP1044HelpUri);
        }
    }

    private static IEnumerable<JObject> FindNativeDiagnosticSuppressions(string file, string content)
    {
        foreach (Match match in NativeDiagnosticSuppressionPattern().Matches(content))
        {
            string suppressionText = GetLineAt(content, match.Index).Trim();
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1044",
                message: $"Native diagnostic suppression '{Truncate(suppressionText, 96)}' hides compiler output instead of fixing the root cause. Remove the suppression and address the underlying warning.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: suppressionText,
                helpUri: BP1044HelpUri);
        }
    }

    // ── BP1004: Empty or comment-only catch block (C#) ───────────────────────

    public static IEnumerable<JObject> FindEmptyCatchBlocks(string file, string content)
    {
        string[] lines = content.Split('\n');
        MatchCollection matches = CatchBlockPattern().Matches(content);
        int findingCount = 0;
        foreach (Match match in matches)
        {
            if (!TryGetCommentOnlyCatchBlockInfo(lines, content, match.Index, out int startLine, out string message))
            {
                continue;
            }

            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1004",
                message: message,
                file: file,
                line: startLine,
                symbol: "catch",
                helpUri: BP1004HelpUri);

            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
        }
    }

    // ── BP1045: TODO/FIXME-style marker comments left in code ──────────────────

    public static IEnumerable<JObject> FindTodoComments(string file, string content)
    {
        int findingCount = 0;
        foreach (Match match in TodoCommentPattern().Matches(content))
        {
            string marker = match.Groups["marker"].Value.ToUpperInvariant();
            string markerText = match.Groups["text"].Value.Trim();
            string message = string.IsNullOrWhiteSpace(markerText)
                ? $"{marker} comment found. Track the work in an issue or resolve it before shipping."
                : $"{marker} comment found: \"{Truncate(markerText, 72)}\". Track the work in an issue or resolve it before shipping.";

            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1045",
                message: message,
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: marker,
                helpUri: BP1045HelpUri);

            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
        }
    }

    // ── BP1020: DateTime.Now/UtcNow in loops (C#) ─────────────────────────────

    public static IEnumerable<JObject> FindDateTimeInLoop(string file, string content)
    {
        string[] lines = content.Split('\n');
        int findingCount = 0;
        bool inLoop = false;
        int loopBraceDepth = 0;
        int braceDepth = 0;

        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i];
#if NET5_0_OR_GREATER
            if (LoopKeywordPattern().IsMatch(line))
#else
            if (Regex.IsMatch(line, @"\b(?:for|foreach|while)\s*\("))
#endif
            {
                inLoop = true;
                loopBraceDepth = braceDepth;
            }

            foreach (char ch in line)
            {
                if (ch == '{') { braceDepth++; }
                else if (ch == '}') { braceDepth--; }
            }

            if (inLoop && braceDepth <= loopBraceDepth)
            {
                inLoop = false;
            }

            if (inLoop)
            {
                Match dtMatch = DateTimeNowSimplePattern().Match(line);
                if (dtMatch.Success)
                {
                    yield return DiagnosticRowFactory.CreateBestPracticeRow(
                        code: "BP1020",
                        message: $"DateTime.{dtMatch.Groups["prop"].Value} called inside a loop. Capture it once before the loop.",
                        file: file,
                        line: i + 1,
                        symbol: $"DateTime.{dtMatch.Groups["prop"].Value}",
                        helpUri: BP1020HelpUri);
                    findingCount++;
                    if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
                }
            }
        }
    }

    // ── BP1021: Overuse of dynamic/object (C#) ────────────────────────────────

    public static IEnumerable<JObject> FindDynamicObjectOveruse(string file, string content)
    {
        MatchCollection matches = DynamicObjectParamPattern().Matches(content);
        if (matches.Count >= DynamicObjectThreshold)
        {
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1021",
                message: $"'dynamic' or 'object' used as parameter type {matches.Count} times. Use specific types or generics for type safety.",
                file: file,
                line: GetLineNumber(content, matches[0].Index),
                symbol: "dynamic/object",
                helpUri: BP1021HelpUri);
        }
    }

    // ── BP1022: Raw new without smart pointer (C++) ───────────────────────────

    // ── BP1008: C-style cast (C/C++) ──

    public static IEnumerable<JObject> FindCStyleCasts(string file, string content)
    {
        CodeLanguage language = GetLanguage(file);
        // C++ named casts do not exist in C, so .c files get C-appropriate guidance.
        bool isCSource = file.EndsWith(".c", StringComparison.OrdinalIgnoreCase);
        int findingCount = 0;
        foreach (Match match in CStyleCastPattern().Matches(content))
        {
            // Skip cast-shaped text inside comments or string literals (e.g. "image(float)"
            // in documentation prose) so only executable casts are flagged.
            if (IsInsideComment(content, match.Index, language) || IsInsideStringLiteral(content, match.Index))
            {
                continue;
            }

            string guidance = isCSource
                ? "Avoid the cast by using the correct type or printf format specifier; C has no named casts."
                : "Prefer static_cast, reinterpret_cast, or const_cast.";
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1008",
                message: $"C-style cast '{match.Value.TrimEnd()}' detected. {guidance}",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: match.Value.Trim(),
                helpUri: BP1008HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile)
            {
                yield break;
            }
        }
    }

    public static IEnumerable<JObject> FindRawNew(string file, string content)
    {
        CodeLanguage language = GetLanguage(file);
        int findingCount = 0;
        foreach (Match match in RawNewPattern().Matches(content))
        {
            // Skip "new" inside comments or string literals (e.g. the English word in
            // "creates a new edge") so only real C++ allocations are flagged.
            if (IsInsideComment(content, match.Index, language) || IsInsideStringLiteral(content, match.Index))
            {
                continue;
            }

            string line = BestPracticeAnalyzerHelpers.GetLineAt(content, match.Index);
            if (line.Contains("make_unique") || line.Contains("make_shared") || line.Contains("reset("))
            {
                continue;
            }

            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1022",
                message: "Raw 'new' detected. Prefer std::make_unique or std::make_shared for automatic memory management.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: "new",
                helpUri: BP1022HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
        }
    }

    // ── BP1023: Heavy macro usage (C++) ───────────────────────────────────────

    public static IEnumerable<JObject> FindMacroOveruse(string file, string content)
    {
        MatchCollection matches = PreprocessorDefinePattern().Matches(content);
        if (matches.Count >= MacroOveruseThreshold)
        {
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1023",
                message: $"File has {matches.Count} #define macros (threshold: {MacroOveruseThreshold}). Prefer constexpr, inline functions, or templates.",
                file: file,
                line: GetLineNumber(content, matches[0].Index),
                symbol: "#define",
                helpUri: BP1023HelpUri);
        }
    }

    // ── BP1024: Deep inheritance (C++) ────────────────────────────────────────

    public static IEnumerable<JObject> FindDeepInheritance(string file, string content)
    {
        int findingCount = 0;
#if NET5_0_OR_GREATER
        foreach (Match match in CppClassInheritancePattern().Matches(content))
#else
        foreach (Match match in new Regex(@"^[ \t]*(?:class|struct)\s+(\w+)\s*:\s*(.+?)(?:\{|$)", RegexOptions.Compiled | RegexOptions.Multiline).Matches(content))
#endif
        {
            string className = match.Groups[1].Value;
            string[] bases = match.Groups[2].Value.Split(',');
            if (bases.Length >= DeepNestingThreshold)
            {
                yield return DiagnosticRowFactory.CreateBestPracticeRow(
                    code: "BP1024",
                    message: $"Class '{className}' inherits from {bases.Length} bases. Deep/wide inheritance is hard to maintain; prefer composition.",
                    file: file,
                    line: GetLineNumber(content, match.Index),
                    symbol: className,
                    helpUri: BP1024HelpUri);
                findingCount++;
                if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
            }
        }
    }

    // ── BP1025: Missing const correctness (C++) ───────────────────────────────

    public static IEnumerable<JObject> FindMissingConst(string file, string content)
    {
        int findingCount = 0;
#if NET5_0_OR_GREATER
        foreach (Match match in CppPassByValuePattern().Matches(content))
#else
        foreach (Match match in new Regex(@"\b(?:std::(?:string|vector|map|unordered_map|set|list|deque|array)|string|vector|map)\s+(\w+)\s*[,)]", RegexOptions.Compiled).Matches(content))
#endif
        {
            if (IsInsideCppLineComment(content, match.Index)) continue;
            if (IsInsideCppStringLiteral(content, match.Index)) continue;
            if (!IsInsideFunctionParams(content, match.Index)) continue;
            string paramName = match.Groups[1].Value;
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1025",
                message: $"Parameter '{paramName}' is passed by value. Use 'const &' to avoid unnecessary copies.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: paramName,
                helpUri: BP1025HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
        }
    }

    // ── BP1026: Python == True/False ──────────────────────────────────────────

    public static IEnumerable<JObject> FindBooleanComparison(string file, string content)
    {
        int findingCount = 0;
        foreach (Match match in PythonBoolComparePattern().Matches(content))
        {
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1026",
                message: $"'{match.Value.Trim()}' is redundant. Use truthiness directly: 'if x:' not 'if x == True:', 'if not x:' not 'if x == False:'.",
                file: file,
                line: GetLineNumber(content, match.Index),
                symbol: match.Value.Trim(),
                helpUri: BP1026HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile) { yield break; }
        }
    }

    // ── BP1028: Unnecessary or redundant comments (C#) ───────────────────────

    public static IEnumerable<JObject> FindUnnecessaryComments(string file, string content)
    {
        MatchCollection commentMatches = CSharpCommentPattern().Matches(content);
        foreach (Match match in commentMatches)
        {
            string raw = match.Value.Trim();

            // Skip XML documentation comments (///) entirely — they are API documentation,
            // not inline comments, and follow conventions like "Gets the X." that are never low-value.
            if (raw.StartsWith("///", StringComparison.Ordinal))
                continue;

            string commentText = raw.StartsWith("//", StringComparison.Ordinal)
#if NET5_0_OR_GREATER
                ? raw[2..].Trim()
#else
                ? raw.Substring(2).Trim()
#endif
                : raw.TrimStart('/').TrimStart('*').TrimEnd('*').TrimEnd('/').Trim();

            if (string.IsNullOrWhiteSpace(commentText))
            {
                continue;
            }

            if (IsTrulyUsefulComment(commentText))
            {
                continue;
            }

            if (IsUnnecessaryComment(commentText))
            {
                yield return DiagnosticRowFactory.CreateBestPracticeRow(
                    code: "BP1028",
                    message: $"Unnecessary comment detected: \"{BestPracticeAnalyzerHelpers.Truncate(commentText, 72)}\". " +
                             "This comment restates what the code already clearly expresses or adds no meaningful value. " +
                             "Remove it to reduce visual noise.",
                    file: file,
                    line: GetLineNumber(content, match.Index),
                    symbol: "Redundant comment",
                    helpUri: BP1028HelpUri);
            }
        }
    }

    // ── BP1029: Namespace/folder structure issues (C#) ────────────────────────

    public static IEnumerable<JObject> FindNamespaceFolderStructureIssues(string file, string content)
    {
        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file);
        string normalizedPath = file.Replace('/', '\\');
        if (string.Equals(fileNameWithoutExtension, "IsExternalInit", StringComparison.Ordinal)
            && normalizedPath.EndsWith(@"\System\Runtime\CompilerServices\IsExternalInit.cs", StringComparison.OrdinalIgnoreCase))
        {
            yield break;
        }

#if NET5_0_OR_GREATER
        if (fileNameWithoutExtension.Contains('.'))
#else
        if (fileNameWithoutExtension.Contains("."))
#endif
        {
            string suggestedPath = fileNameWithoutExtension.Replace('.', '\\') + ".cs";
            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1029",
                message: $"File '{Path.GetFileName(file)}' uses dotted naming that hides the owning folder structure. Move it into folders such as '{suggestedPath}' instead of keeping partial groups in the filename.",
                file: file,
                line: 1,
                symbol: fileNameWithoutExtension,
                helpUri: BP1029HelpUri);
            yield break;
        }

        Match namespaceMatch = CSharpNamespacePattern().Match(content);
        if (!namespaceMatch.Success)
        {
            yield break;
        }

        string declaredNamespace = namespaceMatch.Groups["name"].Value;
        if (IsExternalInitShim(file, fileNameWithoutExtension, declaredNamespace))
        {
            yield break;
        }

        string directoryPath = Path.GetDirectoryName(file) ?? string.Empty;
        string[] relativeSegments = BestPracticeAnalyzerHelpers.GetRelativeDirectorySegments(directoryPath);
        if (relativeSegments.Length == 0)
        {
            yield break;
        }

        string[] partialTypeNames = BestPracticeAnalyzerHelpers.GetDeclaredPartialTypeNames(content);
        string[] structuralSegments = BestPracticeAnalyzerHelpers.TrimTypeGroupSegments(relativeSegments, partialTypeNames);
        string[] namespaceSegments = declaredNamespace.Split('.');
        if (BestPracticeAnalyzerHelpers.NamespaceMatchesFolderStructure(structuralSegments, namespaceSegments))
        {
            yield break;
        }

        string actualFolder = string.Join("\\", structuralSegments);
        yield return DiagnosticRowFactory.CreateBestPracticeRow(
            code: "BP1029",
            message: $"Namespace '{declaredNamespace}' does not match the folder structure '{actualFolder}'. Align namespace segments with folders, and keep extra partial organization under an owning-type folder instead of dotted filenames.",
            file: file,
            line: GetLineNumber(content, namespaceMatch.Index),
            symbol: declaredNamespace,
            helpUri: BP1029HelpUri);
    }

    private static bool IsExternalInitShim(string file, string fileNameWithoutExtension, string declaredNamespace)
    {
        if (!string.Equals(fileNameWithoutExtension, "IsExternalInit", StringComparison.Ordinal)
            || !string.Equals(declaredNamespace, "System.Runtime.CompilerServices", StringComparison.Ordinal))
        {
            return false;
        }

        string normalizedPath = file.Replace('/', '\\');
        return normalizedPath.EndsWith(@"\System\Runtime\CompilerServices\IsExternalInit.cs", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsTopLevelProgramCatch(string file, string content, int matchIndex)
    {
        if (!string.Equals(Path.GetFileName(file), "Program.cs", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        int lineNumber = GetLineNumber(content, matchIndex);
#if NET5_0_OR_GREATER
        return lineNumber <= 100 && content.Contains("static void Main", StringComparison.Ordinal);
#else
        return lineNumber <= 100 && content.IndexOf("static void Main", StringComparison.Ordinal) >= 0;
#endif
    }

    // Returns true when matchIndex falls inside a C/C++ line comment (// ...).
    // Used by BP1025 to avoid flagging identifiers that appear in comment text.
    private static bool IsInsideCppLineComment(string content, int matchIndex)
    {
        int lineStart = matchIndex > 0 ? content.LastIndexOf('\n', matchIndex - 1) + 1 : 0;
        return content.IndexOf("//", lineStart, matchIndex - lineStart, StringComparison.Ordinal) >= 0;
    }

    // Returns true when matchIndex falls inside a double-quoted C/C++ string literal
    // on its line. Used by BP1025 to avoid flagging identifiers that appear in string
    // text (e.g. the word 'vector' inside the message "... is a string vector ..."
    // passed to throw_exception), which would otherwise look like a by-value parameter.
    private static bool IsInsideCppStringLiteral(string content, int matchIndex)
    {
        int lineStart = matchIndex > 0 ? content.LastIndexOf('\n', matchIndex - 1) + 1 : 0;
        bool inString = false;
        for (int i = lineStart; i < matchIndex; i++)
        {
            char ch = content[i];
            if (ch == '"') { inString = !inString; continue; }
            if (inString && ch == '\\') i++; // skip escaped character inside the string
        }
        return inString;
    }

    // Returns true when matchIndex is inside a function parameter list.
    // Walk backwards: an unmatched '(' before any statement boundary (';', '{', '}')  
    // means we are in params; hitting a boundary first means it is a local declaration.
    private static bool IsInsideFunctionParams(string content, int matchIndex)
    {
        int depth = 0;
        for (int i = matchIndex - 1; i >= 0; i--)
        {
            char ch = content[i];
            if      (ch == ')')  { depth++; }
            else if (ch == '(')  { if (depth == 0) return true; depth--; }
            else if (ch == ';' || ch == '{' || ch == '}') { return false; }
        }
        return false;
    }

    // ── BP1027: Property bag class (C#) ───────────────────────────────────────

    public static IEnumerable<JObject> FindPropertyBagClass(string file, string content)
    {
        MatchCollection classMatches = CSharpClassDeclPattern().Matches(content);
        int findingCount = 0;
        string[] lines = content.Split('\n');

        foreach (Match classMatch in classMatches)
        {
            string className = classMatch.Groups[1].Value;
            // Partial classes spread their members across multiple files; a file
            // containing only properties is normal and not a property-bag smell.
#if NET5_0_OR_GREATER
            if (classMatch.Value.Contains("partial", StringComparison.Ordinal))
#else
            if (classMatch.Value.IndexOf("partial", StringComparison.Ordinal) >= 0)
#endif
                continue;
            int classStartLine = BestPracticeAnalyzerHelpers.GetLineNumber(content, classMatch.Index);
            string classBody = BestPracticeAnalyzerHelpers.ExtractBracedBlock(lines, classStartLine - 1);
            int propertyCount = CSharpAutoPropertyPattern().Matches(classBody).Count;
            int behaviorCount = CSharpMethodSignaturePattern().Matches(classBody).Count;

            if (propertyCount < PropertyBagPropertyThreshold || behaviorCount > PropertyBagBehaviorThreshold)
            {
                continue;
            }

            yield return DiagnosticRowFactory.CreateBestPracticeRow(
                code: "BP1027",
                message: $"Class '{className}' has {propertyCount} auto-properties and only {behaviorCount} behavioral methods. Move shared state behind a focused service or model instead of growing an accessor-only class.",
                file: file,
                line: classStartLine,
                symbol: className,
                helpUri: BP1027HelpUri);
            findingCount++;
            if (findingCount >= MaxSuppressionFindingsPerFile)
            {
                yield break;
            }
        }
    }

#if NET5_0_OR_GREATER
    [System.Text.RegularExpressions.GeneratedRegex(@"\b(?:for|foreach|while)\s*\(")]
    private static partial Regex LoopKeywordPattern();

    [System.Text.RegularExpressions.GeneratedRegex(@"^[ \t]*(?:class|struct)\s+(\w+)\s*:\s*(.+?)(?:\{|$)", RegexOptions.Multiline)]
    private static partial Regex CppClassInheritancePattern();

    [System.Text.RegularExpressions.GeneratedRegex(@"\b(?:std::(?:string|vector|map|unordered_map|set|list|deque|array)|string|vector|map)\s+(\w+)\s*[,)]")]
    private static partial Regex CppPassByValuePattern();
#endif

}
