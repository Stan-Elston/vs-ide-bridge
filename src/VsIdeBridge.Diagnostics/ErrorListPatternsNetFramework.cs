using System.Text.RegularExpressions;

namespace VsIdeBridge.Diagnostics;

internal static partial class ErrorListPatterns
{
    private static readonly Regex _explicitCodePattern = new(
        @"\b(?:LINK|LNK|MSB|VCR|E|C)\d+\b|\blnt-[a-z0-9-]+\b|\bInt-make\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    public static Regex ExplicitCodePattern() { return _explicitCodePattern; }

    private static readonly Regex _msBuildDiagnosticPattern = new(
        @"^(?<file>[A-Za-z]:\\.*?|\S.*?)(?:\((?<line>\d+)(?:,(?<column>\d+))?\))?\s*:\s*(?<severity>warning|error)\s+(?<code>[A-Za-z]+[A-Za-z0-9-]*)\s*:\s*(?<message>.+?)(?:\s+\[(?<project>[^\]]+)\])?$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    public static Regex MsBuildDiagnosticPattern() { return _msBuildDiagnosticPattern; }

    private static readonly Regex _structuredOutputPattern = new(
        @"^(?<project>.+?)\s*>\s*(?<file>[A-Za-z]:\\.*?|\S.*?)(?:\((?<line>\d+)(?:,(?<column>\d+))?\))?\s*:\s*(?<severity>warning|error)\s+(?<code>[A-Za-z]+[A-Za-z0-9-]*)\s*:\s*(?<message>.+)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    public static Regex StructuredOutputPattern() { return _structuredOutputPattern; }

    private static readonly Regex _stringLiteralPattern = new("\"([^\"\\r\\n]{8,})\"", RegexOptions.Compiled);
    public static Regex StringLiteralPattern() { return _stringLiteralPattern; }

    private static readonly Regex _constStringDeclPattern = new(@"\bconst\s+string\s+\w+\s*=", RegexOptions.Compiled);
    public static Regex ConstStringDeclPattern() { return _constStringDeclPattern; }

    private static readonly Regex _numberLiteralPattern = new(@"(?<![A-Za-z0-9_\.])(?<value>-?\d+(?:\.\d+)?)\b", RegexOptions.Compiled);
    public static Regex NumberLiteralPattern() { return _numberLiteralPattern; }

    private static readonly Regex _suspiciousRoundDownPattern = new(@"\(int\)\s*Math\s*\.\s*(?<op>Floor|Truncate)\s*\(", RegexOptions.Compiled);
    public static Regex SuspiciousRoundDownPattern() { return _suspiciousRoundDownPattern; }

    private static readonly Regex _catchBlockPattern = new(@"catch\s*(?:\([^)]*\))?\s*\{", RegexOptions.Compiled);
    public static Regex CatchBlockPattern() { return _catchBlockPattern; }

    private static readonly Regex _asyncVoidPattern = new(@"\basync\s+void\s+(\w+)", RegexOptions.Compiled);
    public static Regex AsyncVoidPattern() { return _asyncVoidPattern; }

    private static readonly Regex _rawDeletePattern = new(@"\bdelete\s*(\[\])?\s", RegexOptions.Compiled);
    public static Regex RawDeletePattern() { return _rawDeletePattern; }

    private static readonly Regex _usingNamespacePattern = new(@"^\s*using\s+namespace\s+([\w:]+)\s*;", RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex UsingNamespacePattern() { return _usingNamespacePattern; }

    // The trailing lookahead rejects declaration contexts such as "size_t vertexCount(size_t) const",
    // where "(size_t) c" is an unnamed parameter followed by a qualifier, not a cast.
    private static readonly Regex _cStyleCastPattern = new(@"\((?:const\s+)?(?:unsigned\s+)?(?:int|long|short|char|float|double|size_t|uint\d+_t|int\d+_t|void)\s*\*?\)\s*(?!(?:const|noexcept|override|final)\b)[a-zA-Z_\(]", RegexOptions.Compiled);
    public static Regex CStyleCastPattern() { return _cStyleCastPattern; }

    private static readonly Regex _bareExceptPattern = new(@"^\s*except\s*:", RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex BareExceptPattern() { return _bareExceptPattern; }

    private static readonly Regex _mutableDefaultArgPattern = new(@"\bdef\s+(\w+)\s*\([^)]*=\s*(?:\[\s*\]|\{\s*\}|set\s*\(\s*\))", RegexOptions.Compiled);
    public static Regex MutableDefaultArgPattern() { return _mutableDefaultArgPattern; }

    private static readonly Regex _importStarPattern = new(@"^\s*from\s+(\S+)\s+import\s+\*", RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex ImportStarPattern() { return _importStarPattern; }

    private static readonly Regex _cSharpCommentPattern = new(@"//.*?$|/\*.*?\*/", RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex CSharpCommentPattern() { return _cSharpCommentPattern; }

    private static readonly Regex _cSharpMethodSignaturePattern = new(
        @"^[ \t]*(?:(?:public|private|protected|internal|static|virtual|override|abstract|sealed|async|new|partial|readonly)\s+)*(?" +
        @"!return\b|if\b|else\b|while\b|for\b|foreach\b|switch\b|catch\b|using\b|lock\b|yield\b|class\b|struct\b|interface\b|enum\b" +
        @"|record\b|namespace\b|delegate\b)[\w<>\[\],\?]+\s+(\w+)\s*(?:<[^>]+>)?\s*\(",
        RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex CSharpMethodSignaturePattern() { return _cSharpMethodSignaturePattern; }

    private static readonly Regex _pythonDefPattern = new(@"^[ \t]*def\s+(\w+)\s*\(", RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex PythonDefPattern() { return _pythonDefPattern; }

    private static readonly Regex _cppFunctionPattern = new(
        @"^[ \t]*(?:(?:static|virtual|inline|explicit|constexpr|const|unsigned|signed|volatile|extern|friend|template\s*<[^>]*>)\s+)*[\w:*&<>,\s]+\s+(\w+)\s*\([^;]*$",
        RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex CppFunctionPattern() { return _cppFunctionPattern; }

    private static readonly Regex _poorCSharpNamingPattern = new(
        @"\b(?:var|int|string|bool|double|float|long|object|dynamic)\s+(?<name>tmp|temp|data|val|res|ret|result|process|handle|flag|item|stuff|thing|manager|helper|util|misc)\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    public static Regex PoorCSharpNamingPattern() { return _poorCSharpNamingPattern; }

    private static readonly Regex _singleLetterVarPattern = new(
        @"(?:(?:var|int|string|bool|double|float|long|short|byte|char|object|decimal)\s+(?<name>[a-zA-Z])\s*[=;,)])",
        RegexOptions.Compiled);
    public static Regex SingleLetterVarPattern() { return _singleLetterVarPattern; }

    private static readonly Regex _implicitVarPattern = new(@"(?<!\.)\bvar\s+(?<name>@?[A-Za-z_][A-Za-z0-9_]*)\s*(?:=|;|,)", RegexOptions.Compiled);
    public static Regex ImplicitVarPattern() { return _implicitVarPattern; }

    private static readonly Regex _broadCatchPattern = new(@"catch\s*\(\s*(?:System\.)?Exception(?:\s+[A-Za-z_][A-Za-z0-9_]*)?\s*\)(?!\s*when\b)", RegexOptions.Compiled);
    public static Regex BroadCatchPattern() { return _broadCatchPattern; }

    private static readonly Regex _pragmaWarningDisablePattern = new(@"^[ \t]*#pragma\s+warning(?:\s+disable\b.*|\s*\(\s*disable\s*:\s*[^)]*\))$", RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
    public static Regex PragmaWarningDisablePattern() { return _pragmaWarningDisablePattern; }

    private static readonly Regex _nativeDiagnosticSuppressionPattern = new(@"(?im)^[ \t]*(?:#\s*define\b.*__pragma\s*\(\s*warning\s*\(\s*disable\b.*|.*_Pragma\s*\(\s*""(?:[^""\\]|\\.)*diagnostic\s+ignored.*)$", RegexOptions.Compiled);
    public static Regex NativeDiagnosticSuppressionPattern() { return _nativeDiagnosticSuppressionPattern; }

    private static readonly Regex _editorConfigDiagnosticSuppressionPattern = new(@"(?im)^[ \t]*dotnet_diagnostic\.(?<code>[A-Za-z0-9_.-]+)\.severity\s*=\s*(?<severity>none|silent)\b.*$", RegexOptions.Compiled);
    public static Regex EditorConfigDiagnosticSuppressionPattern() { return _editorConfigDiagnosticSuppressionPattern; }

    private static readonly Regex _noWarnElementPattern = new(@"(?im)<NoWarn>\s*(?<codes>[^<]+?)\s*</NoWarn>", RegexOptions.Compiled);
    public static Regex NoWarnElementPattern() { return _noWarnElementPattern; }

    private static readonly Regex _noWarnAttributePattern = new(@"(?im)\bNoWarn\s*=\s*""(?<codes>[^""]+?)""", RegexOptions.Compiled);
    public static Regex NoWarnAttributePattern() { return _noWarnAttributePattern; }

    private static readonly Regex _suppressMessagePattern = new(@"(?is)\[\s*(?:assembly\s*:\s*)?SuppressMessage\s*\((?<args>.*?)\)\s*\]", RegexOptions.Compiled);
    public static Regex SuppressMessagePattern() { return _suppressMessagePattern; }

    private static readonly Regex _ruleSetSuppressionPattern = new(@"(?im)<Rule\b[^>]*\bId\s*=\s*""(?<code>[^""]+)""[^>]*\bAction\s*=\s*""(?<action>None|Hidden)""[^>]*/?>", RegexOptions.Compiled);
    public static Regex RuleSetSuppressionPattern() { return _ruleSetSuppressionPattern; }

    private static readonly Regex _todoCommentPattern = new(@"(?im)^[ \t]*(?://+|#|'+|/\*+|\*+|\(\*)\s*(?<marker>TODO|FIXME|XXX|HACK|TBD|BUGBUG)\b(?<text>.*)$", RegexOptions.Compiled);
    public static Regex TodoCommentPattern() { return _todoCommentPattern; }

    private static readonly Regex _frameworkTypePattern = new(@"\bSystem\.(?<type>String|Int16|Int32|Int64|UInt16|UInt32|UInt64|Boolean|Object|Decimal|Double|Single|Byte|SByte|Char)\b", RegexOptions.Compiled);
    public static Regex FrameworkTypePattern() { return _frameworkTypePattern; }

    private static readonly Regex _vbOptionStrictOnPattern = new(@"(?im)^\s*Option\s+Strict\s+On\b", RegexOptions.Compiled);
    public static Regex VbOptionStrictOnPattern() { return _vbOptionStrictOnPattern; }

    private static readonly Regex _vbOptionStrictOffPattern = new(@"(?im)^\s*Option\s+Strict\s+Off\b", RegexOptions.Compiled);
    public static Regex VbOptionStrictOffPattern() { return _vbOptionStrictOffPattern; }

    private static readonly Regex _vbLineContinuationPattern = new(@"(?m)_\s*(?:'.*)?$", RegexOptions.Compiled);
    public static Regex VbLineContinuationPattern() { return _vbLineContinuationPattern; }

    private static readonly Regex _fSharpMutablePattern = new(@"\bmutable\b", RegexOptions.Compiled);
    public static Regex FSharpMutablePattern() { return _fSharpMutablePattern; }

    private static readonly Regex _fSharpBlockCommentPattern = new(@"\(\*.*?\*\)", RegexOptions.Compiled | RegexOptions.Singleline);
    public static Regex FSharpBlockCommentPattern() { return _fSharpBlockCommentPattern; }

    private static readonly Regex _pythonNoneComparisonPattern = new(@"(?:==|!=)\s*None\b|\bNone\s*(?:==|!=)", RegexOptions.Compiled);
    public static Regex PythonNoneComparisonPattern() { return _pythonNoneComparisonPattern; }

    private static readonly Regex _powerShellAliasPattern = new(@"(?im)(?:^|[|;]\s*)(?<alias>ls|dir|gc|cat|echo|sleep|cp|mv|rm|del|%|\?)\b", RegexOptions.Compiled);
    public static Regex PowerShellAliasPattern() { return _powerShellAliasPattern; }

    private static readonly Regex _pythonPoorNamingPattern = new(@"^[ \t]*(?<name>tmp|temp|data|val|res|ret|result|process|handle|flag|stuff|thing)\s*=", RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
    public static Regex PythonPoorNamingPattern() { return _pythonPoorNamingPattern; }

    private static readonly Regex _pythonSingleLetterAssignPattern = new(@"^[ \t]*(?<name>[a-zA-Z])\s*=\s*(?!.*\bfor\b)", RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex PythonSingleLetterAssignPattern() { return _pythonSingleLetterAssignPattern; }

    private static readonly Regex _cSharpCommentedCodePattern = new(
        @"^\s*//\s*(?:(?:public|private|protected|internal|static|var|if|else|for|foreach|while|return|throw|try|catch|class|using|" +
        @"namespace|void|int|string|bool)\b|\w+\s*\(.*\)\s*[;{]|\w+\s*=\s*)",
        RegexOptions.Compiled);
    public static Regex CSharpCommentedCodePattern() { return _cSharpCommentedCodePattern; }

    private static readonly Regex _pythonCommentedCodePattern = new(@"^\s*#\s*(?:(?:def|class|if|else|elif|for|while|return|import|from|try|except|raise|with|yield)\b|\w+\s*\(.*\)\s*$|\w+\s*=\s*)", RegexOptions.Compiled);
    public static Regex PythonCommentedCodePattern() { return _pythonCommentedCodePattern; }

    private static readonly Regex _cppCommentedCodePattern = new(@"^\s*//\s*(?:(?:class|struct|if|else|for|while|return|throw|try|catch|namespace|void|int|auto|const|static|virtual|template)\b|\w+\s*\(.*\)\s*[;{]|\w+\s*=\s*)", RegexOptions.Compiled);
    public static Regex CppCommentedCodePattern() { return _cppCommentedCodePattern; }

    private static readonly Regex _tabIndentedLinePattern = new(@"^\t", RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex TabIndentedLinePattern() { return _tabIndentedLinePattern; }

    private static readonly Regex _spaceIndentedLinePattern = new(@"^ {2,}", RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex SpaceIndentedLinePattern() { return _spaceIndentedLinePattern; }

    private static readonly Regex _cSharpClassDeclPattern = new(@"^[ \t]*(?:(?:public|private|protected|internal|static|sealed|abstract|partial)\s+)*class\s+(\w+)", RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex CSharpClassDeclPattern() { return _cSharpClassDeclPattern; }

    private static readonly Regex _cSharpFieldDeclPattern = new(@"^[ \t]*(?:(?:public|private|protected|internal|static|readonly|volatile|const)\s+)+[\w<>\[\],\?\s]+\s+_?\w+\s*[=;]", RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex CSharpFieldDeclPattern() { return _cSharpFieldDeclPattern; }

    private static readonly Regex _dateTimeInLoopPattern = new(@"(?:for\s*\(|foreach\s*\(|while\s*\()[^{]*\{[^}]*DateTime\s*\.\s*(?:Now|UtcNow)", RegexOptions.Compiled | RegexOptions.Singleline);
    public static Regex DateTimeInLoopPattern() { return _dateTimeInLoopPattern; }

    private static readonly Regex _dateTimeNowSimplePattern = new(@"DateTime\s*\.\s*(?<prop>Now|UtcNow)", RegexOptions.Compiled);
    public static Regex DateTimeNowSimplePattern() { return _dateTimeNowSimplePattern; }

    private static readonly Regex _dynamicObjectParamPattern = new(@"\b(?:dynamic|object)\s+\w+\s*[,)]", RegexOptions.Compiled);
    public static Regex DynamicObjectParamPattern() { return _dynamicObjectParamPattern; }

    private static readonly Regex _rawNewPattern = new(
        @"(?<!(?:unique_ptr|shared_ptr|make_unique|make_shared|reset|emplace)\s*(?:<[^>]*>\s*)?\()\bnew\s+\w+",
        RegexOptions.Compiled);
    public static Regex RawNewPattern() { return _rawNewPattern; }

    private static readonly Regex _preprocessorDefinePattern = new(@"^\s*#\s*define\s+(\w+)", RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex PreprocessorDefinePattern() { return _preprocessorDefinePattern; }

    private static readonly Regex _cppNonConstMethodPattern = new(@"^\s*(?:virtual\s+)?(?:[\w:*&<>,\s]+)\s+(\w+)\s*\([^)]*\)\s*(?:override\s*)?(?=\s*\{)", RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex CppNonConstMethodPattern() { return _cppNonConstMethodPattern; }

    private static readonly Regex _pythonBoolComparePattern = new(@"(?:==\s*True|==\s*False|is\s+True|is\s+False)\b", RegexOptions.Compiled);
    public static Regex PythonBoolComparePattern() { return _pythonBoolComparePattern; }

    private static readonly Regex _cSharpAutoPropertyPattern = new(
        @"^[ \t]*(?:(?:public|private|protected|internal|static|virtual|override|required|sealed|abstract)\s+)+[\w<>\[\],\?\s]+\s+\w+\s*\{\s*get;\s*(?:set;|init;)?\s*\}",
        RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex CSharpAutoPropertyPattern() { return _cSharpAutoPropertyPattern; }

    private static readonly Regex _cSharpNamespacePattern = new(@"^\s*namespace\s+(?<name>[A-Za-z_][A-Za-z0-9_.]*)\s*(?:;|\{)", RegexOptions.Compiled | RegexOptions.Multiline);
    public static Regex CSharpNamespacePattern() { return _cSharpNamespacePattern; }

    private static readonly Regex _partialTypeDeclarationPattern = new(@"\bpartial\s+(?:class|struct|interface|record)\s+(?<name>@?[A-Za-z_][A-Za-z0-9_]*)\b", RegexOptions.Compiled);
    public static Regex PartialTypeDeclarationPattern() { return _partialTypeDeclarationPattern; }
}
