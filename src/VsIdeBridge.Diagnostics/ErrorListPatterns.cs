using System.Text.RegularExpressions;

namespace VsIdeBridge.Diagnostics;

internal static partial class ErrorListPatterns
{
    [GeneratedRegex(@"\b(?:LINK|LNK|MSB|VCR|E|C)\d+\b|\blnt-[a-z0-9-]+\b|\bInt-make\b", RegexOptions.IgnoreCase)]
    public static partial Regex ExplicitCodePattern();

    [GeneratedRegex(@"^(?<file>[A-Za-z]:\\.*?|\S.*?)(?:\((?<line>\d+)(?:,(?<column>\d+))?\))?\s*:\s*(?<severity>warning|error)\s+(?<code>[A-Za-z]+[A-Za-z0-9-]*)\s*:\s*(?<message>.+?)(?:\s+\[(?<project>[^\]]+)\])?$", RegexOptions.IgnoreCase)]
    public static partial Regex MsBuildDiagnosticPattern();

    [GeneratedRegex(@"^(?<project>.+?)\s*>\s*(?<file>[A-Za-z]:\\.*?|\S.*?)(?:\((?<line>\d+)(?:,(?<column>\d+))?\))?\s*:\s*(?<severity>warning|error)\s+(?<code>[A-Za-z]+[A-Za-z0-9-]*)\s*:\s*(?<message>.+)$", RegexOptions.IgnoreCase)]
    public static partial Regex StructuredOutputPattern();

    [GeneratedRegex("\"([^\"\\r\\n]{8,})\"")]
    public static partial Regex StringLiteralPattern();

    [GeneratedRegex(@"\bconst\s+string\s+\w+\s*=")]
    public static partial Regex ConstStringDeclPattern();

    [GeneratedRegex(@"(?<![A-Za-z0-9_\.])(?<value>-?\d+(?:\.\d+)?)\b")]
    public static partial Regex NumberLiteralPattern();

    [GeneratedRegex(@"\(int\)\s*Math\s*\.\s*(?<op>Floor|Truncate)\s*\(")]
    public static partial Regex SuspiciousRoundDownPattern();

    [GeneratedRegex(@"catch\s*(?:\([^)]*\))?\s*\{")]
    public static partial Regex CatchBlockPattern();

    [GeneratedRegex(@"\basync\s+void\s+(\w+)")]
    public static partial Regex AsyncVoidPattern();

    [GeneratedRegex(@"\bdelete\s*(\[\])?\s")]
    public static partial Regex RawDeletePattern();

    [GeneratedRegex(@"^\s*using\s+namespace\s+([\w:]+)\s*;", RegexOptions.Multiline)]
    public static partial Regex UsingNamespacePattern();

    // The trailing lookahead rejects declaration contexts such as "size_t vertexCount(size_t) const",
    // where "(size_t) c" is an unnamed parameter followed by a qualifier, not a cast.
    [GeneratedRegex(@"\((?:const\s+)?(?:unsigned\s+)?(?:int|long|short|char|float|double|size_t|uint\d+_t|int\d+_t|void)\s*\*?\)\s*(?!(?:const|noexcept|override|final)\b)[a-zA-Z_\(]")]
    public static partial Regex CStyleCastPattern();

    [GeneratedRegex(@"^\s*except\s*:", RegexOptions.Multiline)]
    public static partial Regex BareExceptPattern();

    [GeneratedRegex(@"\bdef\s+(\w+)\s*\([^)]*=\s*(?:\[\s*\]|\{\s*\}|set\s*\(\s*\))")]
    public static partial Regex MutableDefaultArgPattern();

    [GeneratedRegex(@"^\s*from\s+(\S+)\s+import\s+\*", RegexOptions.Multiline)]
    public static partial Regex ImportStarPattern();

    [GeneratedRegex(@"//.*?$|/\*.*?\*/", RegexOptions.Multiline)]
    public static partial Regex CSharpCommentPattern();

    [GeneratedRegex(
        @"^[ \t]*(?:(?:public|private|protected|internal|static|virtual|override|abstract|sealed|async|new|partial|readonly)\s+)*(?" +
        @"!return\b|if\b|else\b|while\b|for\b|foreach\b|switch\b|catch\b|using\b|lock\b|yield\b|class\b|struct\b|interface\b|enum\b" +
        @"|record\b|namespace\b|delegate\b)[\w<>\[\],\?]+\s+(\w+)\s*(?:<[^>]+>)?\s*\(",
        RegexOptions.Multiline)]
    public static partial Regex CSharpMethodSignaturePattern();

    [GeneratedRegex(@"^[ \t]*def\s+(\w+)\s*\(", RegexOptions.Multiline)]
    public static partial Regex PythonDefPattern();

    [GeneratedRegex(@"^[ \t]*(?:(?:static|virtual|inline|explicit|constexpr|const|unsigned|signed|volatile|extern|friend|template\s*<[^>]*>)\s+)*[\w:*&<>,\s]+\s+(\w+)\s*\([^;]*$", RegexOptions.Multiline)]
    public static partial Regex CppFunctionPattern();

    [GeneratedRegex(@"\b(?:var|int|string|bool|double|float|long|object|dynamic)\s+(?<name>tmp|temp|data|val|res|ret|result|process|handle|flag|item|stuff|thing|manager|helper|util|misc)\b", RegexOptions.IgnoreCase)]
    public static partial Regex PoorCSharpNamingPattern();

    [GeneratedRegex(@"(?:(?:var|int|string|bool|double|float|long|short|byte|char|object|decimal)\s+(?<name>[a-zA-Z])\s*[=;,)])")]
    public static partial Regex SingleLetterVarPattern();

    [GeneratedRegex(@"(?<!\.)\bvar\s+(?<name>@?[A-Za-z_][A-Za-z0-9_]*)\s*(?:=|;|,)")]
    public static partial Regex ImplicitVarPattern();

    [GeneratedRegex(@"catch\s*\(\s*(?:System\.)?Exception(?:\s+[A-Za-z_][A-Za-z0-9_]*)?\s*\)(?!\s*when\b)")]
    public static partial Regex BroadCatchPattern();

    [GeneratedRegex(@"^[ \t]*#pragma\s+warning(?:\s+disable\b.*|\s*\(\s*disable\s*:\s*[^)]*\))$", RegexOptions.Multiline | RegexOptions.IgnoreCase)]
    public static partial Regex PragmaWarningDisablePattern();

    [GeneratedRegex(@"(?im)^[ \t]*(?:#\s*define\b.*__pragma\s*\(\s*warning\s*\(\s*disable\b.*|.*_Pragma\s*\(\s*""(?:[^""\\]|\\.)*diagnostic\s+ignored.*)$")]
    public static partial Regex NativeDiagnosticSuppressionPattern();

    [GeneratedRegex(@"(?im)^[ \t]*dotnet_diagnostic\.(?<code>[A-Za-z0-9_.-]+)\.severity\s*=\s*(?<severity>none|silent)\b.*$")]
    public static partial Regex EditorConfigDiagnosticSuppressionPattern();

    [GeneratedRegex(@"(?im)<NoWarn>\s*(?<codes>[^<]+?)\s*</NoWarn>")]
    public static partial Regex NoWarnElementPattern();

    [GeneratedRegex(@"(?im)\bNoWarn\s*=\s*""(?<codes>[^""]+?)""")]
    public static partial Regex NoWarnAttributePattern();

    [GeneratedRegex(@"(?is)\[\s*(?:assembly\s*:\s*)?SuppressMessage\s*\((?<args>.*?)\)\s*\]")]
    public static partial Regex SuppressMessagePattern();

    [GeneratedRegex(@"(?im)<Rule\b[^>]*\bId\s*=\s*""(?<code>[^""]+)""[^>]*\bAction\s*=\s*""(?<action>None|Hidden)""[^>]*/?>")]
    public static partial Regex RuleSetSuppressionPattern();

    [GeneratedRegex(@"(?im)^[ \t]*(?://+|#|'+|/\*+|\*+|\(\*)\s*(?<marker>TODO|FIXME|XXX|HACK|TBD|BUGBUG)\b(?<text>.*)$")]
    public static partial Regex TodoCommentPattern();

    [GeneratedRegex(@"\bSystem\.(?<type>String|Int16|Int32|Int64|UInt16|UInt32|UInt64|Boolean|Object|Decimal|Double|Single|Byte|SByte|Char)\b")]
    public static partial Regex FrameworkTypePattern();

    [GeneratedRegex(@"(?im)^\s*Option\s+Strict\s+On\b")]
    public static partial Regex VbOptionStrictOnPattern();

    [GeneratedRegex(@"(?im)^\s*Option\s+Strict\s+Off\b")]
    public static partial Regex VbOptionStrictOffPattern();

    [GeneratedRegex(@"(?m)_\s*(?:'.*)?$")]
    public static partial Regex VbLineContinuationPattern();

    [GeneratedRegex(@"\bmutable\b")]
    public static partial Regex FSharpMutablePattern();

    [GeneratedRegex(@"\(\*.*?\*\)", RegexOptions.Singleline)]
    public static partial Regex FSharpBlockCommentPattern();

    [GeneratedRegex(@"(?:==|!=)\s*None\b|\bNone\s*(?:==|!=)")]
    public static partial Regex PythonNoneComparisonPattern();

    [GeneratedRegex(@"(?im)(?:^|[|;]\s*)(?<alias>ls|dir|gc|cat|echo|sleep|cp|mv|rm|del|%|\?)\b")]
    public static partial Regex PowerShellAliasPattern();

    [GeneratedRegex(@"^[ \t]*(?<name>tmp|temp|data|val|res|ret|result|process|handle|flag|stuff|thing)\s*=", RegexOptions.Multiline | RegexOptions.IgnoreCase)]
    public static partial Regex PythonPoorNamingPattern();

    [GeneratedRegex(@"^[ \t]*(?<name>[a-zA-Z])\s*=\s*(?!.*\bfor\b)", RegexOptions.Multiline)]
    public static partial Regex PythonSingleLetterAssignPattern();

    [GeneratedRegex(@"^\s*//\s*(?:(?:public|private|protected|internal|static|var|if|else|for|foreach|while|return|throw|try|catch|class|using|namespace|void|int|string|bool)\b|\w+\s*\(.*\)\s*[;{]|\w+\s*=\s*)")]
    public static partial Regex CSharpCommentedCodePattern();

    [GeneratedRegex(@"^\s*#\s*(?:(?:def|class|if|else|elif|for|while|return|import|from|try|except|raise|with|yield)\b|\w+\s*\(.*\)\s*$|\w+\s*=\s*)")]
    public static partial Regex PythonCommentedCodePattern();

    [GeneratedRegex(@"^\s*//\s*(?:(?:class|struct|if|else|for|while|return|throw|try|catch|namespace|void|int|auto|const|static|virtual|template)\b|\w+\s*\(.*\)\s*[;{]|\w+\s*=\s*)")]
    public static partial Regex CppCommentedCodePattern();

    [GeneratedRegex(@"^\t", RegexOptions.Multiline)]
    public static partial Regex TabIndentedLinePattern();

    [GeneratedRegex(@"^ {2,}", RegexOptions.Multiline)]
    public static partial Regex SpaceIndentedLinePattern();

    [GeneratedRegex(@"^[ \t]*(?:(?:public|private|protected|internal|static|sealed|abstract|partial)\s+)*class\s+(\w+)", RegexOptions.Multiline)]
    public static partial Regex CSharpClassDeclPattern();

    [GeneratedRegex(@"^[ \t]*(?:(?:public|private|protected|internal|static|readonly|volatile|const)\s+)+[\w<>\[\],\?\s]+\s+_?\w+\s*[=;]", RegexOptions.Multiline)]
    public static partial Regex CSharpFieldDeclPattern();

    [GeneratedRegex(@"(?:for\s*\(|foreach\s*\(|while\s*\()[^{]*\{[^}]*DateTime\s*\.\s*(?:Now|UtcNow)", RegexOptions.Singleline)]
    public static partial Regex DateTimeInLoopPattern();

    [GeneratedRegex(@"DateTime\s*\.\s*(?<prop>Now|UtcNow)")]
    public static partial Regex DateTimeNowSimplePattern();

    [GeneratedRegex(@"\b(?:dynamic|object)\s+\w+\s*[,)]")]
    public static partial Regex DynamicObjectParamPattern();

    [GeneratedRegex(@"(?<!(?:unique_ptr|shared_ptr|make_unique|make_shared|reset|emplace)\s*(?:<[^>]*>\s*)?\()\bnew\s+\w+")]
    public static partial Regex RawNewPattern();

    [GeneratedRegex(@"^\s*#\s*define\s+(\w+)", RegexOptions.Multiline)]
    public static partial Regex PreprocessorDefinePattern();

    [GeneratedRegex(@"^\s*(?:virtual\s+)?(?:[\w:*&<>,\s]+)\s+(\w+)\s*\([^)]*\)\s*(?:override\s*)?(?=\s*\{)", RegexOptions.Multiline)]
    public static partial Regex CppNonConstMethodPattern();

    [GeneratedRegex(@"(?:==\s*True|==\s*False|is\s+True|is\s+False)\b")]
    public static partial Regex PythonBoolComparePattern();

    [GeneratedRegex(@"^[ \t]*(?:(?:public|private|protected|internal|static|virtual|override|required|sealed|abstract)\s+)+[\w<>\[\],\?\s]+\s+\w+\s*\{\s*get;\s*(?:set;|init;)?\s*\}", RegexOptions.Multiline)]
    public static partial Regex CSharpAutoPropertyPattern();

    [GeneratedRegex(@"^\s*namespace\s+(?<name>[A-Za-z_][A-Za-z0-9_.]*)\s*(?:;|\{)", RegexOptions.Multiline)]
    public static partial Regex CSharpNamespacePattern();

    [GeneratedRegex(@"\bpartial\s+(?:class|struct|interface|record)\s+(?<name>@?[A-Za-z_][A-Za-z0-9_]*)\b")]
    public static partial Regex PartialTypeDeclarationPattern();
}
