using System.Linq;
using VsIdeBridge.Diagnostics;
using Xunit;

namespace VsIdeBridgeService.Tests;

public sealed class BestPracticeAnalyzerTests
{
    [Theory]
    [InlineData("sample.cs", "/// version 2\n// version 2\n/* version 2\nversion 2 */\nint value = 7;\n")]
    [InlineData("sample.cpp", "/* version 2\nversion 2\nversion 2\nversion 2 */\nint value = 7;\n")]
    [InlineData("sample.py", "# version 2\n# version 2\n# version 2\n# version 2\nvalue = 7\n")]
    [InlineData("sample.ps1", "# version 2\n# version 2\n# version 2\n# version 2\n$value = 7\n")]
    [InlineData("sample.vb", "' version 2\n' version 2\n' version 2\n' version 2\nDim value = 7\n")]
    [InlineData("sample.fs", "(* version 2\nversion 2\nversion 2\nversion 2 *)\nlet value = 7\n")]
    [InlineData("sample.csproj", "<!-- version 2\nversion 2\nversion 2\nversion 2 -->\n<Project />\n")]
    [InlineData(".editorconfig", "; version 2\n; version 2\n; version 2\n; version 2\nroot = true\n")]
    public void FindMagicNumbersIgnoresCommentLiterals(string file, string content)
    {
        Assert.DoesNotContain(BestPracticeAnalyzer.FindMagicNumbers(file, content),
            row => row["code"]?.ToString() == "BP1002");
    }

    [Fact]
    public void FindMagicNumbersStillReportsRepeatedCodeLiterals()
    {
        string content =
            "int firstCount = 42;\n" +
            "int secondCount = 42;\n" +
            "int thirdCount = 42;\n" +
            "int fourthCount = 42;\n";

        Newtonsoft.Json.Linq.JObject finding = Assert.Single(BestPracticeAnalyzer.FindMagicNumbers("sample.cs", content));
        Assert.Equal("BP1002", finding["code"]?.ToString());
        Assert.Equal("42", finding["symbol"]?.ToString());
    }

    [Fact]
    public void FindRepeatedStringLiteralsIgnoresXmlDocParamNames()
    {
        string content =
            "/// <param name=\"patternString\">First use.</param>\n" +
            "/// <param name=\"patternString\">Second use.</param>\n" +
            "/// <param name=\"patternString\">Third use.</param>\n" +
            "/// <param name=\"patternString\">Fourth use.</param>\n" +
            "/// <param name=\"patternString\">Fifth use.</param>\n" +
            "public static bool Matches(string patternString) => patternString.Length > 0;\n";

        Assert.DoesNotContain(BestPracticeAnalyzer.FindRepeatedStringLiterals("sample.cs", content),
            row => row["code"]?.ToString() == "BP1001");
    }

    [Fact]
    public void FindRepeatedStringLiteralsStillReportsCodeLiterals()
    {
        string content =
            "string first = \"repeat-me\";\n" +
            "string second = \"repeat-me\";\n" +
            "string third = \"repeat-me\";\n" +
            "string fourth = \"repeat-me\";\n" +
            "string fifth = \"repeat-me\";\n";

        Newtonsoft.Json.Linq.JObject finding =
            Assert.Single(BestPracticeAnalyzer.FindRepeatedStringLiterals("sample.cs", content));
        Assert.Equal("BP1001", finding["code"]?.ToString());
        Assert.Equal("repeat-me", finding["symbol"]?.ToString());
    }

    [Theory]
    [InlineData("sample.cpp", "// creates a new node\nint value = 7;\n")]
    [InlineData("sample.cpp", "/* returns a new mesh from two new meshes */\nint value = 7;\n")]
    [InlineData("sample.c", "const char* msg = \"allocate a new buffer\";\n")]
    public void FindRawNewIgnoresCommentAndStringOccurrences(string file, string content)
    {
        Assert.DoesNotContain(BestPracticeAnalyzer.FindRawNew(file, content),
            row => row["code"]?.ToString() == "BP1022");
    }

    [Fact]
    public void FindRawNewStillReportsRealAllocation()
    {
        string content = "Widget* w = new Widget();\n";

        Newtonsoft.Json.Linq.JObject finding =
            Assert.Single(BestPracticeAnalyzer.FindRawNew("sample.cpp", content));
        Assert.Equal("BP1022", finding["code"]?.ToString());
    }

    [Fact]
    public void FindLongLines_FlagsLineOverLimit()
    {
        string longLine = new('x', ErrorListConstants.LineTooLongThreshold + 1);
        string content = $"short line\n{longLine}\nshort again\n";

        Newtonsoft.Json.Linq.JObject finding =
            Assert.Single(BestPracticeAnalyzer.FindLongLines("sample.cs", content));
        Assert.Equal("BP1046", finding["code"]?.ToString());
        Assert.Equal("2", finding["line"]?.ToString());
    }

    [Fact]
    public void FindLongLines_DoesNotFlagExactlyAtLimit()
    {
        string atLimit = new('x', ErrorListConstants.LineTooLongThreshold);
        string content = $"short line\n{atLimit}\nshort again\n";

        Assert.Empty(BestPracticeAnalyzer.FindLongLines("sample.cs", content));
    }

    [Fact]
    public void FindLongLines_SkipsUrlOnlyLines()
    {
        string urlLine = "https://" + new string('a', ErrorListConstants.LineTooLongThreshold);
        string content = $"{urlLine}\n";

        Assert.Empty(BestPracticeAnalyzer.FindLongLines("sample.cs", content));
    }

    [Theory]
    [InlineData("sample.cs")]
    [InlineData("sample.cpp")]
    [InlineData("sample.py")]
    [InlineData("sample.ps1")]
    public void FindLongLines_AppliesToAllLanguages(string fileName)
    {
        string longLine = new('x', ErrorListConstants.LineTooLongThreshold + 5);
        string content = $"{longLine}\n";

        Newtonsoft.Json.Linq.JObject finding =
            Assert.Single(BestPracticeAnalyzer.FindLongLines(fileName, content));
        Assert.Equal("BP1046", finding["code"]?.ToString());
    }

    [Fact]
    public void FindDiagnosticSuppressionsFlagsNativePragmaMacros()
    {
        const string suppressionBegin = "#define BOOST_NOWIDE_SUPPRESS_UTF_CODECVT_DEPRECATION_BEGIN ";
        string content =
            "// Suppress C++20 deprecation warnings on std::codecvt<char16/32_t,...>.\n" +
            "#if defined(_MSC_VER)\n" +
            suppressionBegin + "__pragma(warning(push)) __pragma(warning(disable : 4996))\n" +
            "#define BOOST_NOWIDE_SUPPRESS_UTF_CODECVT_DEPRECATION_END   __pragma(warning(pop))\n" +
            "#elif(__cplusplus >= 202002L) && defined(__clang__)\n" +
            "#define BOOST_NOWIDE_SUPPRESS_UTF_CODECVT_DEPRECATION_BEGIN \\\n" +
            "    _Pragma(\"clang diagnostic push\") _Pragma(\"clang diagnostic ignored \\\"-Wdeprecated-declarations\\\"\")\n" +
            "#define BOOST_NOWIDE_SUPPRESS_UTF_CODECVT_DEPRECATION_END _Pragma(\"clang diagnostic pop\")\n" +
            "#elif(__cplusplus >= 202002L) && defined(__GNUC__)\n" +
            "#define BOOST_NOWIDE_SUPPRESS_UTF_CODECVT_DEPRECATION_BEGIN \\\n" +
            "    _Pragma(\"GCC diagnostic push\") _Pragma(\"GCC diagnostic ignored \\\"-Wdeprecated-declarations\\\"\")\n" +
            "#define BOOST_NOWIDE_SUPPRESS_UTF_CODECVT_DEPRECATION_END _Pragma(\"GCC diagnostic pop\")\n" +
            "#endif\n";

        Newtonsoft.Json.Linq.JObject[] findings =
            [.. BestPracticeAnalyzer.FindDiagnosticSuppressions("nowide.hpp", content)];

        Assert.Equal(3, findings.Length);
        Assert.All(findings, row => Assert.Equal("BP1044", row["code"]?.ToString()));
        Assert.Contains(findings, row => row["symbol"]?.ToString().Contains("__pragma(warning(disable") == true);
        Assert.Contains(findings, row => row["symbol"]?.ToString().Contains("clang diagnostic ignored") == true);
        Assert.Contains(findings, row => row["symbol"]?.ToString().Contains("GCC diagnostic ignored") == true);
    }
}

