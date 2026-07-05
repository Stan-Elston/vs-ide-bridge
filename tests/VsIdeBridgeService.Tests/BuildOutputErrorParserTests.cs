using System.Collections.Generic;
using VsIdeBridge.Tooling.Build;
using Xunit;

namespace VsIdeBridgeService.Tests;

public class BuildOutputErrorParserTests
{
    [Fact]
    public void ParsesLinkerErrorsWithoutLocation()
    {
        // Repro from the 2026-07-01 SuperSlicer session: these never reached the Error List.
        string output =
            "1>glu-libtess.lib(sweep.obj) : error LNK2001: unresolved external symbol DebugEvent\r\n" +
            "1>C:\\repos\\SuperSlicer\\build-default\\src\\Slic3r.dll : fatal error LNK1120: 1 unresolved externals\r\n" +
            "2>stl_to_cpp.exe : fatal error LNK1120: 1 unresolved externals\r\n";

        IReadOnlyList<BuildOutputError> rows = BuildOutputErrorParser.ParseErrors(output);

        Assert.Equal(3, rows.Count);
        Assert.Equal("LNK2001", rows[0].Code);
        Assert.Equal("unresolved external symbol DebugEvent", rows[0].Message);
        Assert.Equal("glu-libtess.lib(sweep.obj)", rows[0].File);
        Assert.Null(rows[0].Line);
        Assert.Equal("LNK1120", rows[1].Code);
        Assert.Equal("LNK1120", rows[2].Code);
        Assert.Equal("stl_to_cpp.exe", rows[2].File);
    }

    [Fact]
    public void ParsesCompilerErrorsWithLineAndColumn()
    {
        string output = "3>C:\\repos\\src\\Foo.cpp(123,45): error C2065: 'bar': undeclared identifier\n";

        IReadOnlyList<BuildOutputError> rows = BuildOutputErrorParser.ParseErrors(output);

        BuildOutputError row = Assert.Single(rows);
        Assert.Equal("C2065", row.Code);
        Assert.Equal("C:\\repos\\src\\Foo.cpp", row.File);
        Assert.Equal(123, row.Line);
        Assert.Equal(45, row.Column);
    }

    [Fact]
    public void ParsesMsbuildErrors()
    {
        string output = @"C:\repos\proj\App.vcxproj(21,3): error MSB4278: The imported file does not exist.";

        IReadOnlyList<BuildOutputError> rows = BuildOutputErrorParser.ParseErrors(output);

        BuildOutputError row = Assert.Single(rows);
        Assert.Equal("MSB4278", row.Code);
        Assert.Equal(21, row.Line);
    }

    [Fact]
    public void ParseDiagnosticsKeepsWarningsAndProjectOriginForExternalHeaderNotes()
    {
        string output =
            "1>C:\\Program Files\\Microsoft Visual Studio\\2022\\Community\\VC\\Tools\\MSVC\\14.44.35207\\include\\vector(123,45): warning C4244: 'argument': conversion from 'double' to 'int'\r\n" +
            "1>        C:\\Users\\elsto\\source\\repos\\SuperSlicer\\src\\libslic3r\\Foo.cpp(456,12): note: see reference to function template instantiation 'void foo()' being compiled\r\n";

        IReadOnlyList<BuildOutputError> rows = BuildOutputErrorParser.ParseDiagnostics(output);

        BuildOutputError row = Assert.Single(rows);
        Assert.Equal("Warning", row.Severity);
        Assert.Equal("C4244", row.Code);
        Assert.Equal("C:\\Program Files\\Microsoft Visual Studio\\2022\\Community\\VC\\Tools\\MSVC\\14.44.35207\\include\\vector", row.File);
        Assert.True(row.HasProjectOrigin);
        Assert.Equal("C:\\Users\\elsto\\source\\repos\\SuperSlicer\\src\\libslic3r\\Foo.cpp", row.OriginFile);
        Assert.Equal(456, row.OriginLine);
        Assert.Equal(12, row.OriginColumn);
        Assert.Empty(BuildOutputErrorParser.ParseErrors(output));
    }

    [Fact]
    public void ParseDiagnosticsUsesCompilingSourceFileAsExternalHeaderOriginFallback()
    {
        string output =
            "1>C:\\Program Files\\Microsoft Visual Studio\\2022\\Community\\VC\\Tools\\MSVC\\14.44.35207\\include\\xmemory(99): warning C4267: conversion from 'size_t' to 'int'\n" +
            "1>  (compiling source file '../../../../src/libslic3r/Bar.cpp')\n";

        BuildOutputError row = Assert.Single(BuildOutputErrorParser.ParseDiagnostics(output));

        Assert.Equal("Warning", row.Severity);
        Assert.True(row.HasProjectOrigin);
        Assert.Equal("../../../../src/libslic3r/Bar.cpp", row.OriginFile);
        Assert.Null(row.OriginLine);
    }

    [Fact]
    public void ParseDiagnosticsFoldsMsvcTemplateContinuationLinesIntoPrimaryWarning()
    {
        string output =
            "5>C:\\Program Files\\Microsoft Visual Studio\\2022\\Community\\VC\\Tools\\MSVC\\14.44.35207\\include\\optional(83,17): warning C4244: 'initializing': conversion from '_Ty' to 'int', possible loss of data\n" +
            "5>C:\\Program Files\\Microsoft Visual Studio\\2022\\Community\\VC\\Tools\\MSVC\\14.44.35207\\include\\optional(83,17): warning C4244:         with\n" +
            "5>C:\\Program Files\\Microsoft Visual Studio\\2022\\Community\\VC\\Tools\\MSVC\\14.44.35207\\include\\optional(83,17): warning C4244:         [\n" +
            "5>C:\\Program Files\\Microsoft Visual Studio\\2022\\Community\\VC\\Tools\\MSVC\\14.44.35207\\include\\optional(83,17): warning C4244:             _Ty=__int64\n" +
            "5>C:\\Program Files\\Microsoft Visual Studio\\2022\\Community\\VC\\Tools\\MSVC\\14.44.35207\\include\\optional(83,17): warning C4244:         ]\n" +
            "5>  (compiling source file '../../../src/libslic3r/Config.cpp')\n";

        BuildOutputError row = Assert.Single(BuildOutputErrorParser.ParseDiagnostics(output));

        Assert.Equal("Warning", row.Severity);
        Assert.Equal("C4244", row.Code);
        Assert.Contains("possible loss of data", row.Message);
        Assert.True(row.HasProjectOrigin);
        Assert.Equal("../../../src/libslic3r/Config.cpp", row.OriginFile);
    }

    [Fact]
    public void DeduplicatesRepeatedErrorLines()
    {
        string line = "lib.obj : error LNK2001: unresolved external symbol Foo";
        string output = line + "\n" + line + "\n";

        Assert.Single(BuildOutputErrorParser.ParseErrors(output));
    }

    [Fact]
    public void IgnoresNonErrorLines()
    {
        string output =
            "1>Build started...\n" +
            "========== Build: 0 succeeded, 1 failed, 0 up-to-date ==========\n" +
            "1>Done building project (Build target(s)) -- FAILED.\n" +
            "    0 Error(s)\n";

        Assert.Empty(BuildOutputErrorParser.ParseErrors(output));
    }

    [Fact]
    public void ReturnsEmptyForNullOrBlankInput()
    {
        Assert.Empty(BuildOutputErrorParser.ParseErrors(null));
        Assert.Empty(BuildOutputErrorParser.ParseErrors("   "));
    }

    [Fact]
    public void CapsRowCount()
    {
        System.Text.StringBuilder output = new();
        for (int i = 0; i < BuildOutputErrorParser.MaxRows + 10; i++)
        {
            output.AppendLine($"lib{i}.obj : error LNK2001: unresolved external symbol Sym{i}");
        }

        Assert.Equal(BuildOutputErrorParser.MaxRows, BuildOutputErrorParser.ParseErrors(output.ToString()).Count);
    }
}
