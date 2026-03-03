using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed class PatchService
{
    private sealed class FilePatch
    {
        public string OldPath { get; set; } = string.Empty;

        public string NewPath { get; set; } = string.Empty;

        public List<Hunk> Hunks { get; set; } = new List<Hunk>();
    }

    private sealed class Hunk
    {
        public int OriginalStart { get; set; }

        public int OriginalCount { get; set; }

        public int NewStart { get; set; }

        public int NewCount { get; set; }

        public List<HunkLine> Lines { get; set; } = new List<HunkLine>();
    }

    private sealed class HunkLine
    {
        public char Kind { get; set; }

        public string Text { get; set; } = string.Empty;
    }

    public async Task<JObject> ApplyUnifiedDiffAsync(
        DTE2 dte,
        DocumentService documentService,
        string? patchFilePath,
        string? patchText,
        string? baseDirectory,
        bool openChangedFiles,
        bool saveChangedFiles)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var previousActiveDocument = CaptureActiveDocumentLocation(dte);

        if (string.IsNullOrWhiteSpace(patchFilePath) == string.IsNullOrWhiteSpace(patchText))
        {
            throw new CommandErrorException("invalid_arguments", "Specify exactly one of --patch-file or --patch-text-base64.");
        }

        string patchSource;
        if (!string.IsNullOrWhiteSpace(patchFilePath))
        {
            patchSource = PathNormalization.NormalizeFilePath(patchFilePath);
            if (!File.Exists(patchSource))
            {
                throw new CommandErrorException("document_not_found", $"Patch file not found: {patchSource}");
            }

            patchText = File.ReadAllText(patchSource);
        }
        else
        {
            patchSource = "inline-base64";
            patchText = patchText ?? string.Empty;
        }

        var filePatches = ParseUnifiedDiff(patchText);
        if (filePatches.Count == 0)
        {
            throw new CommandErrorException("invalid_arguments", "Patch file did not contain any unified diff file entries.");
        }

        var resolvedBaseDirectory = ResolveBaseDirectory(dte, baseDirectory);
        var appliedFiles = new JArray();
        var filesToFocus = new List<(string Path, int Line)>();

        foreach (var filePatch in filePatches)
        {
            var target = ResolveTargetPath(dte, resolvedBaseDirectory, filePatch);
            EnsureSafeToModifyOpenDocument(dte, target.Path);

            var result = ApplyFilePatch(target.Path, filePatch);

            if (result.DeleteFile)
            {
                CloseOpenDocumentIfPresent(dte, target.Path);
                if (File.Exists(target.Path))
                {
                    File.Delete(target.Path);
                }

                appliedFiles.Add(new JObject
                {
                    ["path"] = target.Path,
                    ["status"] = "deleted",
                    ["firstChangedLine"] = result.FirstChangedLine,
                    ["hunkCount"] = filePatch.Hunks.Count,
                });
            }
            else
            {
                var writeResult = await documentService.WriteDocumentTextAsync(
                    dte,
                    target.Path,
                    result.Content,
                    result.FirstChangedLine,
                    1,
                    saveChangedFiles).ConfigureAwait(true);
                filesToFocus.Add((target.Path, result.FirstChangedLine));

                appliedFiles.Add(new JObject
                {
                    ["path"] = target.Path,
                    ["status"] = target.IsNewFile ? "added" : "modified",
                    ["firstChangedLine"] = result.FirstChangedLine,
                    ["hunkCount"] = filePatch.Hunks.Count,
                    ["editorBacked"] = writeResult["editorBacked"] ?? false,
                    ["saved"] = writeResult["saved"] ?? saveChangedFiles,
                });
            }
        }

        if (openChangedFiles && filesToFocus.Count > 0)
        {
            await documentService.OpenDocumentAsync(dte, filesToFocus[0].Path, filesToFocus[0].Line, 1).ConfigureAwait(true);
        }
        else
        {
            await RestoreActiveDocumentAsync(dte, documentService, previousActiveDocument).ConfigureAwait(true);
        }

        return new JObject
        {
            ["patchSource"] = patchSource,
            ["baseDirectory"] = resolvedBaseDirectory,
            ["count"] = appliedFiles.Count,
            ["openChangedFiles"] = openChangedFiles,
            ["saveChangedFiles"] = saveChangedFiles,
            ["visibleEdits"] = true,
            ["items"] = appliedFiles,
        };
    }

    private static void EnsureSafeToModifyOpenDocument(DTE2 dte, string path)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var normalizedPath = PathNormalization.NormalizeFilePath(path);
        var openDocument = dte.Documents.Cast<Document>().FirstOrDefault(document =>
            !string.IsNullOrWhiteSpace(document.FullName) &&
            PathNormalization.AreEquivalent(document.FullName, normalizedPath));

        if (openDocument is null)
        {
            return;
        }

        if (!openDocument.Saved)
        {
            throw new CommandErrorException("unsupported_operation", $"Open document has unsaved changes: {normalizedPath}");
        }
    }

    private static void CloseOpenDocumentIfPresent(DTE2 dte, string path)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var normalizedPath = PathNormalization.NormalizeFilePath(path);
        var openDocument = dte.Documents.Cast<Document>().FirstOrDefault(document =>
            !string.IsNullOrWhiteSpace(document.FullName) &&
            PathNormalization.AreEquivalent(document.FullName, normalizedPath));

        if (openDocument is null)
        {
            return;
        }

        if (!openDocument.Saved)
        {
            throw new CommandErrorException("unsupported_operation", $"Open document has unsaved changes: {normalizedPath}");
        }

        openDocument.Close(vsSaveChanges.vsSaveChangesNo);
    }

    private static (string Path, bool IsNewFile) ResolveTargetPath(DTE2 dte, string baseDirectory, FilePatch patch)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var targetRelativePath = patch.NewPath == "/dev/null" ? patch.OldPath : patch.NewPath;
        var isNewFile = patch.OldPath == "/dev/null";
        if (string.IsNullOrWhiteSpace(targetRelativePath) || targetRelativePath == "/dev/null")
        {
            throw new CommandErrorException("invalid_arguments", "Patch entry did not contain a usable target path.");
        }

        if (Path.IsPathRooted(targetRelativePath))
        {
            return (PathNormalization.NormalizeFilePath(targetRelativePath), isNewFile);
        }

        var solutionDirectory = dte.Solution?.IsOpen == true
            ? Path.GetDirectoryName(dte.Solution.FullName) ?? baseDirectory
            : baseDirectory;

        var searchRoots = new List<string>();
        if (!string.IsNullOrWhiteSpace(baseDirectory))
        {
            searchRoots.Add(baseDirectory);
        }

        var current = solutionDirectory;
        for (var depth = 0; depth < 6 && !string.IsNullOrWhiteSpace(current); depth++)
        {
            searchRoots.Add(current);
            current = Path.GetDirectoryName(current);
        }

        foreach (var root in searchRoots.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            var candidate = PathNormalization.NormalizeFilePath(Path.Combine(root, targetRelativePath));
            if (File.Exists(candidate))
            {
                return (candidate, isNewFile);
            }

            if (isNewFile && Directory.Exists(Path.GetDirectoryName(candidate)!))
            {
                return (candidate, true);
            }
        }

        return (PathNormalization.NormalizeFilePath(Path.Combine(baseDirectory, targetRelativePath)), isNewFile);
    }

    private static string ResolveBaseDirectory(DTE2 dte, string? baseDirectory)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (!string.IsNullOrWhiteSpace(baseDirectory))
        {
            return PathNormalization.NormalizeFilePath(baseDirectory);
        }

        if (dte.Solution?.IsOpen == true)
        {
            var solutionDirectory = Path.GetDirectoryName(dte.Solution.FullName);
            if (!string.IsNullOrWhiteSpace(solutionDirectory))
            {
                return PathNormalization.NormalizeFilePath(solutionDirectory);
            }
        }

        return PathNormalization.NormalizeFilePath(Environment.CurrentDirectory);
    }

    private static (string Path, int Line, int Column)? CaptureActiveDocumentLocation(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var activeDocument = dte.ActiveDocument;
        if (activeDocument is null || string.IsNullOrWhiteSpace(activeDocument.FullName))
        {
            return null;
        }

        var line = 1;
        var column = 1;
        try
        {
            if (activeDocument.Object("TextDocument") is TextDocument textDocument)
            {
                line = Math.Max(1, textDocument.Selection.ActivePoint.Line);
                column = Math.Max(1, textDocument.Selection.ActivePoint.DisplayColumn);
            }
        }
        catch
        {
        }

        return (PathNormalization.NormalizeFilePath(activeDocument.FullName), line, column);
    }

    private static async Task RestoreActiveDocumentAsync(
        DTE2 dte,
        DocumentService documentService,
        (string Path, int Line, int Column)? previousActiveDocument)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        if (previousActiveDocument is null || string.IsNullOrWhiteSpace(previousActiveDocument.Value.Path))
        {
            return;
        }

        if (!File.Exists(previousActiveDocument.Value.Path))
        {
            return;
        }

        await documentService.OpenDocumentAsync(
            dte,
            previousActiveDocument.Value.Path,
            previousActiveDocument.Value.Line,
            previousActiveDocument.Value.Column).ConfigureAwait(true);
    }

    private static (string Content, int FirstChangedLine, bool DeleteFile) ApplyFilePatch(string path, FilePatch patch)
    {
        var existingText = File.Exists(path) ? File.ReadAllText(path) : string.Empty;
        var newline = DetectNewline(existingText);
        var existingLines = SplitLines(existingText, out var hadFinalNewline);
        var resultLines = new List<string>();
        var sourceIndex = 0;
        var firstChangedLine = 1;
        var firstChangeCaptured = false;

        foreach (var hunk in patch.Hunks)
        {
            var targetIndex = Math.Max(0, hunk.OriginalStart - 1);
            if (targetIndex < sourceIndex)
            {
                throw new CommandErrorException("invalid_arguments", $"Patch hunks overlap for file: {path}");
            }

            while (sourceIndex < targetIndex && sourceIndex < existingLines.Count)
            {
                resultLines.Add(existingLines[sourceIndex]);
                sourceIndex++;
            }

            foreach (var line in hunk.Lines)
            {
                switch (line.Kind)
                {
                    case ' ':
                        EnsureLineMatches(path, existingLines, sourceIndex, line.Text, "context");
                        resultLines.Add(existingLines[sourceIndex]);
                        sourceIndex++;
                        break;
                    case '-':
                        EnsureLineMatches(path, existingLines, sourceIndex, line.Text, "deletion");
                        if (!firstChangeCaptured)
                        {
                            firstChangedLine = Math.Max(1, resultLines.Count + 1);
                            firstChangeCaptured = true;
                        }

                        sourceIndex++;
                        break;
                    case '+':
                        if (!firstChangeCaptured)
                        {
                            firstChangedLine = Math.Max(1, resultLines.Count + 1);
                            firstChangeCaptured = true;
                        }

                        resultLines.Add(line.Text);
                        break;
                    default:
                        throw new CommandErrorException("invalid_arguments", $"Unsupported hunk line prefix '{line.Kind}' in patch for {path}.");
                }
            }
        }

        while (sourceIndex < existingLines.Count)
        {
            resultLines.Add(existingLines[sourceIndex]);
            sourceIndex++;
        }

        var deleteFile = patch.NewPath == "/dev/null";
        var content = JoinLines(resultLines, newline, deleteFile ? false : hadFinalNewline || patch.OldPath == "/dev/null");
        return (content, firstChangedLine, deleteFile);
    }

    private static void EnsureLineMatches(string path, IReadOnlyList<string> existingLines, int index, string expected, string operation)
    {
        if (index >= existingLines.Count)
        {
            throw new CommandErrorException("invalid_arguments", $"Patch {operation} exceeded file length for {path}.");
        }

        if (!string.Equals(existingLines[index], expected, StringComparison.Ordinal))
        {
            throw new CommandErrorException(
                "invalid_arguments",
                $"Patch {operation} mismatch in {path} at line {index + 1}.",
                new { expected, actual = existingLines[index], line = index + 1 });
        }
    }

    private static List<FilePatch> ParseUnifiedDiff(string patchText)
    {
        var lines = patchText.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var patches = new List<FilePatch>();
        FilePatch? currentFile = null;
        var lineIndex = 0;

        while (lineIndex < lines.Length)
        {
            var line = lines[lineIndex];
            if (line.StartsWith("--- ", StringComparison.Ordinal))
            {
                var oldPath = NormalizePatchPath(line.Substring(4));
                lineIndex++;
                if (lineIndex >= lines.Length || !lines[lineIndex].StartsWith("+++ ", StringComparison.Ordinal))
                {
                    throw new CommandErrorException("invalid_arguments", "Unified diff is missing a +++ header.");
                }

                var newPath = NormalizePatchPath(lines[lineIndex].Substring(4));
                currentFile = new FilePatch
                {
                    OldPath = oldPath,
                    NewPath = newPath,
                    Hunks = new List<Hunk>(),
                };
                patches.Add(currentFile);
                lineIndex++;
                continue;
            }

            if (line.StartsWith("@@ ", StringComparison.Ordinal))
            {
                if (currentFile is null)
                {
                    throw new CommandErrorException("invalid_arguments", "Encountered a hunk before a file header.");
                }

                var hunk = ParseHunkHeader(line);
                lineIndex++;
                while (lineIndex < lines.Length)
                {
                    var hunkLine = lines[lineIndex];
                    if (hunkLine.StartsWith("@@ ", StringComparison.Ordinal) || hunkLine.StartsWith("--- ", StringComparison.Ordinal))
                    {
                        break;
                    }

                    if (hunkLine == "\\ No newline at end of file")
                    {
                        lineIndex++;
                        continue;
                    }

                    if (hunkLine.Length == 0)
                    {
                        lineIndex++;
                        continue;
                    }

                    var prefix = hunkLine[0];
                    if (prefix != ' ' && prefix != '+' && prefix != '-')
                    {
                        throw new CommandErrorException("invalid_arguments", $"Unsupported hunk line prefix '{prefix}'.");
                    }

                    hunk.Lines.Add(new HunkLine
                    {
                        Kind = prefix,
                        Text = hunkLine.Length > 1 ? hunkLine.Substring(1) : string.Empty,
                    });
                    lineIndex++;
                }

                currentFile.Hunks.Add(hunk);
                continue;
            }

            lineIndex++;
        }

        return patches;
    }

    private static Hunk ParseHunkHeader(string line)
    {
        var match = Regex.Match(line, @"^@@ -(?<oldStart>\d+)(,(?<oldCount>\d+))? \+(?<newStart>\d+)(,(?<newCount>\d+))? @@");
        if (!match.Success)
        {
            throw new CommandErrorException("invalid_arguments", $"Invalid unified diff hunk header: {line}");
        }

        return new Hunk
        {
            OriginalStart = int.Parse(match.Groups["oldStart"].Value),
            OriginalCount = ParseHunkCount(match.Groups["oldCount"].Value),
            NewStart = int.Parse(match.Groups["newStart"].Value),
            NewCount = ParseHunkCount(match.Groups["newCount"].Value),
            Lines = new List<HunkLine>(),
        };
    }

    private static int ParseHunkCount(string value)
    {
        return string.IsNullOrWhiteSpace(value) ? 1 : int.Parse(value);
    }

    private static string NormalizePatchPath(string value)
    {
        var trimmed = value.Trim();
        if (trimmed == "/dev/null")
        {
            return trimmed;
        }

        if ((trimmed.StartsWith("a/", StringComparison.Ordinal) || trimmed.StartsWith("b/", StringComparison.Ordinal)) && trimmed.Length > 2)
        {
            return trimmed.Substring(2);
        }

        return trimmed;
    }

    private static string DetectNewline(string text)
    {
        return text.Contains("\r\n", StringComparison.Ordinal) ? "\r\n" : "\n";
    }

    private static List<string> SplitLines(string text, out bool hadFinalNewline)
    {
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        hadFinalNewline = normalized.EndsWith("\n", StringComparison.Ordinal);
        if (hadFinalNewline)
        {
            normalized = normalized.Substring(0, normalized.Length - 1);
        }

        return normalized.Length == 0 ? new List<string>() : normalized.Split('\n').ToList();
    }

    private static string JoinLines(IReadOnlyList<string> lines, string newline, bool includeTrailingNewline)
    {
        var content = string.Join(newline, lines);
        if (includeTrailingNewline)
        {
            content += newline;
        }

        return content;
    }
}
