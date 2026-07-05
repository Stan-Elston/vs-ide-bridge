using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using VsIdeBridge.Infrastructure;
using VsIdeBridge.Diagnostics;
using VsIdeBridge.Shared;

namespace VsIdeBridge.Services;

internal sealed partial class PatchService
{
    private const int PatchedContentVerificationAttemptCount = 3;
    private const int PatchedContentVerificationDelayMilliseconds = 50;

    /// <summary>Injected after construction; resolves handle IDs in path arguments.</summary>
    internal HandleService? HandleService { get; set; }

    public async Task<JObject> ApplyEditorPatchAsync(
        DTE2 dte,
        DocumentService documentService,
        string? patchFilePath,
        string? patchText,
        string? baseDirectory,
        bool openChangedFiles,
        bool saveChangedFiles,
        bool includeBestPracticeWarnings = false,
        bool replaceAll = false)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        (string Path, int Line, int Column)? previousActiveDocument = CaptureActiveDocumentLocation(dte);

        if (string.IsNullOrWhiteSpace(patchFilePath) == string.IsNullOrWhiteSpace(patchText))
        {
            throw new CommandErrorException(InvalidArgumentsCode, "Specify exactly one of --patch-file or --patch-text-base64.");
        }

        string patchSource;
        if (!string.IsNullOrWhiteSpace(patchFilePath))
        {
            patchSource = PathNormalization.NormalizeFilePath(patchFilePath);
            if (!File.Exists(patchSource))
            {
                throw new CommandErrorException(
                    "document_not_found",
                    $"Patch file not found: {patchSource}. Verify the path exists on disk and retry with the correct absolute path, or omit --patch-file and use --patch-text-base64 to pass the patch inline.");
            }

            patchText = File.ReadAllText(patchSource);
        }
        else
        {
            patchSource = "inline-base64";
            patchText ??= string.Empty;
        }

        var (filePatches, patchFormat) = ParseSupportedPatchFormat(patchText);
        if (filePatches.Count == 0)
        {
            throw new CommandErrorException(InvalidArgumentsCode, BuildMissingPatchFormatMessage(patchText));
        }

        if (replaceAll)
        {
            foreach (FilePatch filePatch in filePatches)
            {
                filePatch.ReplaceAll = true;
            }
        }

        string resolvedBaseDirectory = ResolveBaseDirectory(dte, baseDirectory);
        JArray appliedFiles = [];
        List<(string Path, List<ChangedRange> Ranges)> filesToFocus = [];
        List<PreparedPatchOperation> preparedOperations = [];

        foreach (var filePatch in filePatches)
        {
            PatchPaths paths = ResolvePatchPaths(dte, resolvedBaseDirectory, filePatch);
            EnsurePatchHasMeaningfulOperations(filePatch, paths);
            EnsureSafeToModifyOpenDocument(dte, paths.SourcePath);

            // Auto-reload if the document is open in the editor.
            // Prevents state drift between patches — very important for local models.
            if (IsDocumentOpenInEditor(dte, paths.SourcePath))
            {
                await documentService.ReloadDocumentAsync(paths.SourcePath).ConfigureAwait(true);
            }

            if (paths.IsMove)
            {
                EnsureSafeToModifyOpenDocument(dte, paths.TargetPath);
                if (File.Exists(paths.TargetPath))
                {
                    throw new CommandErrorException(
                        "unsupported_operation",
                        $"Patch move target already exists: {paths.TargetPath}. Delete or rename the existing file first, or change the *** Move to: path in the patch to a location that does not exist.");
                }
            }

            bool targetExists = PatchTargetExists(dte, paths.TargetPath);
            string sourceContent = paths.IsNewFile ? string.Empty : ReadContentFromEditorOrDisk(dte, paths.SourcePath);
            string requestedTargetContent = paths.IsNewFile && !targetExists
                ? string.Empty
                : PathNormalization.AreEquivalent(paths.SourcePath, paths.TargetPath) && !paths.IsNewFile
                    ? sourceContent
                    : ReadContentFromEditorOrDisk(dte, paths.TargetPath);
            ApplyFilePatchResult result;
            bool alreadySatisfied = false;
            if (paths.IsNewFile)
            {
                result = ApplyFilePatch(paths.TargetPath, string.Empty, filePatch);
                AddFilePatchDecision addFileDecision = AddFilePatchSemantics.Evaluate(result.Content, targetExists ? requestedTargetContent : null);
                switch (addFileDecision)
                {
                    case AddFilePatchDecision.Create:
                        break;
                    case AddFilePatchDecision.AlreadySatisfied:
                        result = CreateAlreadySatisfiedResult(requestedTargetContent, result);
                        alreadySatisfied = true;
                        break;
                    default:
                        throw new CommandErrorException(
                            "unsupported_operation",
                            $"Patch add target already exists with different content: {paths.TargetPath}. Use *** Update File: instead of *** Add File: in the patch header to modify an existing file.");
                }
            }
            else
            {
                try
                {
                    result = ApplyFilePatch(paths.SourcePath, sourceContent, filePatch);
                }
                catch (CommandErrorException ex) when (!paths.IsMove && IsPatchContentMismatch(ex))
                {
                    if (!IsCurrentContentAlreadyPatched(paths.TargetPath, requestedTargetContent, filePatch))
                    {
                        throw;
                    }

                    ApplyFilePatchResult reverseResult = ApplyFilePatch(paths.TargetPath, requestedTargetContent, CreateReversePatch(filePatch));
                    result = CreateAlreadySatisfiedResult(requestedTargetContent, reverseResult);
                    alreadySatisfied = true;
                }
            }

            preparedOperations.Add(new PreparedPatchOperation
            {
                FilePatch = filePatch,
                Paths = paths,
                Result = result,
                RequestedTargetContent = requestedTargetContent,
                AlreadySatisfied = alreadySatisfied,
            });
        }

        List<PreparedPatchOperation> orderedOperations =
        [
            .. preparedOperations.Where(operation => !operation.Result.DeleteFile),
            .. preparedOperations.Where(operation => operation.Result.DeleteFile),
        ];

        foreach (PreparedPatchOperation operation in orderedOperations)
        {
            FilePatch filePatch = operation.FilePatch;
            PatchPaths paths = operation.Paths;
            ApplyFilePatchResult result = operation.Result;
            string requestedTargetContent = operation.RequestedTargetContent;
            bool alreadySatisfied = operation.AlreadySatisfied;

            if (result.DeleteFile)
            {
                CloseOpenDocumentIfPresent(dte, paths.SourcePath);
                if (File.Exists(paths.SourcePath))
                {
                    File.Delete(paths.SourcePath);
                }

                appliedFiles.Add(new JObject
                {
                    ["path"] = paths.SourcePath,
                    ["status"] = "deleted",
                    ["firstChangedLine"] = result.FirstChangedLine,
                    ["hunkCount"] = filePatch.Hunks.Count + filePatch.SearchBlocks.Count,
                    ["changedRanges"] = CreateChangedRangesArray(result.ChangedRanges),
                });
            }
            else
            {
                bool requestedContentChange = !string.Equals(requestedTargetContent, result.Content, StringComparison.Ordinal);
                if (!requestedContentChange && alreadySatisfied)
                {
                    if (CanRevealEditedFile(paths.TargetPath))
                    {
                        filesToFocus.Add((paths.TargetPath, result.ChangedRanges));
                    }

                    appliedFiles.Add(new JObject
                    {
                        ["path"] = paths.TargetPath,
                        ["status"] = "already-satisfied",
                        ["firstChangedLine"] = result.FirstChangedLine,
                        ["hunkCount"] = filePatch.Hunks.Count + filePatch.SearchBlocks.Count,
                        ["changedRanges"] = CreateChangedRangesArray(result.ChangedRanges),
                        ["matchedLineCount"] = result.MatchedLineCount,
                        ["mutationLineCount"] = result.MutationLineCount,
                        ["editorBacked"] = false,
                        ["verified"] = true,
                        ["contentChanged"] = false,
                        ["requestedContentChange"] = false,
                        ["alreadySatisfied"] = true,
                        ["saved"] = true,
                    });
                    continue;
                }

                if (!requestedContentChange && !paths.IsMove && !paths.IsNewFile)
                {
                    alreadySatisfied = IsCurrentContentAlreadyPatched(paths.TargetPath, requestedTargetContent, filePatch);
                    ThrowIfPatchProducedNoContentChange(paths.TargetPath, filePatch, result, alreadySatisfied);
                    if (CanRevealEditedFile(paths.TargetPath))
                    {
                        filesToFocus.Add((paths.TargetPath, result.ChangedRanges));
                    }
                    appliedFiles.Add(CreateAlreadySatisfiedFileItem(paths.TargetPath, filePatch, result, alreadySatisfied));
                    continue;
                }

                Directory.CreateDirectory(Path.GetDirectoryName(paths.TargetPath)!);

                // Pre-write best-practice analysis - only when caller opts in.
                IReadOnlyList<JObject> preWriteWarnings = includeBestPracticeWarnings
                    ? ErrorListService.AnalyzeContentBeforeWrite(paths.TargetPath, result.Content)
                    : [];
                string? projectUniqueName = preWriteWarnings.Count > 0
                    ? SolutionFileLocator.TryFindProjectUniqueName(dte, paths.TargetPath)
                    : null;

                JObject writeResult = await documentService.WriteDocumentTextAsync(
                    dte,
                    paths.TargetPath,
                    result.Content,
                    result.FirstChangedLine,
                    1,
                    saveChangedFiles,
                    [.. result.ChangedRanges.Select(range => (range.StartLine, range.EndLine))],
                    result.DeletedLineMarkers).ConfigureAwait(true);

                bool contentChanged = (bool?)writeResult["contentChanged"] ?? true;
                bool verified = await VerifyPatchedContentAsync(dte, documentService, paths.TargetPath, result.Content).ConfigureAwait(true);
                alreadySatisfied = requestedContentChange && !contentChanged && verified;

                if (paths.IsMove)
                {
                    CloseOpenDocumentIfPresent(dte, paths.SourcePath);
                    if (File.Exists(paths.SourcePath))
                    {
                        File.Delete(paths.SourcePath);
                    }
                }

                if (CanRevealEditedFile(paths.TargetPath))
                {
                    filesToFocus.Add((paths.TargetPath, result.ChangedRanges));
                }

                JObject fileItem = new()
                {
                    ["path"] = paths.TargetPath,
                    ["status"] = paths.IsMove ? "moved" : alreadySatisfied ? "already-satisfied" : paths.IsNewFile ? "added" : "modified",
                    ["firstChangedLine"] = result.FirstChangedLine,
                    ["hunkCount"] = filePatch.Hunks.Count + filePatch.SearchBlocks.Count,
                    ["changedRanges"] = CreateChangedRangesArray(result.ChangedRanges),
                    ["editorBacked"] = writeResult["editorBacked"] ?? false,
                    ["verified"] = verified,
                    ["contentChanged"] = contentChanged,
                    ["requestedContentChange"] = requestedContentChange,
                    ["alreadySatisfied"] = alreadySatisfied,
                    ["saved"] = writeResult["saved"] ?? saveChangedFiles,
                };

                if (preWriteWarnings.Count > 0)
                {
                    fileItem["bestPracticeWarnings"] = BestPracticeWarningProjector.CreateResponseWarnings(preWriteWarnings, projectUniqueName);
                }

                appliedFiles.Add(fileItem);
            }
        }

        if (openChangedFiles && filesToFocus.Count > 0)
        {
            ChangedRange? primaryRange = filesToFocus[0].Ranges.FirstOrDefault();
            if (primaryRange is not null)
            {
                await documentService.RevealDocumentRangeAsync(
                    dte,
                    filesToFocus[0].Path,
                    primaryRange.StartLine,
                    1,
                    primaryRange.EndLine,
                    1).ConfigureAwait(true);
            }
            else
            {
                await documentService.OpenDocumentAsync(dte, filesToFocus[0].Path, 1, 1).ConfigureAwait(true);
            }
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
            ["patchFormat"] = patchFormat,
            ["openChangedFiles"] = openChangedFiles,
            ["saveChangedFiles"] = saveChangedFiles,
            ["visibleEdits"] = openChangedFiles && filesToFocus.Count > 0,
            ["items"] = appliedFiles,
        };
    }

    private static bool CanRevealEditedFile(string path)
    {
        return !IsDiskBackedPatchFile(path);
    }

    private static bool IsDiskBackedPatchFile(string path)
    {
        string extension = Path.GetExtension(path);
        return string.Equals(extension, ".csproj", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".vcxproj", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".vbproj", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".fsproj", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".props", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".targets", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".sln", StringComparison.OrdinalIgnoreCase);
    }

    private static void ThrowIfPatchProducedNoContentChange(string targetPath, FilePatch filePatch, ApplyFilePatchResult result, bool alreadySatisfied)
    {
        if (alreadySatisfied || result.MutationLineCount <= 0 || result.MatchedLineCount != 0)
        {
            return;
        }

        throw new CommandErrorException(
            InvalidArgumentsCode,
            $"Patch produced no content change for {targetPath} - {result.MutationLineCount} mutation line(s) but 0 matched context lines. " +
            "The patch content does not match the file. Fix: call read_file to check the actual content before retrying.",
            new
            {
                path = targetPath,
                hunkCount = filePatch.Hunks.Count + filePatch.SearchBlocks.Count,
                result.MutationLineCount,
                result.MatchedLineCount,
            });
    }

    private static JObject CreateAlreadySatisfiedFileItem(string targetPath, FilePatch filePatch, ApplyFilePatchResult result, bool alreadySatisfied)
    {
        return new JObject
        {
            ["path"] = targetPath,
            ["status"] = "already-satisfied",
            ["firstChangedLine"] = result.FirstChangedLine,
            ["hunkCount"] = filePatch.Hunks.Count + filePatch.SearchBlocks.Count,
            ["changedRanges"] = CreateChangedRangesArray(result.ChangedRanges),
            ["matchedLineCount"] = result.MatchedLineCount,
            ["mutationLineCount"] = result.MutationLineCount,
            ["editorBacked"] = false,
            ["verified"] = true,
            ["contentChanged"] = false,
            ["requestedContentChange"] = false,
            ["alreadySatisfied"] = alreadySatisfied,
            ["saved"] = true,
        };
    }

    private static async Task<bool> VerifyPatchedContentAsync(DTE2 dte, DocumentService documentService, string targetPath, string expectedContent)
    {
        string currentContent = string.Empty;
        for (int attempt = 0; attempt < PatchedContentVerificationAttemptCount; attempt++)
        {
            currentContent = await ReadPatchVerificationContentAsync(dte, targetPath).ConfigureAwait(true);
            if (string.Equals(currentContent, expectedContent, StringComparison.Ordinal))
            {
                return true;
            }

            if (attempt == 0)
            {
                await documentService.ReloadDocumentAsync(targetPath).ConfigureAwait(true);
            }

            if (attempt < PatchedContentVerificationAttemptCount - 1)
            {
                await Task.Delay(PatchedContentVerificationDelayMilliseconds).ConfigureAwait(true);
            }
        }

        throw new CommandErrorException(
            "write_failed",
            $"Patched content could not be verified after write: {targetPath}",
            new
            {
                path = targetPath,
                expectedLength = expectedContent.Length,
                actualLength = currentContent.Length,
                firstDifferenceIndex = FindFirstDifferenceIndex(expectedContent, currentContent),
            });
    }

    private static async Task<string> ReadPatchVerificationContentAsync(DTE2 dte, string targetPath)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        return ReadContentFromEditorOrDisk(dte, targetPath);
    }

    private static int FindFirstDifferenceIndex(string expectedContent, string currentContent)
    {
        int comparisonLength = Math.Min(expectedContent.Length, currentContent.Length);
        for (int index = 0; index < comparisonLength; index++)
        {
            if (expectedContent[index] != currentContent[index])
            {
                return index;
            }
        }

        return expectedContent.Length == currentContent.Length ? -1 : comparisonLength;
    }

    private static bool IsDocumentOpenInEditor(DTE2 dte, string path)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (string.IsNullOrWhiteSpace(path))
            return false;

        foreach (Document doc in dte.Documents)
        {
            if (string.Equals(doc.FullName, path, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

}

