using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using VsIdeBridge.Diagnostics;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed partial class DocumentService
{
    public async Task<JObject> WriteDocumentTextAsync(
        DTE2 dte,
        string filePath,
        string content,
        int line,
        int column,
        bool saveChanges,
        IReadOnlyCollection<(int StartLine, int EndLine)>? changedRanges = null,
        IReadOnlyCollection<int>? deletedLines = null,
        bool includeBestPracticeWarnings = false)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        string normalizedPath = PathNormalization.NormalizeFilePath(filePath);
        string? directory = Path.GetDirectoryName(normalizedPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        if (!File.Exists(normalizedPath))
        {
            File.WriteAllText(normalizedPath, string.Empty);
        }

        // Project and solution files cannot be edited via the text-editor buffer.
        // Skip ItemOperations.OpenFile entirely: old-style .csproj opens a project
        // designer (not a text window) and closing it can unload the project; .vcxproj
        // and others may throw ArgumentException. Write directly to disk instead.
        if (RequiresDiskBackedWrite(normalizedPath))
        {
            return BuildDiskWriteResult(normalizedPath, content, saveChanges, line, column, includeBestPracticeWarnings, dte);
        }

        await ActivateDocumentForWriteAsync(dte, normalizedPath, line, column).ConfigureAwait(true);

        try
        {
            Document? targetDocument = TryFindOpenDocumentByPath(dte, normalizedPath);
            Window window = targetDocument?.ActiveWindow ?? dte.ActiveWindow;
            if (window?.Document is null || !PathNormalization.AreEquivalent(TryGetDocumentFullName(window.Document), normalizedPath))
            {
                window = dte.ItemOperations.OpenFile(normalizedPath);
                window.Activate();
                targetDocument = window.Document ?? TryFindOpenDocumentByPath(dte, normalizedPath);
            }

            Document document = targetDocument ?? window.Document ?? dte.ActiveDocument
                ?? throw new CommandErrorException(DocumentNotFoundCode, $"Unable to activate: {normalizedPath}");

            (bool usedEditorBuffer, string originalContent, window) = ApplyDocumentContent(dte, window, document, normalizedPath, content, line, column);

            Document finalDocument = window.Document ?? document;
            string finalContent = TryReadDocumentText(finalDocument) ?? ReadFileText(normalizedPath);
            if (!WriteContentMatches(content, finalContent))
            {
                throw new CommandErrorException("write_failed",
                    $"Write verification failed for {normalizedPath}: the VS editor buffer ({finalContent.Length} chars) differs from the expected content ({content.Length} chars). " +
                    "VS may have applied auto-formatting or normalized line endings. Check that no editor extension is reformatting on write.");
            }

            bool contentChanged = !string.Equals(originalContent, finalContent, StringComparison.Ordinal);
            if (saveChanges)
            {
                // The window reference can lose its Document after a large buffer replace;
                // re-resolve instead of silently skipping a requested save (issue 17 left
                // a mutated document unsaved while the result still claimed saved:true).
                Document? documentToSave = window.Document ?? TryFindOpenDocumentByPath(dte, normalizedPath) ?? document;
                documentToSave?.Save();
            }

            return BuildWriteResult(window, normalizedPath, content, usedEditorBuffer, contentChanged,
                saveChanges, line, column, changedRanges, deletedLines, includeBestPracticeWarnings, dte);
        }
        catch (ArgumentException) when (RequiresDiskBackedWrite(normalizedPath))
        {
            return BuildDiskWriteResult(normalizedPath, content, saveChanges, line, column, includeBestPracticeWarnings, dte);
        }
    }

    private async Task ActivateDocumentForWriteAsync(DTE2 dte, string normalizedPath, int line, int column)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        if (RequiresDiskBackedWrite(normalizedPath))
        {
            return;
        }

        Document? openDocument = TryFindOpenDocumentByPath(dte, normalizedPath);
        if (openDocument is not null)
        {
            openDocument.Activate();
            if (openDocument.Object(TextDocumentKind) is TextDocument textDocument)
            {
                _ = await TryNavigateToLineAsync(textDocument, line, column).ConfigureAwait(true);
            }

            return;
        }

        _ = await OpenDocumentAsync(dte, normalizedPath, line, column, allowDiskFallback: false).ConfigureAwait(true);
    }

    private JObject BuildDiskWriteResult(
        string normalizedPath,
        string content,
        bool saveChanges,
        int line,
        int column,
        bool includeBestPracticeWarnings,
        DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        string originalContent = ReadFileText(normalizedPath);
        File.WriteAllText(normalizedPath, content);
        string finalContent = ReadFileText(normalizedPath);
        if (!WriteContentMatches(content, finalContent))
        {
            throw new CommandErrorException("write_failed",
                $"Write verification failed for {normalizedPath}: read back {finalContent.Length} chars but expected {content.Length} chars. " +
                "This may be a file encoding issue. Verify the file is writable and that no process is locking or modifying it.");
        }

        IReadOnlyList<JObject> preWriteWarnings = includeBestPracticeWarnings
            ? ErrorListService.AnalyzeContentBeforeWrite(normalizedPath, content)
            : [];
        string? projectUniqueName = includeBestPracticeWarnings
            ? SolutionFileLocator.TryFindProjectUniqueName(dte, normalizedPath)
            : null;

        return new JObject
        {
            [ResolvedPathProperty] = normalizedPath,
            ["editorBacked"] = false,
            ["verified"] = true,
            ["contentChanged"] = !string.Equals(originalContent, finalContent, StringComparison.Ordinal),
            ["saved"] = saveChanges,
            ["line"] = Math.Max(1, line),
            ["column"] = Math.Max(1, column),
            ["windowCaption"] = Path.GetFileName(normalizedPath),
            ["bestPracticeWarnings"] = preWriteWarnings.Count > 0
                ? BestPracticeWarningProjector.CreateResponseWarnings(preWriteWarnings, projectUniqueName)
                : null,
        };
    }

    private static bool RequiresDiskBackedWrite(string normalizedPath)
    {
        string extension = Path.GetExtension(normalizedPath);
        return string.Equals(extension, ".csproj", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".vcxproj", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".vbproj", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".fsproj", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".props", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".targets", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".sln", StringComparison.OrdinalIgnoreCase);
    }

    private static (bool UsedEditorBuffer, string OriginalContent, Window ResultWindow) ApplyDocumentContent(
        DTE2 dte,
        Window window,
        Document document,
        string normalizedPath,
        string content,
        int line,
        int column)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        string originalContent = TryReadDocumentText(document) ?? ReadFileText(normalizedPath);
        bool usedEditorBuffer = false;
        try
        {
            if (document.Object(TextDocumentKind) is TextDocument textDocument)
            {
                EditPoint start = textDocument.StartPoint.CreateEditPoint();
                start.ReplaceText(textDocument.EndPoint, content, 0);
                TextSelection selection = textDocument.Selection;
                selection.MoveToLineAndOffset(Math.Max(1, line), Math.Max(1, column), false);
                TryShowActivePoint(selection);
                usedEditorBuffer = string.Equals(TryReadDocumentText(document), content, StringComparison.Ordinal);
            }
        }
        catch (COMException)
        {
            usedEditorBuffer = false;
        }

        if (!usedEditorBuffer)
        {
            File.WriteAllText(normalizedPath, content);
            if (window.Document is not null)
            {
                window.Document.Close(vsSaveChanges.vsSaveChangesNo);
                window = dte.ItemOperations.OpenFile(normalizedPath);
                window.Activate();
            }
        }

        return (usedEditorBuffer, originalContent, window);
    }

    private JObject BuildWriteResult(
        Window window,
        string normalizedPath,
        string content,
        bool usedEditorBuffer,
        bool contentChanged,
        bool saveChanges,
        int line,
        int column,
        IReadOnlyCollection<(int StartLine, int EndLine)>? changedRanges,
        IReadOnlyCollection<int>? deletedLines,
        bool includeBestPracticeWarnings,
        DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if ((changedRanges?.Count > 0) || (deletedLines?.Count > 0))
        {
            IWpfTextView? view = TryGetActiveWpfTextView();
            if (view is not null)
            {
                BridgeEditHighlightService.Instance.ApplyHighlights(view, changedRanges ?? [], deletedLines ?? []);
            }
        }

        // Best-practice analysis is opt-in to keep the response lean.
        // Pass includeBestPracticeWarnings: true only when the caller explicitly wants them.
        IReadOnlyList<JObject> preWriteWarnings = includeBestPracticeWarnings
            ? ErrorListService.AnalyzeContentBeforeWrite(normalizedPath, content)
            : [];
        string? projectUniqueName = includeBestPracticeWarnings
            ? (TryGetDocumentProjectUniqueName(window.Document) ?? SolutionFileLocator.TryFindProjectUniqueName(dte, normalizedPath))
            : null;

        return new JObject
        {
            [ResolvedPathProperty] = normalizedPath,
            ["editorBacked"] = usedEditorBuffer,
            ["verified"] = true,
            ["contentChanged"] = contentChanged,
            ["saved"] = (window.Document ?? TryFindOpenDocumentByPath(dte, normalizedPath))?.Saved ?? saveChanges,
            ["line"] = Math.Max(1, line),
            ["column"] = Math.Max(1, column),
            ["windowCaption"] = window.Caption,
            ["bestPracticeWarnings"] = preWriteWarnings.Count > 0
                ? BestPracticeWarningProjector.CreateResponseWarnings(preWriteWarnings, projectUniqueName)
                : null,
        };
    }

    private static string ReadFileText(string path)
    {
        return File.Exists(path) ? File.ReadAllText(path) : string.Empty;
    }

    private static bool WriteContentMatches(string written, string readBack)
        => string.Equals(
            written.Replace("\r\n", "\n"),
            readBack.Replace("\r\n", "\n"),
            StringComparison.Ordinal);

    private static string? TryReadDocumentText(Document? document)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (document is null)
        {
            return null;
        }

        try
        {
            if (document.Object(TextDocumentKind) is not TextDocument textDocument)
            {
                return null;
            }

            EditPoint start = textDocument.StartPoint.CreateEditPoint();
            return start.GetText(textDocument.EndPoint);
        }
        catch (COMException)
        {
            return null;
        }
        catch (Exception ex) when (IsDeferredDocumentLoadFailure(ex))
        {
            return null;
        }
    }

    private IWpfTextView? TryGetActiveWpfTextView()
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (_serviceProvider.GetService(typeof(SVsTextManager)) is not IVsTextManager textManager)
        {
            return null;
        }

        if (textManager.GetActiveView(1, null, out IVsTextView? activeView) != 0 || activeView is null)
        {
            return null;
        }

        IComponentModel? componentModel = _serviceProvider.GetService(typeof(SComponentModel)) as IComponentModel;
        IVsEditorAdaptersFactoryService? adapters = componentModel?.GetService<IVsEditorAdaptersFactoryService>();
        return adapters?.GetWpfTextView(activeView);
    }

    public async Task<JObject> SaveDocumentAsync(DTE2 dte, string? filePath, bool saveAll)
    {
        (bool saveAllResult, int count, string? path, bool? saved) = await SaveDocumentOnMainThreadAsync(dte, filePath, saveAll).ConfigureAwait(false);

        return new JObject
        {
            ["saveAll"] = saveAllResult,
            ["count"] = count,
            ["path"] = path,
            ["saved"] = saved,
        };
    }

    public async Task<JObject> ReloadDocumentAsync(string filePath)
    {
        (bool reloaded, string normalizedPath, string? reason) = await ReloadDocumentOnMainThreadAsync(filePath).ConfigureAwait(false);
        JObject result = new()
        {
            ["reloaded"] = reloaded,
            ["path"] = normalizedPath,
        };
        if (reason is not null)
        {
            result["reason"] = reason;
        }

        return result;
    }

    private async Task<(bool SaveAll, int Count, string? Path, bool? Saved)> SaveDocumentOnMainThreadAsync(DTE2 dte, string? filePath, bool saveAll)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        if (saveAll)
        {
            dte.Documents.SaveAll();
            int count = dte.Documents.Count;
            await Task.Yield();
            return (true, count, null, null);
        }

        if (!string.IsNullOrWhiteSpace(filePath))
        {
            string normalized = ResolveDocumentPath(dte, filePath, allowDiskFallback: false);
            Document document = TryFindOpenDocumentByPath(dte, normalized) ?? throw new CommandErrorException(DocumentNotFoundCode, $"No open document matching path: {filePath}");
            document.Save();
            string? path = TryGetDocumentFullName(document);
            bool? saved = TryGetDocumentSaved(document);
            await Task.Yield();
            return (false, 1, path, saved);
        }

        Document active = dte.ActiveDocument ?? throw new CommandErrorException("no_active_document", "No active document to save.");
        active.Save();
        string? activePath = TryGetDocumentFullName(active);
        bool? activeSaved = TryGetDocumentSaved(active);
        await Task.Yield();
        return (false, 1, activePath, activeSaved);
    }

    private async Task<(bool Reloaded, string Path, string? Reason)> ReloadDocumentOnMainThreadAsync(string filePath)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        string normalized = PathNormalization.NormalizeFilePath(filePath);
        IVsRunningDocumentTable rdt = _serviceProvider.GetService(typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable
            ?? throw new CommandErrorException("service_unavailable", "Running Document Table not available.");

        int hr = rdt.FindAndLockDocument(
            (uint)_VSRDTFLAGS.RDT_NoLock,
            normalized,
            out _,
            out _,
            out IntPtr docDataPtr,
            out uint cookie);

        if (hr != VSConstants.S_OK || cookie == 0 || docDataPtr == IntPtr.Zero)
        {
            await Task.Yield();
            return (false, normalized, "not_open");
        }

        try
        {
            if (Marshal.GetObjectForIUnknown(docDataPtr) is not IVsPersistDocData docData)
            {
                await Task.Yield();
                return (false, normalized, "not_editable");
            }

            ErrorHandler.ThrowOnFailure(docData.ReloadDocData(0));
            await Task.Yield();
            return (true, normalized, null);
        }
        finally
        {
            Marshal.Release(docDataPtr);
        }
    }

    private static string ReadDocumentText(DTE2 dte, string resolvedPath)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        Document? openDocument = TryFindOpenDocumentByPath(dte, resolvedPath);
        string? editorText = TryReadDocumentText(openDocument);
        return editorText ?? File.ReadAllText(resolvedPath);
    }

    private static bool HasDocumentPath(Document document)
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        return !string.IsNullOrWhiteSpace(TryGetDocumentFullName(document));
    }

    private static bool MatchesDocumentExactPath(Document document, string normalizedQueryPath)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        string? fullName = TryGetDocumentFullName(document);
        return !string.IsNullOrWhiteSpace(fullName) &&
            string.Equals(fullName, normalizedQueryPath, StringComparison.OrdinalIgnoreCase);
    }

    private static bool MatchesDocumentExactName(Document document, string rawQuery)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        string? name = TryGetDocumentName(document);
        string? fullName = TryGetDocumentFullName(document);
        string fileName = string.IsNullOrWhiteSpace(fullName) ? string.Empty : Path.GetFileName(fullName);
        return string.Equals(name, rawQuery, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(fileName, rawQuery, StringComparison.OrdinalIgnoreCase);
    }

    private static bool MatchesDocumentNameContains(Document document, string rawQuery)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        string? name = TryGetDocumentName(document);
        string? fullName = TryGetDocumentFullName(document);
        string? fileName = string.IsNullOrWhiteSpace(fullName) ? null : Path.GetFileName(fullName);
        return ContainsOrdinalIgnoreCase(name, rawQuery) ||
            ContainsOrdinalIgnoreCase(fileName, rawQuery);
    }

    private static bool MatchesDocumentPathContains(Document document, string rawQuery)
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        return ContainsOrdinalIgnoreCase(TryGetDocumentFullName(document), rawQuery);
    }

    private static IReadOnlyList<string> GetDistinctDocumentNames(IEnumerable<Document> matches)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        return
        [
            .. matches
                .Select(GetDocumentLabel)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase),
        ];
    }

    private static string GetDocumentLabel(Document document)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        string fullName = TryGetDocumentFullName(document) ?? string.Empty;
        string? project = TryGetDocumentProjectUniqueName(document);
        if (!string.IsNullOrWhiteSpace(fullName) && !string.IsNullOrWhiteSpace(project))
        {
            return $"{Path.GetFileName(fullName)} [{project}]";
        }

        return TryGetDocumentName(document) ?? Path.GetFileName(fullName);
    }

    private static bool IsProjectBackedDocument(Document document)
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        return !string.IsNullOrWhiteSpace(TryGetDocumentProjectUniqueName(document));
    }

    private static bool IsReviewArtifactDocument(Document document)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        string fullName = TryGetDocumentFullName(document) ?? string.Empty;
        return fullName.IndexOf("\\output\\pr-review\\", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static bool IsOutputDocument(Document document)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        string fullName = TryGetDocumentFullName(document) ?? string.Empty;
        return fullName.IndexOf("\\output\\", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static bool ContainsOrdinalIgnoreCase(string? value, string query)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return value!.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static List<string> SplitLines(string text)
    {
        string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        return [.. normalized.Split('\n')];
    }

    private static JObject CreateDocumentInfo(Document document, string? activePath, int? tabIndex = null)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        string fullName = TryGetDocumentFullName(document) ?? string.Empty;
        string normalizedPath = string.IsNullOrWhiteSpace(fullName) ? string.Empty : PathNormalization.NormalizeFilePath(fullName);
        string normalizedActivePath = string.IsNullOrWhiteSpace(activePath) ? string.Empty : PathNormalization.NormalizeFilePath(activePath);
        bool? saved = TryGetDocumentSaved(document);

        return new JObject
        {
            ["name"] = TryGetDocumentName(document) ?? Path.GetFileName(fullName),
            ["path"] = normalizedPath,
            ["tabIndex"] = tabIndex,
            ["isActive"] = !string.IsNullOrWhiteSpace(normalizedPath) &&
                string.Equals(normalizedPath, normalizedActivePath, StringComparison.OrdinalIgnoreCase),
            ["project"] = TryGetDocumentProjectUniqueName(document) ?? string.Empty,
            ["isProjectBacked"] = IsProjectBackedDocument(document),
            ["isReviewArtifact"] = IsReviewArtifactDocument(document),
            ["saved"] = (JToken?)saved ?? JValue.CreateNull(),
        };
    }

    private static string? TryGetDocumentFullName(Document? document)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (document is null)
        {
            return null;
        }

        try
        {
            string fullName = document.FullName;
            return string.IsNullOrWhiteSpace(fullName) ? null : PathNormalization.NormalizeFilePath(fullName);
        }
        catch (COMException)
        {
            return null;
        }
        catch (Exception ex) when (IsDeferredDocumentLoadFailure(ex))
        {
            return null;
        }
    }

    private static string? TryGetDocumentName(Document? document)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (document is null)
        {
            return null;
        }

        try
        {
            return string.IsNullOrWhiteSpace(document.Name) ? null : document.Name;
        }
        catch (COMException)
        {
            return null;
        }
        catch (Exception ex) when (IsDeferredDocumentLoadFailure(ex))
        {
            return null;
        }
    }

    private static string? TryGetDocumentProjectUniqueName(Document? document)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (document is null)
        {
            return null;
        }

        try
        {
            string? project = document.ProjectItem?.ContainingProject?.UniqueName;
            return string.IsNullOrWhiteSpace(project) ? null : project;
        }
        catch (COMException)
        {
            return null;
        }
        catch (Exception ex) when (IsDeferredDocumentLoadFailure(ex))
        {
            return null;
        }
    }

    private static bool? TryGetDocumentSaved(Document? document)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (document is null)
        {
            return null;
        }

        try
        {
            return document.Saved;
        }
        catch (COMException)
        {
            return null;
        }
        catch (Exception ex) when (IsDeferredDocumentLoadFailure(ex))
        {
            return null;
        }
    }

    private static bool IsDeferredDocumentLoadFailure(Exception ex)
    {
        return string.Equals(ex.GetType().FullName, "Microsoft.Assumes+InternalErrorException", StringComparison.Ordinal)
            || string.Equals(ex.GetType().Name, "InternalErrorException", StringComparison.Ordinal);
    }
}
