using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed partial class DocumentService
{
    public async Task<JObject> OpenDocumentAsync(DTE2 dte, string filePath, int line, int column, bool allowDiskFallback = true)
    {
        (string normalizedPath, bool navigated, string windowCaption) =
            await OpenDocumentOnMainThreadAsync(dte, filePath, line, column, allowDiskFallback).ConfigureAwait(false);

        return new JObject
        {
            [ResolvedPathProperty] = normalizedPath,
            ["name"] = Path.GetFileName(normalizedPath),
            ["line"] = Math.Max(1, line),
            ["column"] = Math.Max(1, column),
            ["navigated"] = navigated,
            ["windowCaption"] = windowCaption,
        };
    }

    public async Task<JObject> RevealDocumentRangeAsync(
        DTE2 dte,
        string filePath,
        int startLine,
        int startColumn,
        int endLine,
        int endColumn)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        string normalizedPath = PathNormalization.NormalizeFilePath(filePath);
        if (!File.Exists(normalizedPath))
        {
            throw new CommandErrorException(DocumentNotFoundCode, $"File not found: {normalizedPath}. Call find_files to locate the correct path, then retry.");
        }

        Window window = dte.ItemOperations.OpenFile(normalizedPath);
        window.Activate();

        bool navigated = false;
        if (window.Document?.Object(TextDocumentKind) is TextDocument textDocument)
        {
            navigated = TryNavigateToRange(textDocument, startLine, startColumn, endLine, endColumn);
            if (!navigated)
            {
                navigated = await TryNavigateToLineAsync(textDocument, startLine, startColumn).ConfigureAwait(true);
            }
        }

        return new JObject
        {
            [ResolvedPathProperty] = normalizedPath,
            ["name"] = Path.GetFileName(normalizedPath),
            ["startLine"] = Math.Max(1, startLine),
            ["startColumn"] = Math.Max(1, startColumn),
            ["endLine"] = Math.Max(1, endLine),
            ["endColumn"] = Math.Max(1, endColumn),
            ["navigated"] = navigated,
            ["windowCaption"] = window.Caption,
        };
    }

    public async Task<JObject> PositionTextSelectionAsync(
        DTE2 dte,
        string? filePath,
        string? documentQuery,
        int? line,
        int? column,
        bool selectWord)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        if (!string.IsNullOrWhiteSpace(filePath) && !string.IsNullOrWhiteSpace(documentQuery))
        {
            throw new CommandErrorException("invalid_arguments", "Use either --file or --document, not both.");
        }

        Document document = ResolveDocumentForTextSelection(dte, filePath, documentQuery);

        if (document.Object(TextDocumentKind) is not TextDocument textDocument)
        {
            throw new CommandErrorException(UnsupportedOperationCode, $"Document is not a text document: {document.FullName}. Only source code files support cursor positioning — open a .cs, .cpp, .py, or other text-based file.");
        }

        TextSelection selection = textDocument.Selection;
        int targetLine = Math.Max(1, line ?? selection.ActivePoint.Line);
        int targetColumn = Math.Max(1, column ?? selection.ActivePoint.DisplayColumn);
        selection.MoveToLineAndOffset(targetLine, targetColumn, false);
        TryShowActivePoint(selection);
        if (selectWord)
        {
            if (!TrySelectIdentifierAtLocation(textDocument, selection, targetLine, targetColumn))
            {
                TrySelectCurrentWord(selection);
            }
            TryShowActivePoint(selection);
        }

        int activeLine = selection.ActivePoint.Line;
        int activeColumn = selection.ActivePoint.DisplayColumn;
        return new JObject
        {
            [ResolvedPathProperty] = PathNormalization.NormalizeFilePath(document.FullName),
            ["name"] = document.Name,
            ["line"] = activeLine,
            ["column"] = activeColumn,
            [SelectedTextProperty] = selection.Text,
            ["lineText"] = GetLineText(textDocument, activeLine),
        };
    }

    public async Task<JObject> ActivateOpenDocumentAsync(DTE2 dte, string query)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        return ActivateOpenDocumentOnMainThread(dte, query);
    }

    private static JObject ActivateOpenDocumentOnMainThread(DTE2 dte, string query)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        (List<Document> documents, string matchedBy) = ResolveDocumentMatches(dte, query, fallbackToActive: false, allowMultiple: false);
        documents[0].Activate();

        return new JObject
        {
            ["query"] = query,
            ["matchedBy"] = matchedBy,
            ["document"] = TryCreateDocumentInfo(documents[0], TryGetDocumentFullName(documents[0]))
                ?? new JObject
                {
                    ["name"] = TryGetDocumentName(documents[0]) ?? string.Empty,
                    ["path"] = TryGetDocumentFullName(documents[0]) ?? string.Empty,
                    ["tabIndex"] = JValue.CreateNull(),
                    ["isActive"] = true,
                    ["project"] = TryGetDocumentProjectUniqueName(documents[0]) ?? string.Empty,
                    ["isProjectBacked"] = IsProjectBackedDocument(documents[0]),
                    ["isReviewArtifact"] = IsReviewArtifactDocument(documents[0]),
                    ["saved"] = JValue.CreateNull(),
                },
        };
    }

    public async Task<JObject> CloseOpenDocumentsAsync(DTE2 dte, string? query, bool closeAllMatches, bool saveChanges)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        return CloseOpenDocumentsOnMainThread(dte, query, closeAllMatches, saveChanges);
    }

    private static JObject CloseOpenDocumentsOnMainThread(DTE2 dte, string? query, bool closeAllMatches, bool saveChanges)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        string? activePath = TryGetDocumentFullName(dte.ActiveDocument);
        (List<Document> documents, string matchedBy) = ResolveDocumentMatches(dte, query, fallbackToActive: true, allowMultiple: closeAllMatches);
        JArray closed = [];
        JArray skipped = [];
        foreach (Document document in documents)
        {
            JObject info = CreateDocumentInfoOrFallback(document, activePath);
            TryCloseDocument(document, saveChanges, info, closed, skipped);
        }

        return new JObject
        {
            ["query"] = query ?? string.Empty,
            ["matchedBy"] = matchedBy,
            ["saveChanges"] = saveChanges,
            ["count"] = closed.Count,
            ["items"] = closed,
            ["skippedCount"] = skipped.Count,
            ["skipped"] = skipped,
        };
    }

    public async Task<JObject> CloseFileAsync(DTE2 dte, string? filePath, string? query, bool saveChanges)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        string resolvedQuery;
        if (!string.IsNullOrWhiteSpace(filePath))
        {
            resolvedQuery = PathNormalization.NormalizeFilePath(filePath);
        }
        else if (!string.IsNullOrWhiteSpace(query))
        {
            resolvedQuery = query!;
        }
        else
        {
            throw new CommandErrorException("invalid_arguments", "Either file or query is required. Pass file with an absolute path to the document on disk, or query with a text string to match against open document names.");
        }

        return await CloseOpenDocumentsAsync(dte, resolvedQuery, closeAllMatches: false, saveChanges).ConfigureAwait(true);
    }

    public async Task<JObject> CloseAllExceptCurrentAsync(DTE2 dte, bool saveChanges)
    {
        (string activePath, JArray closed, JArray skipped) = await CloseAllExceptCurrentOnMainThreadAsync(dte, saveChanges).ConfigureAwait(false);

        return new JObject
        {
            ["activePath"] = PathNormalization.NormalizeFilePath(activePath),
            ["saveChanges"] = saveChanges,
            ["count"] = closed.Count,
            ["items"] = closed,
            ["skippedCount"] = skipped.Count,
            ["skipped"] = skipped,
        };
    }

    private static JObject CreateDocumentInfoOrFallback(Document document, string? activePath)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        return TryCreateDocumentInfo(document, activePath)
            ?? new JObject
            {
                ["name"] = TryGetDocumentName(document) ?? string.Empty,
                ["path"] = TryGetDocumentFullName(document) ?? string.Empty,
                ["tabIndex"] = JValue.CreateNull(),
                ["isActive"] = false,
                ["project"] = TryGetDocumentProjectUniqueName(document) ?? string.Empty,
                ["isProjectBacked"] = IsProjectBackedDocument(document),
                ["isReviewArtifact"] = IsReviewArtifactDocument(document),
                ["saved"] = JValue.CreateNull(),
            };
    }

    private static void TryCloseDocument(Document document, bool saveChanges, JObject info, JArray closed, JArray skipped)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        vsSaveChanges saveMode = saveChanges ? vsSaveChanges.vsSaveChangesYes : vsSaveChanges.vsSaveChangesNo;
        try
        {
            CloseDocumentWithWindowFallback(document, saveMode);
            closed.Add(info);
        }
        catch (COMException ex)
        {
            skipped.Add(CreateSkippedCloseInfo(info, ex));
        }
        catch (Exception ex) when (IsDeferredDocumentLoadFailure(ex))
        {
            skipped.Add(CreateSkippedCloseInfo(info, ex));
        }
    }

    private static void CloseDocumentWithWindowFallback(Document document, vsSaveChanges saveMode)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            document.Close(saveMode);
        }
        catch (COMException) when (TryCloseDocumentWindow(document, saveMode))
        {
        }
        catch (Exception ex) when (IsDeferredDocumentLoadFailure(ex) && TryCloseDocumentWindow(document, saveMode))
        {
        }
    }

    private static bool TryCloseDocumentWindow(Document document, vsSaveChanges saveMode)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            foreach (Window window in document.Windows)
            {
                window.Close(saveMode);
                return true;
            }
        }
        catch (COMException)
        {
            return false;
        }
        catch (Exception ex) when (IsDeferredDocumentLoadFailure(ex))
        {
            return false;
        }

        return false;
    }

    private static JObject CreateSkippedCloseInfo(JObject info, Exception ex)
    {
        JObject skipped = (JObject)info.DeepClone();
        skipped["closeError"] = ex.Message;
        skipped["closeErrorType"] = ex.GetType().FullName ?? ex.GetType().Name;
        return skipped;
    }

    private async Task<(string NormalizedPath, bool Navigated, string WindowCaption)> OpenDocumentOnMainThreadAsync(
        DTE2 dte,
        string filePath,
        int line,
        int column,
        bool allowDiskFallback)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        string normalizedPath = ResolveDocumentPath(dte, filePath, allowDiskFallback);
        if (!File.Exists(normalizedPath))
        {
            normalizedPath = TryFindExistingOpenDocumentPathByFileName(dte, normalizedPath) ?? normalizedPath;
            if (!File.Exists(normalizedPath))
            {
                throw new CommandErrorException(
                    DocumentNotFoundCode,
                    $"File does not exist on disk: {normalizedPath}");
            }
        }

        Window window = dte.ItemOperations.OpenFile(normalizedPath);
        window.Activate();

        bool navigated = false;
        if (window.Document?.Object(TextDocumentKind) is TextDocument textDocument)
        {
            navigated = await TryNavigateToLineAsync(textDocument, line, column).ConfigureAwait(true);
        }

        string windowCaption = window.Caption;
        await Task.Yield();
        return (normalizedPath, navigated, windowCaption);
    }

    private static async Task<(string ActivePath, JArray Closed, JArray Skipped)> CloseAllExceptCurrentOnMainThreadAsync(DTE2 dte, bool saveChanges)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        Document activeDocument = dte.ActiveDocument ?? throw new CommandErrorException(DocumentNotFoundCode, "There is no active document. Open a file in VS first using open_file.");
        string? activePath = TryGetDocumentFullName(activeDocument);
        if (string.IsNullOrWhiteSpace(activePath))
        {
            throw new CommandErrorException(DocumentNotFoundCode, "The active document does not have a file path. It may be an unsaved scratch buffer — open a file from disk first using open_file.");
        }

        List<Document> documentsToClose = [..EnumerateOpenDocuments(dte)
            .Where(document => !string.Equals(
                TryGetDocumentFullName(document),
                activePath,
                StringComparison.OrdinalIgnoreCase))];

        JArray closed = [];
        JArray skipped = [];
        foreach (Document document in documentsToClose)
        {
            JObject info = CreateDocumentInfoOrFallback(document, activePath);
            TryCloseDocument(document, saveChanges, info, closed, skipped);
        }

        await Task.Yield();
        return (activePath!, closed, skipped);
    }

    public async Task<(string Path, string Text)> GetActiveDocumentTextAsync(DTE2 dte)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        Document document = dte.ActiveDocument ?? throw new CommandErrorException(DocumentNotFoundCode, "There is no active document. Open a file in VS first using open_file.");
        if (document.Object(TextDocumentKind) is not TextDocument textDocument)
        {
            throw new CommandErrorException(UnsupportedOperationCode, $"Active document is not a text document: {document.FullName}. Only source code files support text operations — open a .cs, .cpp, .py, or other text-based file.");
        }

        EditPoint editPoint = textDocument.StartPoint.CreateEditPoint();
        return (document.FullName, editPoint.GetText(textDocument.EndPoint));
    }

    // ── Extracted helpers ─────────────────────────────────────────────────────

    private Document ResolveDocumentForTextSelection(DTE2 dte, string? filePath, string? documentQuery)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (!string.IsNullOrWhiteSpace(filePath))
        {
            // Use the full resolver (handles, solution items, ancestor search roots). Naive
            // normalization resolves relative paths against the solution directory, which for
            // out-of-tree builds is the build tree — "src/foo.cpp" must find the project-backed
            // source file, not fail on a nonexistent build-dir copy.
            string normalizedPath = ResolveDocumentPath(dte, filePath);
            if (!File.Exists(normalizedPath))
            {
                throw new CommandErrorException(DocumentNotFoundCode, $"File not found: {normalizedPath}. Call find_files to locate the correct path, then retry.");
            }

            Window window = dte.ItemOperations.OpenFile(normalizedPath);
            window.Activate();
            return window.Document ?? dte.ActiveDocument
                ?? throw new CommandErrorException(DocumentNotFoundCode, $"Unable to activate document: {normalizedPath}. Verify the file exists and is not locked by another process.");
        }

        if (!string.IsNullOrWhiteSpace(documentQuery))
        {
            (List<Document> documents, _) = ResolveDocumentMatches(dte, documentQuery, fallbackToActive: false, allowMultiple: false);
            Document resolved = documents[0];
            resolved.Activate();
            return resolved;
        }

        Document active = dte.ActiveDocument ?? throw new CommandErrorException(DocumentNotFoundCode, "There is no active document. Open a file in VS first using open_file.");
        active.Activate();
        return active;
    }

    private static bool TryNavigateToRange(
        TextDocument textDocument,
        int startLine,
        int startColumn,
        int endLine,
        int endColumn)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            TextSelection selection = textDocument.Selection;
            SelectRange(
                textDocument,
                selection,
                Math.Max(1, startLine),
                Math.Max(1, startColumn),
                Math.Max(1, endLine),
                Math.Max(1, endColumn));
            TryShowActivePoint(selection);
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (COMException)
        {
            return false;
        }
    }


    // ── Low-level navigation helpers ──────────────────────────────────────────

    private static void TrySelectCurrentWord(TextSelection selection)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        int originalLine = selection.ActivePoint.Line;
        int originalColumn = selection.ActivePoint.DisplayColumn;
        selection.WordLeft(false, 1);
        if (selection.ActivePoint.Line != originalLine)
        {
            selection.MoveToLineAndOffset(originalLine, originalColumn, false);
        }

        selection.WordRight(true, 1);
        if (string.IsNullOrWhiteSpace(selection.Text))
        {
            selection.MoveToLineAndOffset(originalLine, originalColumn, false);
        }
    }

    private static bool TrySelectIdentifierAtLocation(TextDocument textDocument, TextSelection selection, int lineNumber, int column)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        string lineText = GetLineText(textDocument, lineNumber);
        if (string.IsNullOrWhiteSpace(lineText))
        {
            return false;
        }

        Match[] matches = [..Regex.Matches(lineText, @"\b[_\p{L}][_\p{L}\p{Nd}]*\b")
            .Cast<Match>()];
        if (matches.Length == 0)
        {
            return false;
        }

        int zeroBasedColumn = Math.Max(0, column - 1);
        Match? targetMatch = matches.FirstOrDefault(match => zeroBasedColumn >= match.Index && zeroBasedColumn < match.Index + match.Length)
            ?? matches.FirstOrDefault(match => match.Index >= zeroBasedColumn)
            ?? matches.LastOrDefault(match => match.Index < zeroBasedColumn);
        if (targetMatch is null)
        {
            return false;
        }

        int startColumn = targetMatch.Index + 1;
        int endExclusiveColumn = startColumn + targetMatch.Length;

        selection.MoveToLineAndOffset(lineNumber, startColumn, false);
        selection.MoveToLineAndOffset(lineNumber, endExclusiveColumn, true);
        return HasNavigableSymbolText(selection.Text);
    }

    private static bool HasNavigableSymbolText(string? text)
    {
        return !string.IsNullOrWhiteSpace(text)
            && text.Any(static character => char.IsLetterOrDigit(character) || character == '_');
    }

    /// <summary>
    /// Navigates to a line/column, clamping to the document range and handling
    /// transient COM errors (e.g. "Class not registered" when the editor buffer
    /// hasn't fully initialized after first open).
    /// </summary>
    private static async Task<bool> TryNavigateToLineAsync(TextDocument textDocument, int line, int column)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        if (TryNavigateToLineOnce(textDocument, line, column)) return true;

        // Editor buffers can be transiently unavailable immediately after open.
        // Yield the UI thread during the retry delay so the IDE stays responsive.
        await Task.Delay(100).ConfigureAwait(true);
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        if (TryNavigateToLineOnce(textDocument, line, column)) return true;

        // Final fallback: navigate to line 1 col 1 so the file is still visible.
        TryNavigateToLineFallback(textDocument);
        return false;
    }

    private static bool TryNavigateToLineOnce(TextDocument textDocument, int line, int column)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            int maxLine = textDocument.EndPoint.Line;
            int clampedLine = Math.Min(Math.Max(line, 1), Math.Max(1, maxLine));
            TextSelection selection = textDocument.Selection;
            selection.MoveToLineAndOffset(clampedLine, Math.Max(1, column), false);
            TryShowActivePoint(selection);
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (COMException)
        {
            return false;
        }
    }

    private static void TryNavigateToLineFallback(TextDocument textDocument)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            TextSelection selection = textDocument.Selection;
            selection.MoveToLineAndOffset(1, 1, false);
            TryShowActivePoint(selection);
        }
        catch (COMException ex)
        {
            BridgeActivityLog.LogVerbose(nameof(DocumentService), "Editor selection fallback could not position the caret (document not ready or not a text view)", ex);
        }
    }

    private static void TryShowActivePoint(TextSelection selection)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            selection.ActivePoint.TryToShow(vsPaneShowHow.vsPaneShowCentered);
        }
        catch (COMException ex)
        {
            // Viewport centering is best-effort; transient COM failures (E_UNEXPECTED etc.)
            // do not affect the navigation result. Log at verbose only — not to the log file.
            BridgeActivityLog.LogVerbose(nameof(DocumentService), "Failed to center the active editor point in the viewport", ex);
        }
    }

    private static void SelectRange(
        TextDocument textDocument,
        TextSelection selection,
        int startLine,
        int startColumn,
        int endLine,
        int endColumn)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        int resolvedStartLine = Math.Max(1, startLine);
        int resolvedStartColumn = Math.Max(1, startColumn);
        int resolvedEndLine = Math.Max(resolvedStartLine, endLine);
        int resolvedEndColumn = Math.Max(1, endColumn);

        selection.MoveToLineAndOffset(resolvedStartLine, resolvedStartColumn, false);

        if (resolvedEndLine <= resolvedStartLine)
        {
            selection.EndOfLine(true);
            return;
        }

        int documentEndLine = textDocument.EndPoint.Line;
        if (resolvedEndLine >= documentEndLine)
        {
            selection.EndOfDocument(true);
            return;
        }

        selection.MoveToLineAndOffset(resolvedEndLine + 1, resolvedEndColumn, true);
    }

    private static string GetLineText(TextDocument textDocument, int lineNumber)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (lineNumber < 1)
        {
            return string.Empty;
        }

        EditPoint start = textDocument.StartPoint.CreateEditPoint();
        start.MoveToLineAndOffset(lineNumber, 1);
        EditPoint end = start.CreateEditPoint();
        end.EndOfLine();
        return start.GetText(end);
    }
}
