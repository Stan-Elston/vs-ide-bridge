using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed class DocumentService
{
    public async Task<JObject> ListOpenTabsAsync(DTE2 dte)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var activePath = dte.ActiveDocument?.FullName;
        var documents = EnumerateOpenDocuments(dte);
        var items = new JArray();
        for (var i = 0; i < documents.Count; i++)
        {
            items.Add(CreateDocumentInfo(documents[i], activePath, i + 1));
        }

        return new JObject
        {
            ["count"] = items.Count,
            ["activePath"] = string.IsNullOrWhiteSpace(activePath) ? string.Empty : PathNormalization.NormalizeFilePath(activePath),
            ["items"] = items,
        };
    }

    public async Task<JObject> ListOpenDocumentsAsync(DTE2 dte)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var activePath = dte.ActiveDocument?.FullName;
        var items = new JArray(
            EnumerateOpenDocuments(dte)
                .Select((document, index) => CreateDocumentInfo(document, activePath, index + 1)));

        return new JObject
        {
            ["count"] = items.Count,
            ["items"] = items,
        };
    }

    public async Task<JObject> OpenDocumentAsync(DTE2 dte, string filePath, int line, int column)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var normalizedPath = PathNormalization.NormalizeFilePath(filePath);
        if (!File.Exists(normalizedPath))
        {
            throw new CommandErrorException("document_not_found", $"File not found: {normalizedPath}");
        }

        var window = dte.ItemOperations.OpenFile(normalizedPath);
        window.Activate();

        if (window.Document?.Object("TextDocument") is TextDocument textDocument)
        {
            var selection = textDocument.Selection;
            selection.MoveToLineAndOffset(Math.Max(1, line), Math.Max(1, column), false);
            TryShowActivePoint(selection);
        }

        return new JObject
        {
            ["resolvedPath"] = normalizedPath,
            ["name"] = Path.GetFileName(normalizedPath),
            ["line"] = Math.Max(1, line),
            ["column"] = Math.Max(1, column),
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

        Document document;
        if (!string.IsNullOrWhiteSpace(filePath))
        {
            var normalizedPath = PathNormalization.NormalizeFilePath(filePath);
            if (!File.Exists(normalizedPath))
            {
                throw new CommandErrorException("document_not_found", $"File not found: {normalizedPath}");
            }

            var window = dte.ItemOperations.OpenFile(normalizedPath);
            window.Activate();
            document = window.Document ?? dte.ActiveDocument ?? throw new CommandErrorException("document_not_found", $"Unable to activate: {normalizedPath}");
        }
        else if (!string.IsNullOrWhiteSpace(documentQuery))
        {
            var match = ResolveDocumentMatches(dte, documentQuery, fallbackToActive: false, allowMultiple: false);
            document = match.Documents[0];
            document.Activate();
        }
        else
        {
            document = dte.ActiveDocument ?? throw new CommandErrorException("document_not_found", "There is no active document.");
            document.Activate();
        }

        if (document.Object("TextDocument") is not TextDocument textDocument)
        {
            throw new CommandErrorException("unsupported_operation", $"Document is not a text document: {document.FullName}");
        }

        var selection = textDocument.Selection;
        var targetLine = Math.Max(1, line ?? selection.ActivePoint.Line);
        var targetColumn = Math.Max(1, column ?? selection.ActivePoint.DisplayColumn);
        selection.MoveToLineAndOffset(targetLine, targetColumn, false);
        TryShowActivePoint(selection);
        if (selectWord)
        {
            TrySelectCurrentWord(selection);
            TryShowActivePoint(selection);
        }

        var activeLine = selection.ActivePoint.Line;
        var activeColumn = selection.ActivePoint.DisplayColumn;
        return new JObject
        {
            ["resolvedPath"] = PathNormalization.NormalizeFilePath(document.FullName),
            ["name"] = document.Name,
            ["line"] = activeLine,
            ["column"] = activeColumn,
            ["selectedText"] = selection.Text,
            ["lineText"] = GetLineText(textDocument, activeLine),
        };
    }

    public async Task<JObject> ActivateOpenDocumentAsync(DTE2 dte, string query)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var match = ResolveDocumentMatches(dte, query, fallbackToActive: false, allowMultiple: false);
        match.Documents[0].Activate();

        return new JObject
        {
            ["query"] = query,
            ["matchedBy"] = match.MatchedBy,
            ["document"] = CreateDocumentInfo(match.Documents[0], match.Documents[0].FullName),
        };
    }

    public async Task<JObject> CloseOpenDocumentsAsync(DTE2 dte, string? query, bool closeAllMatches, bool saveChanges)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var activePath = dte.ActiveDocument?.FullName;
        var match = ResolveDocumentMatches(dte, query, fallbackToActive: true, allowMultiple: closeAllMatches);
        var closed = new JArray();
        foreach (var document in match.Documents)
        {
            var info = CreateDocumentInfo(document, activePath);
            document.Close(saveChanges ? vsSaveChanges.vsSaveChangesYes : vsSaveChanges.vsSaveChangesNo);
            closed.Add(info);
        }

        return new JObject
        {
            ["query"] = query ?? string.Empty,
            ["matchedBy"] = match.MatchedBy,
            ["saveChanges"] = saveChanges,
            ["count"] = closed.Count,
            ["items"] = closed,
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
            resolvedQuery = query;
        }
        else
        {
            throw new CommandErrorException("invalid_arguments", "Specify --file or --query.");
        }

        return await CloseOpenDocumentsAsync(dte, resolvedQuery, closeAllMatches: false, saveChanges).ConfigureAwait(true);
    }

    public async Task<JObject> CloseAllExceptCurrentAsync(DTE2 dte, bool saveChanges)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var activeDocument = dte.ActiveDocument ?? throw new CommandErrorException("document_not_found", "There is no active document.");
        var activePath = activeDocument.FullName;
        var documentsToClose = EnumerateOpenDocuments(dte)
            .Where(document => !string.Equals(
                PathNormalization.NormalizeFilePath(document.FullName),
                PathNormalization.NormalizeFilePath(activePath),
                StringComparison.OrdinalIgnoreCase))
            .ToList();

        var closed = new JArray();
        foreach (var document in documentsToClose)
        {
            closed.Add(CreateDocumentInfo(document, activePath));
            document.Close(saveChanges ? vsSaveChanges.vsSaveChangesYes : vsSaveChanges.vsSaveChangesNo);
        }

        return new JObject
        {
            ["activePath"] = PathNormalization.NormalizeFilePath(activePath),
            ["saveChanges"] = saveChanges,
            ["count"] = closed.Count,
            ["items"] = closed,
        };
    }

    public async Task<(string Path, string Text)> GetActiveDocumentTextAsync(DTE2 dte)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var document = dte.ActiveDocument;
        if (document is null)
        {
            throw new CommandErrorException("document_not_found", "There is no active document.");
        }

        if (document.Object("TextDocument") is not TextDocument textDocument)
        {
            throw new CommandErrorException("unsupported_operation", $"Active document is not a text document: {document.FullName}");
        }

        var editPoint = textDocument.StartPoint.CreateEditPoint();
        return (document.FullName, editPoint.GetText(textDocument.EndPoint));
    }

    public async Task<JObject> GetDocumentSliceAsync(
        DTE2 dte,
        string? filePath,
        int startLine,
        int endLine,
        bool includeLineNumbers)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var resolvedPath = ResolveDocumentPath(dte, filePath);
        var text = ReadDocumentText(dte, resolvedPath);
        var lines = SplitLines(text);

        var clampedStart = Math.Max(1, startLine);
        var clampedEnd = Math.Max(clampedStart, endLine);
        var actualStart = Math.Min(clampedStart, Math.Max(1, lines.Count));
        var actualEnd = Math.Min(clampedEnd, Math.Max(1, lines.Count));

        var sliceLines = new JArray();
        var builder = new System.Text.StringBuilder();
        for (var lineNumber = actualStart; lineNumber <= actualEnd && lineNumber <= lines.Count; lineNumber++)
        {
            var lineText = lines[lineNumber - 1];
            sliceLines.Add(new JObject
            {
                ["line"] = lineNumber,
                ["text"] = lineText,
            });

            if (builder.Length > 0)
            {
                builder.Append('\n');
            }

            if (includeLineNumbers)
            {
                builder.Append(lineNumber);
                builder.Append(": ");
            }

            builder.Append(lineText);
        }

        return new JObject
        {
            ["resolvedPath"] = resolvedPath,
            ["requestedStartLine"] = clampedStart,
            ["requestedEndLine"] = clampedEnd,
            ["actualStartLine"] = actualStart,
            ["actualEndLine"] = actualEnd,
            ["lineCount"] = lines.Count,
            ["text"] = builder.ToString(),
            ["lines"] = sliceLines,
        };
    }

    public async Task<JObject> GoToDefinitionAsync(
        DTE2 dte,
        string? filePath,
        string? documentQuery,
        int? line,
        int? column)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var sourceLocation = await PositionTextSelectionAsync(dte, filePath, documentQuery, line, column, selectWord: false)
            .ConfigureAwait(true);

        try
        {
            dte.ExecuteCommand("Edit.GoToDefinition");
        }
        catch (COMException ex)
        {
            throw new CommandErrorException("unsupported_operation", $"Edit.GoToDefinition failed: {ex.Message}");
        }

        var activeDoc = dte.ActiveDocument;
        if (activeDoc is null)
        {
            return new JObject
            {
                ["sourceLocation"] = sourceLocation,
                ["definitionLocation"] = null,
                ["definitionFound"] = false,
            };
        }

        var definitionPath = PathNormalization.NormalizeFilePath(activeDoc.FullName);
        int definitionLine = 0, definitionColumn = 0;
        string selectedText = string.Empty, lineText = string.Empty;

        if (activeDoc.Object("TextDocument") is TextDocument defTextDoc)
        {
            var selection = defTextDoc.Selection;
            definitionLine = selection.ActivePoint.Line;
            definitionColumn = selection.ActivePoint.DisplayColumn;
            selectedText = selection.Text ?? string.Empty;
            lineText = GetLineText(defTextDoc, definitionLine);
        }

        var definitionLocation = new JObject
        {
            ["resolvedPath"] = definitionPath,
            ["name"] = activeDoc.Name ?? string.Empty,
            ["line"] = definitionLine,
            ["column"] = definitionColumn,
            ["selectedText"] = selectedText,
            ["lineText"] = lineText,
        };

        var sourcePath = (string?)sourceLocation["resolvedPath"] ?? string.Empty;
        var sourceLine = (int?)sourceLocation["line"] ?? 0;
        var definitionFound = !string.Equals(sourcePath, definitionPath, StringComparison.OrdinalIgnoreCase)
            || sourceLine != definitionLine;

        return new JObject
        {
            ["sourceLocation"] = sourceLocation,
            ["definitionLocation"] = definitionLocation,
            ["definitionFound"] = definitionFound,
        };
    }

    public async Task<JObject> GetFileOutlineAsync(DTE2 dte, string? filePath, int maxDepth)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var resolvedPath = ResolveDocumentPath(dte, filePath);

        ProjectItem? projectItem = null;
        try { projectItem = dte.Solution.FindProjectItem(resolvedPath); } catch { }

        if (projectItem is null)
        {
            return new JObject
            {
                ["resolvedPath"] = resolvedPath,
                ["count"] = 0,
                ["symbols"] = new JArray(),
                ["note"] = "File is not part of any project or code model is unavailable.",
            };
        }

        var symbols = new JArray();
        string? note = null;
        try
        {
            var codeModel = projectItem.FileCodeModel;
            if (codeModel?.CodeElements is not null)
            {
                foreach (CodeElement element in codeModel.CodeElements)
                {
                    try { CollectOutlineSymbols(element, symbols, 0, maxDepth); } catch { }
                }
            }
            else
            {
                note = "No code model available for this file type.";
            }
        }
        catch (Exception ex)
        {
            note = $"Code model unavailable: {ex.Message}";
        }

        var result = new JObject
        {
            ["resolvedPath"] = resolvedPath,
            ["count"] = symbols.Count,
            ["symbols"] = symbols,
        };
        if (note is not null) result["note"] = note;
        return result;
    }

    private static readonly HashSet<vsCMElement> s_outlineKinds = new HashSet<vsCMElement>
    {
        vsCMElement.vsCMElementFunction,
        vsCMElement.vsCMElementClass,
        vsCMElement.vsCMElementStruct,
        vsCMElement.vsCMElementEnum,
        vsCMElement.vsCMElementNamespace,
        vsCMElement.vsCMElementInterface,
    };

    private static void CollectOutlineSymbols(CodeElement element, JArray symbols, int depth, int maxDepth)
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        if (depth > maxDepth) return;

        vsCMElement kind;
        try { kind = element.Kind; } catch { return; }

        string name = string.Empty;
        int startLine = 0, endLine = 0;
        try { name = element.Name ?? string.Empty; } catch { }
        try { startLine = element.StartPoint?.Line ?? 0; } catch { }
        try { endLine = element.EndPoint?.Line ?? 0; } catch { }

        if (s_outlineKinds.Contains(kind))
        {
            symbols.Add(new JObject
            {
                ["name"] = name,
                ["kind"] = kind.ToString().Replace("vsCMElement", string.Empty),
                ["startLine"] = startLine,
                ["endLine"] = endLine,
                ["depth"] = depth,
            });
        }

        CodeElements? children = null;
        try
        {
            if (element is CodeNamespace ns) children = ns.Members;
            else if (element is CodeClass cls) children = cls.Members;
            else if (element is CodeStruct st) children = st.Members;
            else if (element is CodeInterface iface) children = iface.Members;
        }
        catch { }

        if (children is null) return;
        foreach (CodeElement child in children)
        {
            try { CollectOutlineSymbols(child, symbols, depth + 1, maxDepth); } catch { }
        }
    }

    private static string ResolveDocumentPath(DTE2 dte, string? filePath)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (!string.IsNullOrWhiteSpace(filePath))
        {
            var normalizedPath = PathNormalization.NormalizeFilePath(filePath);
            if (!File.Exists(normalizedPath))
            {
                throw new CommandErrorException("document_not_found", $"File not found: {normalizedPath}");
            }

            return normalizedPath;
        }

        if (dte.ActiveDocument is null || string.IsNullOrWhiteSpace(dte.ActiveDocument.FullName))
        {
            throw new CommandErrorException("document_not_found", "There is no active document.");
        }

        return PathNormalization.NormalizeFilePath(dte.ActiveDocument.FullName);
    }

    private static void TrySelectCurrentWord(TextSelection selection)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var originalLine = selection.ActivePoint.Line;
        var originalColumn = selection.ActivePoint.DisplayColumn;
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

    private static void TryShowActivePoint(TextSelection selection)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            selection.ActivePoint.TryToShow(vsPaneShowHow.vsPaneShowCentered);
        }
        catch
        {
            // Some editor surfaces may not support viewport repositioning.
        }
    }

    private static string GetLineText(TextDocument textDocument, int lineNumber)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (lineNumber < 1)
        {
            return string.Empty;
        }

        var start = textDocument.StartPoint.CreateEditPoint();
        start.MoveToLineAndOffset(lineNumber, 1);
        var end = start.CreateEditPoint();
        end.EndOfLine();
        return start.GetText(end);
    }

    private static IReadOnlyList<Document> EnumerateOpenDocuments(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        return dte.Documents.Cast<Document>().Where(document => !string.IsNullOrWhiteSpace(document.FullName)).ToArray();
    }

    private static (List<Document> Documents, string MatchedBy) ResolveDocumentMatches(
        DTE2 dte,
        string? query,
        bool fallbackToActive,
        bool allowMultiple)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var documents = EnumerateOpenDocuments(dte);
        if (documents.Count == 0)
        {
            throw new CommandErrorException("document_not_found", "There are no open documents.");
        }

        if (string.IsNullOrWhiteSpace(query))
        {
            if (!fallbackToActive || dte.ActiveDocument is null)
            {
                throw new CommandErrorException("invalid_arguments", "Missing document query.");
            }

            return (new List<Document> { dte.ActiveDocument }, "active");
        }

        var rawQuery = query.Trim();
        var queryLooksLikePath = rawQuery.IndexOfAny(new[] { '\\', '/', ':' }) >= 0;
        if (queryLooksLikePath)
        {
            var normalizedQueryPath = PathNormalization.NormalizeFilePath(rawQuery);
            var exactPath = documents.Where(document => string.Equals(
                    PathNormalization.NormalizeFilePath(document.FullName),
                    normalizedQueryPath,
                    StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (exactPath.Count > 0)
            {
                return FinalizeMatches(exactPath, allowMultiple, "path");
            }
        }

        var exactName = documents.Where(document =>
                string.Equals(document.Name, rawQuery, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(Path.GetFileName(document.FullName), rawQuery, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (exactName.Count > 0)
        {
            return FinalizeMatches(exactName, allowMultiple, "filename");
        }

        var containsName = documents.Where(document =>
                document.Name.IndexOf(rawQuery, StringComparison.OrdinalIgnoreCase) >= 0 ||
                Path.GetFileName(document.FullName).IndexOf(rawQuery, StringComparison.OrdinalIgnoreCase) >= 0)
            .ToList();
        if (containsName.Count > 0)
        {
            return FinalizeMatches(containsName, allowMultiple, "filename-contains");
        }

        var containsPath = documents.Where(document =>
                document.FullName.IndexOf(rawQuery, StringComparison.OrdinalIgnoreCase) >= 0)
            .ToList();
        if (containsPath.Count > 0)
        {
            return FinalizeMatches(containsPath, allowMultiple, "path-contains");
        }

        throw new CommandErrorException("document_not_found", $"No open document matched '{rawQuery}'.");
    }

    private static (List<Document> Documents, string MatchedBy) FinalizeMatches(List<Document> matches, bool allowMultiple, string matchedBy)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (!allowMultiple && matches.Count > 1)
        {
            var options = string.Join(", ", matches.Select(document => document.Name).Distinct(StringComparer.OrdinalIgnoreCase));
            throw new CommandErrorException("invalid_arguments", $"Document query is ambiguous. Matches: {options}");
        }

        return allowMultiple ? (matches, matchedBy) : (new List<Document> { matches[0] }, matchedBy);
    }

    private static string ReadDocumentText(DTE2 dte, string resolvedPath)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var openDocument = dte.Documents
            .Cast<Document>()
            .FirstOrDefault(document => string.Equals(
                PathNormalization.NormalizeFilePath(document.FullName),
                resolvedPath,
                StringComparison.OrdinalIgnoreCase));

        if (openDocument?.Object("TextDocument") is TextDocument textDocument)
        {
            var editPoint = textDocument.StartPoint.CreateEditPoint();
            return editPoint.GetText(textDocument.EndPoint);
        }

        return File.ReadAllText(resolvedPath);
    }

    private static List<string> SplitLines(string text)
    {
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        return normalized.Split('\n').ToList();
    }

    private static JObject CreateDocumentInfo(Document document, string? activePath, int? tabIndex = null)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var fullName = document.FullName ?? string.Empty;
        var normalizedPath = string.IsNullOrWhiteSpace(fullName) ? string.Empty : PathNormalization.NormalizeFilePath(fullName);
        var normalizedActivePath = string.IsNullOrWhiteSpace(activePath) ? string.Empty : PathNormalization.NormalizeFilePath(activePath);

        return new JObject
        {
            ["name"] = document.Name ?? Path.GetFileName(fullName),
            ["path"] = normalizedPath,
            ["tabIndex"] = tabIndex,
            ["isActive"] = !string.IsNullOrWhiteSpace(normalizedPath) &&
                string.Equals(normalizedPath, normalizedActivePath, StringComparison.OrdinalIgnoreCase),
            ["saved"] = document.Saved,
        };
    }
}
