using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed class OutputWindowService
{
    private const int DefaultChunkLines = 200;
    private const int DefaultMaxChars = 120_000;
    private const int OutputReadAttemptCount = 2;
    private const string BuildPaneName = "Build";

    /// <summary>
    /// Best-effort read of the Build pane's full text for post-build error parsing.
    /// Returns null when the pane does not exist yet (no build has run) or reading fails.
    /// </summary>
    public async Task<string?> TryReadBuildPaneTextAsync(DTE2 dte)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        try
        {
            OutputWindowPane pane = dte.ToolWindows.OutputWindow.OutputWindowPanes.Item(BuildPaneName);
            return ReadOutputPaneText(dte, pane, activate: false);
        }
        catch (Exception ex) when (ex is not null) // pane missing before the first build, or COM read failure
        {
            Debug.WriteLine(ex);
            return null;
        }
    }

    public async Task<JObject> ReadOutputWindowAsync(
        DTE2 dte,
        string? requestedPane,
        int? chunkLines,
        int? chunkIndex,
        int? maxChars,
        bool includeChunks,
        bool activate)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        int effectiveChunkLines = chunkLines ?? DefaultChunkLines;
        int effectiveMaxChars = maxChars ?? DefaultMaxChars;
        ValidateLimits(effectiveChunkLines, chunkIndex, effectiveMaxChars);

        OutputWindow outputWindow = dte.ToolWindows.OutputWindow;
        IReadOnlyList<OutputPaneSnapshot> panes = CaptureOutputPanes(outputWindow);
        if (panes.Count == 0)
        {
            throw new CommandErrorException("output_pane_not_found", "Visual Studio has no Output window panes to read. Trigger a build, run, or debug session first to populate the Output window, then retry.");
        }

        OutputWindowPane pane = SelectPane(outputWindow, panes, requestedPane);
        OutputPaneSnapshot selectedPane = CreatePaneSnapshot(pane, isActive: IsActivePane(outputWindow, pane));

        string fullText = ReadOutputPaneText(dte, pane, activate);
        OutputTextSlice slice = SliceOutputText(fullText, effectiveChunkLines, chunkIndex, effectiveMaxChars);

        return new JObject
        {
            ["pane"] = selectedPane.Name,
            ["paneGuid"] = selectedPane.Guid,
            ["requestedPane"] = string.IsNullOrWhiteSpace(requestedPane) ? null : requestedPane,
            ["usedDefaultPane"] = string.IsNullOrWhiteSpace(requestedPane),
            ["availablePanes"] = CreatePanesArray(panes, selectedPane),
            ["lineCount"] = slice.TotalLineCount,
            ["returnedLineCount"] = slice.ReturnedLineCount,
            ["totalChars"] = fullText.Length,
            ["returnedChars"] = slice.Text.Length,
            ["tailLines"] = effectiveChunkLines,
            ["chunkLines"] = effectiveChunkLines,
            ["chunkCount"] = slice.Chunks.Count,
            ["selectedChunkIndex"] = slice.SelectedChunk.Index,
            ["selectedChunkNumber"] = slice.SelectedChunk.Index + 1,
            ["maxChars"] = effectiveMaxChars,
            ["includeChunks"] = includeChunks,
            ["lineTruncated"] = slice.LineTruncated,
            ["charTruncated"] = slice.CharTruncated,
            ["truncated"] = slice.LineTruncated || slice.CharTruncated,
            ["activated"] = activate,
            ["chunks"] = CreateChunksArray(slice.Chunks, slice.SelectedChunk.Index, includeChunks),
            ["selectedChunk"] = CreateChunkObject(slice.SelectedChunk, includeText: true, selected: true, textOverride: slice.Text, charTruncated: slice.CharTruncated),
            ["text"] = slice.Text,
        };
    }

    private static void ValidateLimits(int chunkLines, int? chunkIndex, int maxChars)
    {
        if (chunkLines < 0)
        {
            throw new CommandErrorException("invalid_arguments", "Argument --chunk-lines must be greater than or equal to zero.");
        }

        if (chunkIndex is < 0)
        {
            throw new CommandErrorException("invalid_arguments", "Argument --chunk-index must be greater than or equal to zero.");
        }

        if (maxChars < 0)
        {
            throw new CommandErrorException("invalid_arguments", "Argument --max-chars must be greater than or equal to zero.");
        }
    }

    private static IReadOnlyList<OutputPaneSnapshot> CaptureOutputPanes(OutputWindow outputWindow)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        List<OutputPaneSnapshot> panes = [];
        OutputWindowPanes outputPanes = outputWindow.OutputWindowPanes;
        for (int i = 1; i <= outputPanes.Count; i++)
        {
            OutputWindowPane pane = outputPanes.Item(i);
            panes.Add(CreatePaneSnapshot(pane, IsActivePane(outputWindow, pane)));
        }

        return panes;
    }

    private static OutputWindowPane SelectPane(OutputWindow outputWindow, IReadOnlyList<OutputPaneSnapshot> panes, string? requestedPane)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (!string.IsNullOrWhiteSpace(requestedPane))
        {
            OutputWindowPane? namedPane = TryFindPane(outputWindow, requestedPane!);
            if (namedPane is not null)
            {
                return namedPane;
            }

            throw new CommandErrorException(
                "output_pane_not_found",
                $"Output pane not found: {requestedPane}",
                new { requestedPane, availablePanes = panes });
        }

        OutputWindowPane? activePane = TryGetActivePane(outputWindow);
        if (activePane is not null)
        {
            return activePane;
        }

        OutputWindowPane? buildPane = TryFindPane(outputWindow, BuildPaneName);
        if (buildPane is not null)
        {
            return buildPane;
        }

        return outputWindow.OutputWindowPanes.Item(1);
    }

    private static OutputWindowPane? TryFindPane(OutputWindow outputWindow, string paneNameOrGuid)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        string normalizedQuery = NormalizeGuid(paneNameOrGuid);
        OutputWindowPanes panes = outputWindow.OutputWindowPanes;
        for (int i = 1; i <= panes.Count; i++)
        {
            OutputWindowPane pane = panes.Item(i);
            string paneName = SafeGetPaneName(pane);
            string paneGuid = SafeGetPaneGuid(pane);
            if (string.Equals(paneName, paneNameOrGuid, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(NormalizeGuid(paneGuid), normalizedQuery, StringComparison.OrdinalIgnoreCase))
            {
                return pane;
            }
        }

        return null;
    }

    private static OutputWindowPane? TryGetActivePane(OutputWindow outputWindow)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            return outputWindow.ActivePane;
        }
        catch (COMException)
        {
            return null;
        }
        catch (ArgumentException)
        {
            return null;
        }
    }

    private static string ReadOutputPaneText(DTE2 dte, OutputWindowPane pane, bool activate)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (activate)
        {
            ActivateOutputPane(dte, pane);
        }

        for (int attempt = 0; attempt < OutputReadAttemptCount; attempt++)
        {
            try
            {
                return pane.TextDocument is TextDocument textDocument
                    ? ReadTextDocument(textDocument)
                    : string.Empty;
            }
            catch (COMException ex)
            {
                if (attempt == 0)
                {
                    ActivateOutputPane(dte, pane);
                    continue;
                }

                throw new CommandErrorException(
                    "output_pane_unreadable",
                    $"Could not read Output pane '{SafeGetPaneName(pane)}': {ex.Message}",
                    new { exception = ex.ToString() });
            }
        }

        return string.Empty;
    }

    private static void ActivateOutputPane(DTE2 dte, OutputWindowPane pane)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            dte.ExecuteCommand("View.Output", string.Empty);
        }
        catch (COMException ex)
        {
            TraceNonCriticalComException(ex);
        }

        try
        {
            pane.Activate();
        }
        catch (COMException ex)
        {
            TraceNonCriticalComException(ex);
        }
    }

    private static void TraceNonCriticalComException(COMException ex)
        => Debug.WriteLine(ex);

    private static string ReadTextDocument(TextDocument textDocument)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        EditPoint start = textDocument.StartPoint.CreateEditPoint();
        return start.GetText(textDocument.EndPoint);
    }

    private static OutputTextSlice SliceOutputText(string text, int chunkLines, int? chunkIndex, int maxChars)
    {
        List<string> lines = SplitLines(text);
        IReadOnlyList<OutputTextChunk> chunks = CreateTextChunks(lines, chunkLines);
        int selectedChunkIndex = chunkIndex ?? chunks.Count - 1;
        if (selectedChunkIndex >= chunks.Count)
        {
            throw new CommandErrorException(
                "invalid_arguments",
                $"Argument --chunk-index is out of range. Valid range is 0 to {chunks.Count - 1}.");
        }

        OutputTextChunk selectedChunk = chunks[selectedChunkIndex];
        string selectedText = selectedChunk.Text;
        bool charTruncated = maxChars > 0 && selectedText.Length > maxChars;
        if (charTruncated)
        {
            selectedText = selectedText.Substring(selectedText.Length - maxChars);
        }

        return new OutputTextSlice(
            selectedText,
            chunks,
            selectedChunk,
            lines.Count,
            SplitLines(selectedText).Count,
            chunks.Count > 1,
            charTruncated);
    }

    private static IReadOnlyList<OutputTextChunk> CreateTextChunks(List<string> lines, int chunkLines)
    {
        List<OutputTextChunk> chunks = [];
        if (lines.Count == 0)
        {
            chunks.Add(new OutputTextChunk(0, 0, 0, string.Empty));
            return chunks;
        }

        if (chunkLines == 0)
        {
            chunks.Add(new OutputTextChunk(0, 1, lines.Count, string.Join(Environment.NewLine, lines)));
            return chunks;
        }

        for (int start = 0, index = 0; start < lines.Count; start += chunkLines, index++)
        {
            int count = Math.Min(chunkLines, lines.Count - start);
            string chunkText = string.Join(Environment.NewLine, lines.GetRange(start, count));
            chunks.Add(new OutputTextChunk(index, start + 1, start + count, chunkText));
        }

        return chunks;
    }

    private static List<string> SplitLines(string text)
    {
        List<string> lines = [];
        using StringReader reader = new(text ?? string.Empty);
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            lines.Add(line);
        }

        return lines;
    }

    private static JArray CreateChunksArray(IReadOnlyList<OutputTextChunk> chunks, int selectedChunkIndex, bool includeText)
    {
        JArray array = [];
        foreach (OutputTextChunk chunk in chunks)
        {
            array.Add(CreateChunkObject(chunk, includeText, selected: chunk.Index == selectedChunkIndex));
        }

        return array;
    }

    private static JObject CreateChunkObject(
        OutputTextChunk chunk,
        bool includeText,
        bool selected,
        string? textOverride = null,
        bool charTruncated = false)
    {
        string returnedText = textOverride ?? chunk.Text;
        JObject obj = new()
        {
            ["index"] = chunk.Index,
            ["number"] = chunk.Index + 1,
            ["startLine"] = chunk.StartLine,
            ["endLine"] = chunk.EndLine,
            ["lineCount"] = chunk.LineCount,
            ["charCount"] = chunk.Text.Length,
            ["returnedChars"] = returnedText.Length,
            ["selected"] = selected,
            ["charTruncated"] = charTruncated,
        };

        if (includeText)
        {
            obj["text"] = returnedText;
        }

        return obj;
    }

    private static JArray CreatePanesArray(IReadOnlyList<OutputPaneSnapshot> panes, OutputPaneSnapshot selectedPane)
    {
        JArray array = [];
        foreach (OutputPaneSnapshot pane in panes)
        {
            bool selected = string.Equals(pane.Name, selectedPane.Name, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(NormalizeGuid(pane.Guid), NormalizeGuid(selectedPane.Guid), StringComparison.OrdinalIgnoreCase);
            array.Add(new JObject
            {
                ["name"] = pane.Name,
                ["guid"] = pane.Guid,
                ["isActive"] = pane.IsActive,
                ["selected"] = selected,
            });
        }

        return array;
    }

    private static OutputPaneSnapshot CreatePaneSnapshot(OutputWindowPane pane, bool isActive)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        return new(SafeGetPaneName(pane), SafeGetPaneGuid(pane), isActive);
    }

    private static bool IsActivePane(OutputWindow outputWindow, OutputWindowPane pane)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        OutputWindowPane? activePane = TryGetActivePane(outputWindow);
        if (activePane is null)
        {
            return false;
        }

        return string.Equals(SafeGetPaneName(activePane), SafeGetPaneName(pane), StringComparison.OrdinalIgnoreCase) &&
            string.Equals(NormalizeGuid(SafeGetPaneGuid(activePane)), NormalizeGuid(SafeGetPaneGuid(pane)), StringComparison.OrdinalIgnoreCase);
    }

    private static string SafeGetPaneName(OutputWindowPane pane)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            return pane.Name ?? string.Empty;
        }
        catch (COMException)
        {
            return string.Empty;
        }
    }

    private static string SafeGetPaneGuid(OutputWindowPane pane)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            return pane.Guid ?? string.Empty;
        }
        catch (COMException)
        {
            return string.Empty;
        }
    }

    private static string NormalizeGuid(string value)
        => string.IsNullOrWhiteSpace(value)
            ? string.Empty
            : value.Trim().Trim('{', '}');

    private sealed record OutputPaneSnapshot(string Name, string Guid, bool IsActive);

    private sealed record OutputTextChunk(int Index, int StartLine, int EndLine, string Text)
    {
        public int LineCount => StartLine == 0 && EndLine == 0 ? 0 : EndLine - StartLine + 1;
    }

    private sealed record OutputTextSlice(
        string Text,
        IReadOnlyList<OutputTextChunk> Chunks,
        OutputTextChunk SelectedChunk,
        int TotalLineCount,
        int ReturnedLineCount,
        bool LineTruncated,
        bool CharTruncated);
}
