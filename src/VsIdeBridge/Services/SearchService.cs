using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.FindResults;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;
using Newtonsoft.Json.Linq;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed class SearchService
{
    private sealed class SearchHit
    {
        public string Path { get; set; } = string.Empty;

        public string ProjectUniqueName { get; set; } = string.Empty;

        public int Line { get; set; }

        public int Column { get; set; }

        public int MatchLength { get; set; }

        public string Preview { get; set; } = string.Empty;

        public int ScoreHint { get; set; }

        public List<string> SourceQueries { get; set; } = new List<string>();
    }

    private sealed class SmartQueryTerm
    {
        public string Text { get; set; } = string.Empty;

        public int Weight { get; set; }

        public bool WholeWord { get; set; }
    }

    public async Task<JObject> FindFilesAsync(IdeCommandContext context, string query)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);

        var matches = new JArray(
            EnumerateSolutionFiles(context.Dte)
                .Where(item => item.Path.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                               Path.GetFileName(item.Path).IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
                .Select(item => new JObject
                {
                    ["path"] = item.Path,
                    ["project"] = item.ProjectUniqueName,
                }));

        return new JObject
        {
            ["query"] = query,
            ["count"] = matches.Count,
            ["matches"] = matches,
        };
    }

    public async Task<JObject> FindTextAsync(
        IdeCommandContext context,
        string query,
        string scope,
        bool matchCase,
        bool wholeWord,
        bool useRegex,
        int resultsWindow,
        string? projectUniqueName)
    {
        var searchData = await SearchTextMatchesAsync(
            context,
            query,
            scope,
            matchCase,
            wholeWord,
            useRegex,
            projectUniqueName).ConfigureAwait(true);

        await PopulateFindResultsAsync(context, searchData.GroupedMatches, query, resultsWindow).ConfigureAwait(true);

        return new JObject
        {
            ["query"] = query,
            ["scope"] = scope,
            ["count"] = searchData.Matches.Count,
            ["resultsWindow"] = resultsWindow,
            ["matches"] = new JArray(searchData.Matches.Select(SerializeHit)),
        };
    }

    public async Task<JObject> GetSmartContextForQueryAsync(
        IdeCommandContext context,
        string query,
        string scope,
        bool matchCase,
        bool wholeWord,
        bool useRegex,
        string? projectUniqueName,
        int maxContexts,
        int contextBefore,
        int contextAfter,
        bool populateResultsWindow,
        int resultsWindow)
    {
        IReadOnlyList<string> searchTerms;
        var searchData = useRegex
            ? await SearchTextMatchesAsync(
                context,
                query,
                scope,
                matchCase,
                wholeWord,
                true,
                projectUniqueName).ConfigureAwait(true)
            : await SearchSmartQueryTermsAsync(
                context,
                query,
                scope,
                projectUniqueName).ConfigureAwait(true);

        if (useRegex)
        {
            searchTerms = new[] { query };
        }
        else
        {
            searchTerms = ExtractSmartQueryTerms(query)
                .Select(term => term.Text)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        if (populateResultsWindow)
        {
            await PopulateFindResultsAsync(context, searchData.GroupedMatches, query, resultsWindow).ConfigureAwait(true);
        }

        var contexts = BuildSmartContexts(
            query,
            matchCase,
            wholeWord,
            useRegex,
            searchData.Matches,
            contextBefore,
            contextAfter,
            maxContexts);

        return new JObject
        {
            ["query"] = query,
            ["scope"] = scope,
            ["searchTerms"] = new JArray(searchTerms),
            ["totalMatchCount"] = searchData.Matches.Count,
            ["contextCount"] = contexts.Count,
            ["populateResultsWindow"] = populateResultsWindow,
            ["resultsWindow"] = resultsWindow,
            ["contexts"] = contexts,
        };
    }

    private async Task PopulateFindResultsAsync(
        IdeCommandContext context,
        IReadOnlyDictionary<string, List<FindResult>> groupedMatches,
        string query,
        int resultsWindow)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);

        var service = await context.Package.GetServiceAsync(typeof(SVsFindResults)).ConfigureAwait(true) as IFindResultsService;
        if (service is null)
        {
            return;
        }

        var title = $"IDE Bridge Find Results {resultsWindow}";
        var description = $"Find all \"{query}\"";
        var identifier = $"VsIdeBridge.FindResults.{resultsWindow}";
        var window = service.StartSearch(title, description, identifier);
        foreach (var item in groupedMatches)
        {
            window.AddResults(item.Key, item.Key, null, item.Value);
        }

        window.Summary = $"Matching lines: {groupedMatches.Sum(item => item.Value.Count)} Matching files: {groupedMatches.Count}";
        window.Complete();
    }

    private static Regex BuildRegex(string query, bool matchCase, bool wholeWord, bool useRegex)
    {
        var pattern = useRegex ? query : Regex.Escape(query);
        if (wholeWord)
        {
            pattern = $@"\b{pattern}\b";
        }

        var options = RegexOptions.Compiled;
        if (!matchCase)
        {
            options |= RegexOptions.IgnoreCase;
        }

        return new Regex(pattern, options);
    }

    private async Task<(List<SearchHit> Matches, Dictionary<string, List<FindResult>> GroupedMatches)> SearchTextMatchesAsync(
        IdeCommandContext context,
        string query,
        string scope,
        bool matchCase,
        bool wholeWord,
        bool useRegex,
        string? projectUniqueName)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);

        var files = scope switch
        {
            "document" => new[] { await GetDocumentTargetAsync(context).ConfigureAwait(true) },
            "project" => EnumerateSolutionFiles(context.Dte)
                .Where(item => string.Equals(item.ProjectUniqueName, projectUniqueName, StringComparison.OrdinalIgnoreCase))
                .ToArray(),
            _ => EnumerateSolutionFiles(context.Dte).ToArray(),
        };

        var regex = BuildRegex(query, matchCase, wholeWord, useRegex);
        var hits = new List<SearchHit>();
        var groupedMatches = new Dictionary<string, List<FindResult>>(StringComparer.OrdinalIgnoreCase);

        foreach (var file in files)
        {
            if (!File.Exists(file.Path))
            {
                continue;
            }

            var lines = File.ReadAllLines(file.Path);
            for (var lineIndex = 0; lineIndex < lines.Length; lineIndex++)
            {
                var line = lines[lineIndex];
                foreach (Match match in regex.Matches(line))
                {
                    hits.Add(new SearchHit
                    {
                        Path = file.Path,
                        ProjectUniqueName = file.ProjectUniqueName,
                        Line = lineIndex + 1,
                        Column = match.Index + 1,
                        MatchLength = match.Length,
                        Preview = line,
                        ScoreHint = 0,
                        SourceQueries = new List<string> { query },
                    });

                    if (!groupedMatches.TryGetValue(file.Path, out var results))
                    {
                        results = new List<FindResult>();
                        groupedMatches[file.Path] = results;
                    }

                    results.Add(new FindResult(line, lineIndex, match.Index, new Span(match.Index, match.Length)));
                }
            }
        }

        return (hits, groupedMatches);
    }

    private async Task<(List<SearchHit> Matches, Dictionary<string, List<FindResult>> GroupedMatches)> SearchSmartQueryTermsAsync(
        IdeCommandContext context,
        string query,
        string scope,
        string? projectUniqueName)
    {
        var terms = ExtractSmartQueryTerms(query);
        var hitMap = new Dictionary<string, SearchHit>(StringComparer.OrdinalIgnoreCase);
        var groupedMatches = new Dictionary<string, List<FindResult>>(StringComparer.OrdinalIgnoreCase);

        foreach (var term in terms)
        {
            var termResults = await SearchTextMatchesAsync(
                context,
                term.Text,
                scope,
                matchCase: false,
                wholeWord: term.WholeWord,
                useRegex: false,
                projectUniqueName).ConfigureAwait(true);

            foreach (var hit in termResults.Matches)
            {
                var key = string.Concat(hit.Path, "|", hit.Line.ToString(), "|", hit.Column.ToString());

                if (!hitMap.TryGetValue(key, out var existing))
                {
                    existing = new SearchHit
                    {
                        Path = hit.Path,
                        ProjectUniqueName = hit.ProjectUniqueName,
                        Line = hit.Line,
                        Column = hit.Column,
                        MatchLength = hit.MatchLength,
                        Preview = hit.Preview,
                        ScoreHint = 0,
                    };
                    hitMap[key] = existing;
                }

                existing.ScoreHint += term.Weight;
                if (!existing.SourceQueries.Contains(term.Text, StringComparer.OrdinalIgnoreCase))
                {
                    existing.SourceQueries.Add(term.Text);
                }

                if (!groupedMatches.TryGetValue(hit.Path, out var results))
                {
                    results = new List<FindResult>();
                    groupedMatches[hit.Path] = results;
                }

                results.Add(new FindResult(hit.Preview, hit.Line - 1, hit.Column - 1, new Span(hit.Column - 1, hit.MatchLength)));
            }
        }

        return (hitMap.Values
            .OrderByDescending(hit => hit.ScoreHint)
            .ThenBy(hit => hit.Path, StringComparer.OrdinalIgnoreCase)
            .ThenBy(hit => hit.Line)
            .ToList(), groupedMatches);
    }

    private static JArray BuildSmartContexts(
        string query,
        bool matchCase,
        bool wholeWord,
        bool useRegex,
        IReadOnlyList<SearchHit> hits,
        int contextBefore,
        int contextAfter,
        int maxContexts)
    {
        var before = Math.Max(0, contextBefore);
        var after = Math.Max(0, contextAfter);
        var limit = Math.Max(1, maxContexts);
        var contexts = new List<(int Score, int FirstLine, JObject Context)>();

        foreach (var fileGroup in hits.GroupBy(hit => hit.Path, StringComparer.OrdinalIgnoreCase))
        {
            var allLines = File.ReadAllLines(fileGroup.Key);
            var windows = new List<(int StartLine, int EndLine, List<SearchHit> Hits)>();
            foreach (var hit in fileGroup.OrderBy(hit => hit.Line).ThenBy(hit => hit.Column))
            {
                var startLine = Math.Max(1, hit.Line - before);
                var endLine = Math.Min(allLines.Length, hit.Line + after);
                var merged = false;

                for (var i = 0; i < windows.Count; i++)
                {
                    var existing = windows[i];
                    if (startLine <= existing.EndLine + 1)
                    {
                        existing.StartLine = Math.Min(existing.StartLine, startLine);
                        existing.EndLine = Math.Max(existing.EndLine, endLine);
                        existing.Hits.Add(hit);
                        windows[i] = existing;
                        merged = true;
                        break;
                    }
                }

                if (!merged)
                {
                    windows.Add((startLine, endLine, new List<SearchHit> { hit }));
                }
            }

            foreach (var window in windows)
            {
                var textLines = new JArray();
                var builder = new System.Text.StringBuilder();
                for (var lineNumber = window.StartLine; lineNumber <= window.EndLine; lineNumber++)
                {
                    var lineText = allLines[lineNumber - 1];
                    textLines.Add(new JObject
                    {
                        ["line"] = lineNumber,
                        ["text"] = lineText,
                    });

                    if (builder.Length > 0)
                    {
                        builder.Append('\n');
                    }

                    builder.Append(lineNumber);
                    builder.Append(": ");
                    builder.Append(lineText);
                }

                var score = ScoreSmartContext(window.Hits);
                contexts.Add((score, window.Hits.Min(hit => hit.Line), new JObject
                {
                    ["path"] = fileGroup.Key,
                    ["project"] = window.Hits[0].ProjectUniqueName,
                    ["startLine"] = window.StartLine,
                    ["endLine"] = window.EndLine,
                    ["score"] = score,
                    ["hits"] = new JArray(window.Hits.Select(SerializeHit)),
                    ["text"] = builder.ToString(),
                    ["lines"] = textLines,
                }));
            }
        }

        return new JArray(contexts
            .OrderByDescending(item => item.Score)
            .ThenBy(item => item.FirstLine)
            .Take(limit)
            .Select(item => item.Context));
    }

    private static int ScoreSmartContext(IReadOnlyList<SearchHit> hits)
    {
        var score = 0;
        foreach (var hit in hits)
        {
            var preview = hit.Preview ?? string.Empty;
            score += Math.Max(20, hit.ScoreHint);

            foreach (var query in hit.SourceQueries.DefaultIfEmpty(string.Empty).Take(3))
            {
                if (string.IsNullOrWhiteSpace(query))
                {
                    continue;
                }

                var escaped = Regex.Escape(query);
                var declarationPattern = new Regex($@"^\s*(class|struct|enum|namespace)\s+{escaped}\b", RegexOptions.IgnoreCase);
                var callablePattern = new Regex($@"(\b{escaped}\s*\()|(::\s*{escaped}\b)", RegexOptions.IgnoreCase);
                var identifierPattern = new Regex($@"\b{escaped}\b", RegexOptions.IgnoreCase);

                if (declarationPattern.IsMatch(preview))
                {
                    score += 120;
                }
                else if (callablePattern.IsMatch(preview))
                {
                    score += 80;
                }
                else if (identifierPattern.IsMatch(preview))
                {
                    score += 40;
                }
                else
                {
                    score += 10;
                }
            }
        }

        return score;
    }

    private static JObject SerializeHit(SearchHit hit)
    {
        return new JObject
        {
            ["path"] = hit.Path,
            ["project"] = hit.ProjectUniqueName,
            ["line"] = hit.Line,
            ["column"] = hit.Column,
            ["matchLength"] = hit.MatchLength,
            ["preview"] = hit.Preview,
            ["scoreHint"] = hit.ScoreHint,
            ["queries"] = new JArray(hit.SourceQueries),
        };
    }

    private static IReadOnlyList<SmartQueryTerm> ExtractSmartQueryTerms(string query)
    {
        var terms = new List<SmartQueryTerm>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void AddTerm(string value, int weight, bool wholeWord)
        {
            var trimmed = value.Trim();
            if (trimmed.Length < 2 || !seen.Add(trimmed))
            {
                return;
            }

            terms.Add(new SmartQueryTerm
            {
                Text = trimmed,
                Weight = weight,
                WholeWord = wholeWord,
            });
        }

        foreach (Match match in Regex.Matches(query, "\"([^\"]+)\""))
        {
            AddTerm(match.Groups[1].Value, 220, wholeWord: false);
        }

        foreach (Match match in Regex.Matches(query, @"[A-Za-z_][A-Za-z0-9_:/\\.\-]*"))
        {
            var token = match.Value;
            var looksLikeIdentifier = token.Contains("::", StringComparison.Ordinal) ||
                                      token.Contains("_", StringComparison.Ordinal) ||
                                      token.Contains(".", StringComparison.Ordinal) ||
                                      char.IsUpper(token[0]);

            if (looksLikeIdentifier)
            {
                AddTerm(token, 160, wholeWord: !token.Contains(".", StringComparison.Ordinal));
            }
        }

        var stopWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "where", "what", "when", "which", "that", "this", "with", "from", "into", "used", "using",
            "call", "calls", "show", "find", "open", "close", "line", "file", "files", "query", "context",
        };

        foreach (Match match in Regex.Matches(query, @"[A-Za-z][A-Za-z0-9_]{3,}"))
        {
            if (!stopWords.Contains(match.Value))
            {
                AddTerm(match.Value, 80, wholeWord: true);
            }
        }

        if (terms.Count == 0)
        {
            AddTerm(query, 120, wholeWord: false);
        }

        return terms;
    }

    private async Task<(string Path, string ProjectUniqueName)> GetDocumentTargetAsync(IdeCommandContext context)
    {
        var activeDocument = context.Dte.ActiveDocument;
        if (activeDocument is null || string.IsNullOrWhiteSpace(activeDocument.FullName))
        {
            throw new CommandErrorException("document_not_found", "There is no active document.");
        }

        await Task.CompletedTask;
        return (PathNormalization.NormalizeFilePath(activeDocument.FullName), activeDocument.ProjectItem?.ContainingProject?.UniqueName ?? string.Empty);
    }

    private static IEnumerable<(string Path, string ProjectUniqueName)> EnumerateSolutionFiles(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (dte.Solution is null || !dte.Solution.IsOpen)
        {
            yield break;
        }

        foreach (Project project in dte.Solution.Projects)
        {
            foreach (var file in EnumerateProjectFiles(project))
            {
                yield return file;
            }
        }
    }

    private static IEnumerable<(string Path, string ProjectUniqueName)> EnumerateProjectFiles(Project? project)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (project is null)
        {
            yield break;
        }

        if (string.Equals(project.Kind, EnvDTE80.ProjectKinds.vsProjectKindSolutionFolder, StringComparison.OrdinalIgnoreCase))
        {
            foreach (ProjectItem item in project.ProjectItems)
            {
                if (item.SubProject is not null)
                {
                    foreach (var file in EnumerateProjectFiles(item.SubProject))
                    {
                        yield return file;
                    }
                }
            }

            yield break;
        }

        foreach (ProjectItem item in project.ProjectItems)
        {
            foreach (var file in EnumerateProjectItemFiles(item, project.UniqueName))
            {
                yield return file;
            }
        }
    }

    private static IEnumerable<(string Path, string ProjectUniqueName)> EnumerateProjectItemFiles(ProjectItem item, string projectUniqueName)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (item.FileCount > 0)
        {
            for (short i = 1; i <= item.FileCount; i++)
            {
                var fileName = item.FileNames[i];
                if (!string.IsNullOrWhiteSpace(fileName))
                {
                    yield return (PathNormalization.NormalizeFilePath(fileName), projectUniqueName);
                }
            }
        }

        if (item.ProjectItems is null)
        {
            yield break;
        }

        foreach (ProjectItem child in item.ProjectItems)
        {
            foreach (var file in EnumerateProjectItemFiles(child, projectUniqueName))
            {
                yield return file;
            }
        }
    }
}
