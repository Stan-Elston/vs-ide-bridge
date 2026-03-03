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

    private sealed class CodeModelHit
    {
        public string Path { get; set; } = string.Empty;

        public string ProjectUniqueName { get; set; } = string.Empty;

        public string Name { get; set; } = string.Empty;

        public string FullName { get; set; } = string.Empty;

        public string Kind { get; set; } = string.Empty;

        public string Signature { get; set; } = string.Empty;

        public int Line { get; set; }

        public int EndLine { get; set; }

        public int Score { get; set; }

        public string MatchKind { get; set; } = string.Empty;
    }

    private sealed class SmartQueryTerm
    {
        public string Text { get; set; } = string.Empty;

        public int Weight { get; set; }

        public bool WholeWord { get; set; }
    }

    private static readonly HashSet<vsCMElement> s_codeModelKinds = new()
    {
        vsCMElement.vsCMElementFunction,
        vsCMElement.vsCMElementClass,
        vsCMElement.vsCMElementStruct,
        vsCMElement.vsCMElementEnum,
        vsCMElement.vsCMElementNamespace,
        vsCMElement.vsCMElementInterface,
        vsCMElement.vsCMElementProperty,
        vsCMElement.vsCMElementVariable,
    };

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
        string? projectUniqueName,
        string? pathFilter = null)
    {
        var searchData = await SearchTextMatchesAsync(
            context,
            query,
            scope,
            matchCase,
            wholeWord,
            useRegex,
            projectUniqueName,
            pathFilter).ConfigureAwait(true);

        await PopulateFindResultsAsync(context, searchData.GroupedMatches, query, resultsWindow).ConfigureAwait(true);

        return new JObject
        {
            ["query"] = query,
            ["scope"] = scope,
            ["pathFilter"] = pathFilter ?? string.Empty,
            ["count"] = searchData.Matches.Count,
            ["resultsWindow"] = resultsWindow,
            ["matches"] = new JArray(searchData.Matches.Select(SerializeHit)),
        };
    }

    public async Task<JObject> SearchSymbolsAsync(
        IdeCommandContext context,
        string name,
        string kind,
        string scope,
        bool matchCase,
        string? projectUniqueName,
        string? pathFilter,
        int max)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);

        var codeModelHits = SearchCodeModelSymbols(
            context.Dte,
            name,
            kind,
            scope,
            matchCase,
            projectUniqueName,
            pathFilter)
            .Take(Math.Max(1, max))
            .ToArray();

        if (codeModelHits.Length > 0)
        {
            return new JObject
            {
                ["query"] = name,
                ["kind"] = kind,
                ["scope"] = scope,
                ["project"] = projectUniqueName ?? string.Empty,
                ["pathFilter"] = pathFilter ?? string.Empty,
                ["count"] = codeModelHits.Length,
                ["totalMatchCount"] = codeModelHits.Length,
                ["source"] = "code-model",
                ["matches"] = new JArray(codeModelHits.Select(SerializeCodeModelHit)),
            };
        }

        // Build a regex pattern targeting definition signatures for the requested kind.
        var escaped = Regex.Escape(name);
        string pattern;
        string resolvedKind;
        switch (kind.ToLowerInvariant())
        {
            case "function":
                pattern = $@"\b{escaped}\s*\(";
                resolvedKind = "function";
                break;
            case "class":
                pattern = $@"\bclass\s+{escaped}\b";
                resolvedKind = "class";
                break;
            case "struct":
                pattern = $@"\bstruct\s+{escaped}\b";
                resolvedKind = "struct";
                break;
            case "enum":
                pattern = $@"\benum(?:\s+class)?\s+{escaped}\b";
                resolvedKind = "enum";
                break;
            case "namespace":
                pattern = $@"\bnamespace\s+{escaped}\b";
                resolvedKind = "namespace";
                break;
            case "interface":
                pattern = $@"\binterface\s+{escaped}\b";
                resolvedKind = "interface";
                break;
            case "member":
                pattern = $@"\b{escaped}\b";
                resolvedKind = "member";
                break;
            case "type":
                pattern = $@"\b{escaped}\b";
                resolvedKind = "type";
                break;
            default:
                // "all" — whole-word match; kind is inferred per-hit
                pattern = $@"\b{escaped}\b";
                resolvedKind = "all";
                break;
        }

        var searchData = await SearchTextMatchesAsync(
            context,
            pattern,
            scope,
            matchCase,
            wholeWord: false,
            useRegex: true,
            projectUniqueName,
            pathFilter).ConfigureAwait(true);

        // Annotate each hit with an inferred kind and cap at max
        var hits = searchData.Matches
            .Take(max)
            .Select(hit =>
            {
                var inferredKind = resolvedKind == "all"
                    ? InferSymbolKind(hit.Preview, name)
                    : resolvedKind;
                var obj = SerializeHit(hit);
                obj["inferredKind"] = inferredKind;
                return obj;
            })
            .ToArray();

        return new JObject
        {
            ["query"] = name,
            ["kind"] = kind,
            ["scope"] = scope,
            ["project"] = projectUniqueName ?? string.Empty,
            ["pathFilter"] = pathFilter ?? string.Empty,
            ["count"] = hits.Length,
            ["totalMatchCount"] = searchData.Matches.Count,
            ["source"] = "text",
            ["matches"] = new JArray(hits),
        };
    }

    private static string InferSymbolKind(string lineText, string name)
    {
        var trimmed = lineText.TrimStart();
        // Class/struct/enum/namespace — look for keyword immediately before name
        if (Regex.IsMatch(trimmed, $@"\bclass\s+{Regex.Escape(name)}\b", RegexOptions.IgnoreCase))
            return "class";
        if (Regex.IsMatch(trimmed, $@"\bstruct\s+{Regex.Escape(name)}\b", RegexOptions.IgnoreCase))
            return "struct";
        if (Regex.IsMatch(trimmed, $@"\benum(?:\s+class)?\s+{Regex.Escape(name)}\b", RegexOptions.IgnoreCase))
            return "enum";
        if (Regex.IsMatch(trimmed, $@"\bnamespace\s+{Regex.Escape(name)}\b", RegexOptions.IgnoreCase))
            return "namespace";
        if (Regex.IsMatch(trimmed, $@"\binterface\s+{Regex.Escape(name)}\b", RegexOptions.IgnoreCase))
            return "interface";
        // Function: name followed by (
        if (Regex.IsMatch(trimmed, $@"\b{Regex.Escape(name)}\s*\(", RegexOptions.IgnoreCase))
            return "function";
        return "unknown";
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
        string? projectUniqueName,
        string? pathFilter = null)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);

        var allFiles = scope switch
        {
            "document" => new[] { await GetDocumentTargetAsync(context).ConfigureAwait(true) },
            "open" => EnumerateOpenFiles(context.Dte).ToArray(),
            "project" => EnumerateSolutionFiles(context.Dte)
                .Where(item => string.Equals(item.ProjectUniqueName, projectUniqueName, StringComparison.OrdinalIgnoreCase))
                .ToArray(),
            _ => EnumerateSolutionFiles(context.Dte).ToArray(),
        };

        var files = string.IsNullOrWhiteSpace(pathFilter)
            ? allFiles
            : allFiles.Where(f => f.Path.IndexOf(pathFilter, StringComparison.OrdinalIgnoreCase) >= 0).ToArray();

        var regex = BuildRegex(query, matchCase, wholeWord, useRegex);
        var hits = new List<SearchHit>();
        var groupedMatches = new Dictionary<string, List<FindResult>>(StringComparer.OrdinalIgnoreCase);

        foreach (var file in files)
        {
            if (!File.Exists(file.Path))
            {
                continue;
            }

            var lines = ReadSearchLines(context.Dte, file.Path);
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

    private static JObject SerializeCodeModelHit(CodeModelHit hit)
    {
        return new JObject
        {
            ["name"] = hit.Name,
            ["fullName"] = hit.FullName,
            ["kind"] = hit.Kind,
            ["signature"] = hit.Signature,
            ["path"] = hit.Path,
            ["project"] = hit.ProjectUniqueName,
            ["line"] = hit.Line,
            ["column"] = 1,
            ["endLine"] = hit.EndLine,
            ["matchKind"] = hit.MatchKind,
            ["scoreHint"] = hit.Score,
            ["preview"] = hit.Signature,
            ["source"] = "code-model",
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
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);

        var activeDocument = context.Dte.ActiveDocument;
        if (activeDocument is null || string.IsNullOrWhiteSpace(activeDocument.FullName))
        {
            throw new CommandErrorException("document_not_found", "There is no active document.");
        }

        return (
            PathNormalization.NormalizeFilePath(activeDocument.FullName),
            activeDocument.ProjectItem?.ContainingProject?.UniqueName ?? string.Empty);
    }

    private static IEnumerable<CodeModelHit> SearchCodeModelSymbols(
        DTE2 dte,
        string query,
        string kind,
        string scope,
        bool matchCase,
        string? projectUniqueName,
        string? pathFilter)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var activeDocument = TryGetActiveDocumentTarget(dte);
        var files = scope switch
        {
            "document" => string.IsNullOrWhiteSpace(activeDocument.Path)
                ? Array.Empty<(string Path, string ProjectUniqueName)>()
                : new[] { activeDocument },
            "open" => EnumerateOpenFiles(dte),
            "project" => EnumerateSolutionFiles(dte)
                .Where(item => string.Equals(item.ProjectUniqueName, projectUniqueName, StringComparison.OrdinalIgnoreCase)),
            _ => EnumerateSolutionFiles(dte),
        };

        if (!string.IsNullOrWhiteSpace(pathFilter))
        {
            files = files.Where(item => item.Path.IndexOf(pathFilter, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var hits = new List<CodeModelHit>();
        foreach (var file in files)
        {
            if (string.IsNullOrWhiteSpace(file.Path))
            {
                continue;
            }

            ProjectItem? projectItem = null;
            CodeElements? elements = null;
            try
            {
                projectItem = dte.Solution.FindProjectItem(file.Path);
                elements = projectItem?.FileCodeModel?.CodeElements;
            }
            catch
            {
            }

            if (projectItem is null || elements is null)
            {
                continue;
            }

            foreach (CodeElement element in elements)
            {
                CollectMatchingSymbols(element, file.Path, file.ProjectUniqueName, query, kind, comparison, hits, seen);
            }
        }

        return hits
            .OrderByDescending(hit => hit.Score)
            .ThenBy(hit => hit.Path, StringComparer.OrdinalIgnoreCase)
            .ThenBy(hit => hit.Line)
            .ThenBy(hit => hit.Name, StringComparer.OrdinalIgnoreCase);
    }

    private static void CollectMatchingSymbols(
        CodeElement element,
        string path,
        string projectUniqueName,
        string query,
        string kindFilter,
        StringComparison comparison,
        List<CodeModelHit> hits,
        HashSet<string> seen)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        vsCMElement kind;
        try
        {
            kind = element.Kind;
        }
        catch
        {
            return;
        }

        if (s_codeModelKinds.Contains(kind))
        {
            var normalizedKind = NormalizeKind(kind);
            if (MatchesKind(kindFilter, normalizedKind))
            {
                var name = TryGetElementName(element);
                var fullName = TryGetFullName(element);
                var score = ScoreSymbolMatch(query, name, fullName, comparison, out var matchKind);
                if (score > 0)
                {
                    var line = TryGetLine(element.StartPoint);
                    var endLine = TryGetLine(element.EndPoint);
                    var signature = TryGetSignature(element, fullName, name);
                    var key = string.Concat(path, "|", normalizedKind, "|", fullName, "|", line.ToString());
                    if (seen.Add(key))
                    {
                        hits.Add(new CodeModelHit
                        {
                            Path = path,
                            ProjectUniqueName = projectUniqueName,
                            Name = name,
                            FullName = fullName,
                            Kind = normalizedKind,
                            Signature = signature,
                            Line = line,
                            EndLine = endLine,
                            Score = score,
                            MatchKind = matchKind,
                        });
                    }
                }
            }
        }

        foreach (var child in EnumerateChildren(element))
        {
            CollectMatchingSymbols(child, path, projectUniqueName, query, kindFilter, comparison, hits, seen);
        }
    }

    private static IEnumerable<CodeElement> EnumerateChildren(CodeElement element)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        CodeElements? children = null;
        try
        {
            children = element switch
            {
                CodeNamespace codeNamespace => codeNamespace.Members,
                CodeClass codeClass => codeClass.Members,
                CodeStruct codeStruct => codeStruct.Members,
                CodeInterface codeInterface => codeInterface.Members,
                _ => null,
            };
        }
        catch
        {
        }

        if (children is null)
        {
            yield break;
        }

        foreach (CodeElement child in children)
        {
            yield return child;
        }
    }

    private static string NormalizeKind(vsCMElement kind)
    {
        return kind switch
        {
            vsCMElement.vsCMElementFunction => "function",
            vsCMElement.vsCMElementClass => "class",
            vsCMElement.vsCMElementStruct => "struct",
            vsCMElement.vsCMElementEnum => "enum",
            vsCMElement.vsCMElementNamespace => "namespace",
            vsCMElement.vsCMElementInterface => "interface",
            vsCMElement.vsCMElementProperty => "member",
            vsCMElement.vsCMElementVariable => "member",
            _ => "unknown",
        };
    }

    private static bool MatchesKind(string kindFilter, string normalizedKind)
    {
        return kindFilter.ToLowerInvariant() switch
        {
            "all" => true,
            "type" => normalizedKind is "class" or "struct" or "enum" or "interface",
            "member" => normalizedKind is "member" or "function",
            _ => string.Equals(kindFilter, normalizedKind, StringComparison.OrdinalIgnoreCase),
        };
    }

    private static string TryGetElementName(CodeElement element)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            return element.Name ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string TryGetFullName(CodeElement element)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            return string.IsNullOrWhiteSpace(element.FullName) ? TryGetElementName(element) : element.FullName;
        }
        catch
        {
            return TryGetElementName(element);
        }
    }

    private static string TryGetSignature(CodeElement element, string fullName, string name)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            if (element is CodeFunction function)
            {
                return function.get_Prototype(
                    (int)(vsCMPrototype.vsCMPrototypeFullname | vsCMPrototype.vsCMPrototypeParamTypes | vsCMPrototype.vsCMPrototypeType))
                    ?? fullName;
            }
        }
        catch
        {
        }

        return string.IsNullOrWhiteSpace(fullName) ? name : fullName;
    }

    private static int TryGetLine(TextPoint? point)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            return point?.Line ?? 0;
        }
        catch
        {
            return 0;
        }
    }

    private static int ScoreSymbolMatch(
        string query,
        string name,
        string fullName,
        StringComparison comparison,
        out string matchKind)
    {
        matchKind = string.Empty;
        if (string.IsNullOrWhiteSpace(query))
        {
            return 0;
        }

        if (string.Equals(name, query, comparison))
        {
            matchKind = "name-exact";
            return 1000;
        }

        if (string.Equals(fullName, query, comparison))
        {
            matchKind = "full-name-exact";
            return 950;
        }

        if (name.StartsWith(query, comparison))
        {
            matchKind = "name-prefix";
            return 875;
        }

        if (fullName.StartsWith(query, comparison))
        {
            matchKind = "full-name-prefix";
            return 850;
        }

        if (name.IndexOf(query, comparison) >= 0)
        {
            matchKind = "name-contains";
            return 760;
        }

        if (fullName.IndexOf(query, comparison) >= 0)
        {
            matchKind = "full-name-contains";
            return 720;
        }

        return 0;
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

    private static IEnumerable<(string Path, string ProjectUniqueName)> EnumerateOpenFiles(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        foreach (Document document in dte.Documents)
        {
            string? fullName = null;
            try
            {
                fullName = document.FullName;
            }
            catch
            {
            }

            if (string.IsNullOrWhiteSpace(fullName))
            {
                continue;
            }

            var normalizedPath = PathNormalization.NormalizeFilePath(fullName!);
            yield return (normalizedPath, document.ProjectItem?.ContainingProject?.UniqueName ?? string.Empty);
        }
    }

    private static (string Path, string ProjectUniqueName) TryGetActiveDocumentTarget(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var activeDocument = dte.ActiveDocument;
        if (activeDocument is null || string.IsNullOrWhiteSpace(activeDocument.FullName))
        {
            return (string.Empty, string.Empty);
        }

        return (PathNormalization.NormalizeFilePath(activeDocument.FullName), activeDocument.ProjectItem?.ContainingProject?.UniqueName ?? string.Empty);
    }

    private static string[] ReadSearchLines(DTE2 dte, string path)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var normalizedPath = PathNormalization.NormalizeFilePath(path);
        foreach (Document document in dte.Documents)
        {
            try
            {
                if (!string.Equals(PathNormalization.NormalizeFilePath(document.FullName), normalizedPath, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (document.Object("TextDocument") is TextDocument textDocument)
                {
                    var editPoint = textDocument.StartPoint.CreateEditPoint();
                    var text = editPoint.GetText(textDocument.EndPoint);
                    return text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
                }
            }
            catch
            {
            }
        }

        return File.ReadAllLines(normalizedPath);
    }
}
