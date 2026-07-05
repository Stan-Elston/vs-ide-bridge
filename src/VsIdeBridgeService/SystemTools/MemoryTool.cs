using System.Globalization;
using System.Text.Json.Nodes;

namespace VsIdeBridgeService.SystemTools;

internal static class MemoryTool
{
    private const int DefaultSearchResults = 25;
    private const int MaxSearchResults = 100;
    private const int DefaultSearchContextLines = 1;
    private const int MaxSearchContextLines = 5;
    private const int DefaultReadLines = 120;
    private const int MaxReadLines = 250;
    private const int DefaultCenteredContextLines = 20;
    private const int MaxCenteredContextLines = 100;
    private const int SearchContextWindowRadius = 2;
    private const int TwoTermQueryCount = 2;
    private const int ThreeTermQueryCount = 3;
    private const int MinimumMultiTermMatches = TwoTermQueryCount;
    private const long MaxSearchFileBytes = 5_000_000;

    public static Task<JsonNode> SearchAsync(JsonNode? id, JsonObject? args, BridgeConnection bridge)
    {
        string query = (args?["query"]?.GetValue<string>() ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(query))
        {
            throw new McpRequestException(id, McpErrorCodes.InvalidParams, "Missing required argument 'query'.");
        }

        int maxResults = Clamp(args?["max_results"]?.GetValue<int?>() ?? DefaultSearchResults, 1, MaxSearchResults);
        int contextLines = Clamp(args?["context_lines"]?.GetValue<int?>() ?? DefaultSearchContextLines, 0, MaxSearchContextLines);
        bool includeRollouts = args?["include_rollouts"]?.GetValue<bool?>() ?? true;
        string memoryRoot = ResolveMemoryRoot(id, bridge);
        List<string> queryTerms = BuildQueryTerms(query);

        List<SearchCandidate> candidates = [];
        int searchedFiles = 0;
        int candidateOrder = 0;

        foreach (string file in EnumerateMemoryFiles(memoryRoot, includeRollouts))
        {
            searchedFiles++;
            if (ShouldSkipSearchFile(file))
            {
                continue;
            }

            string[] lines;
            try
            {
                lines = File.ReadAllLines(file);
            }
            catch (IOException)
            {
                continue;
            }
            catch (UnauthorizedAccessException)
            {
                continue;
            }

            string relativePath = ToMemoryRelativePath(memoryRoot, file);
            for (int index = 0; index < lines.Length; index++)
            {
                int score = ScoreSearchLine(query, queryTerms, relativePath, lines, index,
                    out int matchedTerms, out bool exactMatch);
                if (score <= 0)
                {
                    continue;
                }

                candidates.Add(new SearchCandidate(file, index, score, matchedTerms, exactMatch, candidateOrder++));
            }
        }

        List<SearchCandidate> selected = [.. candidates
            .OrderByDescending(candidate => candidate.ExactMatch)
            .ThenByDescending(candidate => candidate.Score)
            .ThenByDescending(candidate => candidate.MatchedTerms)
            .ThenBy(candidate => candidate.Order)
            .Take(maxResults)];

        JsonArray results = [];
        foreach (SearchCandidate candidate in selected)
        {
            string[] lines = File.ReadAllLines(candidate.File);
            results.Add(BuildSearchHit(memoryRoot, candidate.File, candidate.LineIndex, lines, contextLines,
                candidate.Score, candidate.MatchedTerms, candidate.ExactMatch));
        }

        JsonObject payload = new()
        {
            ["success"] = true,
            ["memoryRoot"] = memoryRoot,
            ["query"] = query,
            ["queryTerms"] = new JsonArray([.. queryTerms.Select(term => JsonValue.Create(term))]),
            ["maxResults"] = maxResults,
            ["contextLines"] = contextLines,
            ["includeRollouts"] = includeRollouts,
            ["searchedFiles"] = searchedFiles,
            ["matchCount"] = candidates.Count,
            ["count"] = results.Count,
            ["truncated"] = candidates.Count > selected.Count,
            ["results"] = results,
        };

        string successText = $"Found {results.Count} memory match(es) across {searchedFiles} file(s) " +
            $"from {candidates.Count} candidate line(s).";
        return Task.FromResult<JsonNode>(ToolResultFormatter.StructuredToolResult(payload, args, successText: successText));
    }

    public static Task<JsonNode> ReadAsync(JsonNode? id, JsonObject? args, BridgeConnection bridge)
    {
        string memoryRoot = ResolveMemoryRoot(id, bridge);
        string relativePath = args?["path"]?.GetValue<string>() ?? "MEMORY.md";
        string fullPath = ResolveMemoryPath(id, memoryRoot, relativePath);

        int? centerLine = args?["line"]?.GetValue<int?>();
        int startLine;
        int endLine;
        if (centerLine is not null)
        {
            if (centerLine <= 0)
            {
                throw new McpRequestException(id, McpErrorCodes.InvalidParams, "Argument 'line' must be greater than zero.");
            }

            int contextLines = Clamp(args?["context_lines"]?.GetValue<int?>() ?? DefaultCenteredContextLines, 0, MaxCenteredContextLines);
            startLine = Math.Max(1, centerLine.Value - contextLines);
            endLine = centerLine.Value + contextLines;
        }
        else
        {
            startLine = Math.Max(1, args?["start_line"]?.GetValue<int?>() ?? 1);
            endLine = args?["end_line"]?.GetValue<int?>() ?? (startLine + DefaultReadLines - 1);
        }

        if (endLine < startLine)
        {
            throw new McpRequestException(id, McpErrorCodes.InvalidParams, "Argument 'end_line' must be greater than or equal to 'start_line'.");
        }

        if (endLine - startLine + 1 > MaxReadLines)
        {
            endLine = startLine + MaxReadLines - 1;
        }

        List<string> selected = [];
        int currentLine = 0;
        foreach (string line in File.ReadLines(fullPath))
        {
            currentLine++;
            if (currentLine < startLine)
            {
                continue;
            }

            if (currentLine > endLine)
            {
                break;
            }

            selected.Add($"{currentLine.ToString(CultureInfo.InvariantCulture)}: {line}");
        }

        JsonObject payload = new()
        {
            ["success"] = true,
            ["memoryRoot"] = memoryRoot,
            ["path"] = ToMemoryRelativePath(memoryRoot, fullPath),
            ["startLine"] = startLine,
            ["endLine"] = startLine + Math.Max(0, selected.Count - 1),
            ["lineCount"] = selected.Count,
            ["text"] = string.Join(Environment.NewLine, selected),
        };

        string successText = $"Read {selected.Count} memory line(s) from {payload["path"]!.GetValue<string>()}.";
        return Task.FromResult<JsonNode>(ToolResultFormatter.StructuredToolResult(payload, args, successText: successText));
    }

    private sealed record SearchCandidate(
        string File,
        int LineIndex,
        int Score,
        int MatchedTerms,
        bool ExactMatch,
        int Order);

    private static JsonObject BuildSearchHit(
        string root,
        string file,
        int lineIndex,
        string[] lines,
        int contextLines,
        int score,
        int matchedTerms,
        bool exactMatch)
    {
        int lineNumber = lineIndex + 1;
        int start = Math.Max(0, lineIndex - contextLines);
        int end = Math.Min(lines.Length - 1, lineIndex + contextLines);
        JsonArray context = [];
        for (int index = start; index <= end; index++)
        {
            context.Add(new JsonObject
            {
                ["line"] = index + 1,
                ["text"] = Truncate(lines[index], 500),
            });
        }

        return new JsonObject
        {
            ["path"] = ToMemoryRelativePath(root, file),
            ["line"] = lineNumber,
            ["preview"] = Truncate(lines[lineIndex], 500),
            ["matchType"] = exactMatch ? "exact" : "terms",
            ["score"] = score,
            ["matchedTerms"] = matchedTerms,
            ["context"] = context,
        };
    }

    private static List<string> BuildQueryTerms(string query)
    {
        List<string> terms = [];
        int start = -1;

        for (int index = 0; index <= query.Length; index++)
        {
            bool inTerm = index < query.Length && IsSearchTermChar(query[index]);
            if (inTerm && start < 0)
            {
                start = index;
                continue;
            }

            if (inTerm || start < 0)
            {
                continue;
            }

            AddQueryTerm(query[start..index], terms);
            start = -1;
        }

        return terms.Count > 0 ? terms : [query];
    }

    private static void AddQueryTerm(string term, List<string> terms)
    {
        if (IsStopWord(term) || terms.Contains(term, StringComparer.OrdinalIgnoreCase))
        {
            return;
        }

        terms.Add(term);
    }

    private static int ScoreSearchLine(
        string query,
        List<string> queryTerms,
        string relativePath,
        string[] lines,
        int lineIndex,
        out int matchedTerms,
        out bool exactMatch)
    {
        string line = lines[lineIndex];
        string searchableText = BuildSearchableText(relativePath, lines, lineIndex);
        exactMatch = line.Contains(query, StringComparison.OrdinalIgnoreCase);
        matchedTerms = 0;
        int lineTermMatches = 0;

        foreach (string term in queryTerms)
        {
            if (!searchableText.Contains(term, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            matchedTerms++;
            if (line.Contains(term, StringComparison.OrdinalIgnoreCase))
            {
                lineTermMatches++;
            }
        }

        if (!exactMatch && (lineTermMatches == 0 || matchedTerms < RequiredTermMatches(queryTerms.Count)))
        {
            return 0;
        }

        int score = exactMatch ? 10_000 : 0;
        score += matchedTerms * 100;
        score += lineTermMatches * 20;
        if (matchedTerms == queryTerms.Count)
        {
            score += 500;
        }

        if (lineTermMatches == queryTerms.Count)
        {
            score += 250;
        }

        return score;
    }

    private static string BuildSearchableText(string relativePath, string[] lines, int lineIndex)
    {
        int start = Math.Max(0, lineIndex - SearchContextWindowRadius);
        int end = Math.Min(lines.Length - 1, lineIndex + SearchContextWindowRadius);
        return relativePath + Environment.NewLine + string.Join(Environment.NewLine, lines.Skip(start).Take(end - start + 1));
    }

    private static int RequiredTermMatches(int termCount)
        => termCount switch
        {
            <= 1 => 1,
            TwoTermQueryCount => TwoTermQueryCount,
            ThreeTermQueryCount => MinimumMultiTermMatches,
            _ => termCount - 1,
        };

    private static bool IsSearchTermChar(char ch)
        => char.IsLetterOrDigit(ch) || ch is '_' or '-' or '#' or '.';

    private static bool IsStopWord(string term)
        => term.Length <= 1 || term.ToLowerInvariant() is
            "a" or "an" or "and" or "are" or "as" or "at" or "be" or "but" or "by" or
            "for" or "from" or "how" or "in" or "into" or "is" or "it" or "of" or "on" or
            "or" or "that" or "the" or "this" or "to" or "use" or "using" or "with";

    private static IEnumerable<string> EnumerateMemoryFiles(string root, bool includeRollouts)
    {
        string[] priorityFiles = ["memory_summary.md", "MEMORY.md"];
        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);

        foreach (string fileName in priorityFiles)
        {
            string fullPath = Path.Combine(root, fileName);
            if (File.Exists(fullPath) && seen.Add(fullPath))
            {
                yield return fullPath;
            }
        }

        foreach (string file in EnumerateFilesSafe(Path.Combine(root, "extensions"), "*.md", SearchOption.AllDirectories))
        {
            if (seen.Add(file))
            {
                yield return file;
            }
        }

        foreach (string file in EnumerateFilesSafe(Path.Combine(root, "skills"), "*.md", SearchOption.AllDirectories))
        {
            if (seen.Add(file))
            {
                yield return file;
            }
        }

        if (!includeRollouts)
        {
            yield break;
        }

        string rolloutDir = Path.Combine(root, "rollout_summaries");
        foreach (string file in EnumerateFilesSafe(rolloutDir, "*.md", SearchOption.TopDirectoryOnly)
            .Concat(EnumerateFilesSafe(rolloutDir, "*.jsonl", SearchOption.TopDirectoryOnly)))
        {
            if (seen.Add(file))
            {
                yield return file;
            }
        }
    }

    private static string[] EnumerateFilesSafe(string directory, string pattern, SearchOption searchOption)
    {
        if (!Directory.Exists(directory))
        {
            return [];
        }

        try
        {
            return [.. Directory.EnumerateFiles(directory, pattern, searchOption)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)];
        }
        catch (IOException)
        {
            return [];
        }
        catch (UnauthorizedAccessException)
        {
            return [];
        }
    }

    private static bool ShouldSkipSearchFile(string file)
    {
        try
        {
            return new FileInfo(file).Length > MaxSearchFileBytes;
        }
        catch (IOException)
        {
            return true;
        }
        catch (UnauthorizedAccessException)
        {
            return true;
        }
    }

    private static string ResolveMemoryRoot(JsonNode? id, BridgeConnection bridge)
    {
        List<string> candidates = [];
        AddMemoryRootCandidate(candidates, Environment.GetEnvironmentVariable("CODEX_HOME"));
        AddHomeCandidate(candidates, Environment.GetEnvironmentVariable("USERPROFILE"));
        AddHomeCandidate(candidates, Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
        AddHomeCandidate(candidates, InferUserHomeFromSolution(bridge));
        AddHomeCandidate(candidates, InferUserHomeFromDiscoveryFile(bridge));

        foreach (string candidate in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (Directory.Exists(candidate))
            {
                return Path.GetFullPath(candidate);
            }
        }

        string tried = string.Join(", ", candidates.Distinct(StringComparer.OrdinalIgnoreCase));
        throw new McpRequestException(id, McpErrorCodes.InvalidParams,
            $"Codex memory root was not found. Tried: {tried}");
    }

    private static void AddMemoryRootCandidate(List<string> candidates, string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        string fullPath = Path.GetFullPath(path);
        if (string.Equals(Path.GetFileName(fullPath), "memories", StringComparison.OrdinalIgnoreCase))
        {
            candidates.Add(fullPath);
        }
        else
        {
            candidates.Add(Path.Combine(fullPath, "memories"));
        }
    }

    private static void AddHomeCandidate(List<string> candidates, string? home)
    {
        if (!string.IsNullOrWhiteSpace(home))
        {
            candidates.Add(Path.Combine(home, ".codex", "memories"));
        }
    }

    private static string? InferUserHomeFromSolution(BridgeConnection bridge)
    {
        try
        {
            DirectoryInfo? directory = new(ServiceToolPaths.ResolveSolutionDirectory(bridge));
            while (directory is not null)
            {
                if (string.Equals(directory.Parent?.Name, "Users", StringComparison.OrdinalIgnoreCase))
                {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }
        }
        catch (Exception ex) when (ex is InvalidOperationException or IOException or UnauthorizedAccessException)
        {
            return null;
        }

        return null;
    }

    // The VSIX publishes its discovery JSON under the interactive user's temp directory
    // (C:\Users\<name>\AppData\Local\Temp\...). When this service runs as LocalSystem its
    // own profile is systemprofile and has no .codex, and the solution may live outside
    // C:\Users - a discovery path is then the most reliable pointer to the real user's home.
    private static string? InferUserHomeFromDiscoveryFile(BridgeConnection bridge)
    {
        string? home = TryInferUserHomeFromDiscoveryPath(bridge.CurrentInstance?.DiscoveryFile);
        if (home is not null)
        {
            return home;
        }

        // Not bound yet (a fresh HTTP session hits this): any discovered instance works,
        // because every VSIX publishes its discovery file under its own user profile.
        try
        {
            IReadOnlyList<BridgeInstance> instances = VsDiscovery.ListAsync(bridge.Mode)
                .ConfigureAwait(false)
                .GetAwaiter()
                .GetResult();
            foreach (BridgeInstance instance in instances)
            {
                home = TryInferUserHomeFromDiscoveryPath(instance.DiscoveryFile);
                if (home is not null)
                {
                    return home;
                }
            }
        }
        catch (BridgeException)
        {
            return null;
        }

        return null;
    }

    private static string? TryInferUserHomeFromDiscoveryPath(string? discoveryFile)
    {
        if (string.IsNullOrWhiteSpace(discoveryFile) || !Path.IsPathRooted(discoveryFile))
        {
            return null;
        }

        try
        {
            DirectoryInfo? directory = new FileInfo(discoveryFile).Directory;
            while (directory is not null)
            {
                if (string.Equals(directory.Parent?.Name, "Users", StringComparison.OrdinalIgnoreCase))
                {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }
        }
        catch (Exception ex) when (ex is ArgumentException or IOException or UnauthorizedAccessException or NotSupportedException)
        {
            return null;
        }

        return null;
    }

    private static string ResolveMemoryPath(JsonNode? id, string root, string relativePath)
    {
        if (string.IsNullOrWhiteSpace(relativePath))
        {
            relativePath = "MEMORY.md";
        }

        string fullPath = Path.GetFullPath(Path.IsPathRooted(relativePath)
            ? relativePath
            : Path.Combine(root, relativePath.Replace('/', Path.DirectorySeparatorChar)));
        string rootWithSeparator = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!fullPath.StartsWith(rootWithSeparator, StringComparison.OrdinalIgnoreCase))
        {
            throw new McpRequestException(id, McpErrorCodes.InvalidParams,
                "Memory path must stay under the resolved Codex memory root.");
        }

        if (!File.Exists(fullPath))
        {
            throw new McpRequestException(id, McpErrorCodes.InvalidParams,
                $"Memory file not found: {ToMemoryRelativePath(root, fullPath)}");
        }

        return fullPath;
    }

    private static string ToMemoryRelativePath(string root, string path)
        => Path.GetRelativePath(root, path).Replace(Path.DirectorySeparatorChar, '/');

    private static int Clamp(int value, int min, int max)
        => Math.Min(max, Math.Max(min, value));

    private static string Truncate(string value, int maxChars)
        => value.Length <= maxChars ? value : value[..maxChars] + "...";
}
