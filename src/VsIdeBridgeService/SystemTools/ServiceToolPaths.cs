using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using LibGit2Sharp;

namespace VsIdeBridgeService.SystemTools;

internal static class ServiceToolPaths
{
    // Cache the repo root per solution directory so ResolveRepoRootDirectory doesn't
    // re-run FindGitRoot + HasTrackedFilesUnder (LibGit2Sharp index scan) on every
    // git tool call in a session.  The repo root for a given solution doesn't change
    // at runtime, so this is safe to cache for the process lifetime.
    private static readonly ConcurrentDictionary<string, string> _repoRootCache =
        new(StringComparer.OrdinalIgnoreCase);
    public static string ResolveInstalledCompanionPath(string fileName, string? solutionPath = null)
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            throw new ArgumentException("Companion file name is required.", nameof(fileName));
        }

        string installedCandidate = Path.Combine(AppContext.BaseDirectory, fileName);
        if (File.Exists(installedCandidate))
        {
            return installedCandidate;
        }

        string? solutionDirectory = TryGetSolutionDirectory(solutionPath);
        if (string.IsNullOrWhiteSpace(solutionDirectory))
        {
            return installedCandidate;
        }

        string[] candidates =
        [
            Path.Combine(solutionDirectory, "src", "VsIdeBridgeLauncher", "bin", "Release", "net472", fileName),
            Path.Combine(solutionDirectory, "src", "VsIdeBridgeLauncher", "bin", "Release", fileName),
            Path.Combine(solutionDirectory, "src", "VsIdeBridgeLauncher", "bin", "Debug", "net472", fileName),
            Path.Combine(solutionDirectory, "src", "VsIdeBridgeLauncher", "bin", "Debug", fileName),
        ];

        foreach (string candidate in candidates)
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return installedCandidate;
    }

    public static string ResolveSolutionDirectory(BridgeConnection bridge)
    {
        string? solutionDirectory = TryGetCurrentSolutionDirectory(bridge);
        if (!string.IsNullOrWhiteSpace(solutionDirectory))
        {
            return solutionDirectory;
        }

        solutionDirectory = TryDiscoverSolutionDirectory(bridge);
        if (!string.IsNullOrWhiteSpace(solutionDirectory))
        {
            return solutionDirectory;
        }

        return Environment.CurrentDirectory;
    }

    private static string? TryGetCurrentSolutionDirectory(BridgeConnection bridge)
    {
        string? solutionDirectory = TryGetSolutionDirectory(bridge.CurrentSolutionPath);
        if (!string.IsNullOrWhiteSpace(solutionDirectory))
        {
            return solutionDirectory;
        }

        BridgeInstance? currentInstance = bridge.CurrentInstance;
        if (currentInstance is not null)
        {
            solutionDirectory = TryGetSolutionDirectory(currentInstance.SolutionPath);
            if (!string.IsNullOrWhiteSpace(solutionDirectory))
            {
                return solutionDirectory;
            }
        }

        return null;
    }

    private static string? TryDiscoverSolutionDirectory(BridgeConnection bridge)
    {
        try
        {
            BridgeInstance instance = VsDiscovery.SelectAsync(bridge.CurrentSelector, bridge.Mode)
                .ConfigureAwait(false)
                .GetAwaiter()
                .GetResult();
            return TryGetSolutionDirectory(instance.SolutionPath);
        }
        catch (BridgeException)
        {
            return null;
        }
    }

    /// <summary>
    /// Returns the root directory of the git repository that contains the current or discoverable solution.
    /// Walks up from the solution directory looking for a <c>.git</c> directory or file.
    /// If the nearest git root is an ancestor directory and the solution has no committed files
    /// tracked under it (i.e. it lives inside a parent mega-repo as an untracked subdirectory),
    /// the solution directory itself is returned so that git tools are scoped correctly.
    /// Falls back to the solution directory when no git root is found at all.
    /// </summary>
    public static string ResolveRepoRootDirectory(BridgeConnection bridge)
    {
        string? solutionDirectory = TryGetCurrentSolutionDirectory(bridge)
            ?? TryDiscoverSolutionDirectory(bridge);
        if (string.IsNullOrWhiteSpace(solutionDirectory))
        {
            throw new McpRequestException(null, McpErrorCodes.BridgeError,
                "Repository-scoped bridge tools require a bound or uniquely discoverable Visual Studio solution. " +
                "Call list_instances and bind_solution or bind_instance before running this tool.");
        }

        return _repoRootCache.GetOrAdd(solutionDirectory, ComputeRepoRoot);
    }

    /// <summary>
    /// Resolves the repo root for <paramref name="solutionDirectory"/> without caching.
    /// Called at most once per unique solution directory per process lifetime.
    /// </summary>
    private static string ComputeRepoRoot(string solutionDirectory)
    {
        string? repoRoot = FindGitRoot(solutionDirectory);
        if (repoRoot is null)
            return solutionDirectory;

        // If the found repo root IS the solution directory, use it directly.
        if (string.Equals(
                repoRoot.TrimEnd(Path.DirectorySeparatorChar),
                solutionDirectory.TrimEnd(Path.DirectorySeparatorChar),
                StringComparison.OrdinalIgnoreCase))
        {
            return repoRoot;
        }

        // The repo root is an ancestor - only accept it when the solution directory has at
        // least one file tracked in that repo.  If nothing is tracked there, the solution is
        // an untracked subdirectory of a parent mega-repo; return the solution directory so
        // that git operations are scoped to the right place.
        return HasTrackedFilesUnder(repoRoot, solutionDirectory) ? repoRoot : solutionDirectory;
    }

    /// <summary>
    /// Returns <see langword="true"/> when at least one file under <paramref name="directory"/>
    /// is present in the index of the repository rooted at <paramref name="repoRoot"/>.
    /// Swallows exceptions and returns <see langword="true"/> on failure so that callers
    /// default to using the found repo root rather than silently discarding it.
    /// </summary>
    private static bool HasTrackedFilesUnder(string repoRoot, string directory)
    {
        try
        {
            string relativePath = Path.GetRelativePath(repoRoot, directory).Replace('\\', '/');

            // GetRelativePath returns "." when the paths are equal - already handled above.
            // A ".." prefix means directory is outside repoRoot, which should not happen;
            // treat it as tracked to avoid a confusing fallback.
            if (relativePath == "." || relativePath.StartsWith("..", StringComparison.Ordinal))
                return true;

            string prefix = relativePath.TrimEnd('/') + "/";
            using Repository repo = new(repoRoot);
            return repo.Index.Any(e =>
                e.Path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
        }
        catch
        {
            return true; // fail-open: use the repo root if we can't inspect the index
        }
    }

    /// <summary>
    /// Walks up from <paramref name="startDirectory"/> to find the directory that contains
    /// a <c>.git</c> entry (directory for normal repos, file for git worktrees).
    /// Returns <see langword="null"/> when no git root is found within a reasonable depth.
    /// </summary>
    public static string? FindGitRoot(string startDirectory)
    {
        string current = startDirectory;
        for (int depth = 0; depth < 10 && !string.IsNullOrWhiteSpace(current); depth++)
        {
            if (Directory.Exists(Path.Combine(current, ".git")) ||
                File.Exists(Path.Combine(current, ".git")))
            {
                return current;
            }

            string parent = Path.GetDirectoryName(current) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(parent) ||
                string.Equals(parent, current, StringComparison.OrdinalIgnoreCase))
            {
                break;
            }

            current = parent;
        }

        return null;
    }

    internal static string? TryGetSolutionDirectory(string? solutionPath)
    {
        if (string.IsNullOrWhiteSpace(solutionPath))
        {
            return null;
        }

        string fullPath = Path.GetFullPath(solutionPath);
        if (Directory.Exists(fullPath))
        {
            return Path.TrimEndingDirectorySeparator(fullPath);
        }

        string? solutionDirectory = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrWhiteSpace(solutionDirectory))
        {
            return null;
        }

        return solutionDirectory;
    }
}
