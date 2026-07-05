using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed partial class SearchService
{
    private static IEnumerable<(string Path, string ProjectUniqueName)> EnumerateSolutionFiles(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (dte.Solution?.IsOpen != true)
        {
            yield break;
        }

        // Snapshot the project list first. Iterating dte.Solution.Projects directly can throw
        // COMException from the enumerator itself for some project types in VS 18+ (.slnx).
        List<Project> projects = [];
        try
        {
            foreach (Project project in dte.Solution.Projects)
            {
                if (project is not null)
                {
                    projects.Add(project);
                }
            }
        }
        catch (COMException ex)
        {
            TraceSearchFailure(nameof(EnumerateSolutionFiles), ex);
        }

        foreach (Project project in projects)
        {
            foreach ((string Path, string ProjectUniqueName) file in EnumerateProjectFiles(project))
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

        // Guard every COM property access — VS 18 SDK-style / .slnx projects can throw
        // "Failed to load the document due to an internal error." from these properties.
        string? kind = null;
        try { kind = project.Kind; }
        catch (COMException ex) { TraceSearchFailure("EnumerateProjectFiles.Kind", ex); }

        if (kind is null)
        {
            yield break;
        }

        if (string.Equals(kind, EnvDTE80.ProjectKinds.vsProjectKindSolutionFolder, StringComparison.OrdinalIgnoreCase))
        {
            ProjectItems? sfProjectItems = null;
            try { sfProjectItems = project.ProjectItems; }
        catch (COMException ex) { TraceSearchFailure("EnumerateProjectFiles.SFProjectItems", ex); }

            foreach (ProjectItem item in TryGetProjectItems(sfProjectItems, "EnumerateProjectFiles.SFEnumerate"))
            {
                Project? subProject = null;
                try { subProject = item.SubProject; }
                catch (COMException ex) { TraceSearchFailure("EnumerateProjectFiles.SubProject", ex); }

                if (subProject is not null)
                {
                    foreach ((string Path, string ProjectUniqueName) file in EnumerateProjectFiles(subProject))
                    {
                        yield return file;
                    }
                }
            }

            yield break;
        }

        string uniqueName = string.Empty;
        try { uniqueName = project.UniqueName; }
        catch (COMException ex) { TraceSearchFailure("EnumerateProjectFiles.UniqueName", ex); }

        ProjectItems? projectItems = null;
        try { projectItems = project.ProjectItems; }
        catch (COMException ex) { TraceSearchFailure("EnumerateProjectFiles.ProjectItems", ex); }

        foreach (ProjectItem item in TryGetProjectItems(projectItems, "EnumerateProjectFiles.Enumerate"))
        {
            foreach ((string Path, string ProjectUniqueName) file in EnumerateProjectItemFiles(item, uniqueName))
            {
                yield return file;
            }
        }
    }

    private static IEnumerable<(string Path, string ProjectUniqueName)> EnumerateProjectItemFiles(ProjectItem item, string projectUniqueName)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (item is null)
        {
            yield break;
        }

        // item.FileCount and item.FileNames can throw for virtual or unresolvable items.
        short fileCount = 0;
        try { fileCount = item.FileCount; }
        catch (COMException ex) { TraceSearchFailure("EnumerateProjectItemFiles.FileCount", ex); }

        for (short i = 1; i <= fileCount; i++)
        {
            string? fileName = null;
            try { fileName = item.FileNames[i]; }
            catch (COMException ex) { TraceSearchFailure("EnumerateProjectItemFiles.FileNames", ex); }

            if (!string.IsNullOrWhiteSpace(fileName))
            {
                yield return (PathNormalization.NormalizeFilePath(fileName!), projectUniqueName);
            }
        }

        ProjectItems? children = null;
        try { children = item.ProjectItems; }
        catch (COMException ex) { TraceSearchFailure("EnumerateProjectItemFiles.ProjectItems", ex); }

        foreach (ProjectItem child in TryGetProjectItems(children, "EnumerateProjectItemFiles.Children"))
        {
            foreach ((string Path, string ProjectUniqueName) file in EnumerateProjectItemFiles(child, projectUniqueName))
            {
                yield return file;
            }
        }
    }

    /// <summary>
    /// Safely materialises a COM <see cref="ProjectItems"/> collection into a list.
    /// Catches any exception thrown by the enumerator itself (e.g. from VS 18 SDK-style projects).
    /// </summary>
    private static List<ProjectItem> TryGetProjectItems(ProjectItems? projectItems, string context)
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        List<ProjectItem> result = [];
        if (projectItems is null)
        {
            return result;
        }

        try
        {
            foreach (ProjectItem item in projectItems)
            {
                if (item is not null)
                {
                    result.Add(item);
                }
            }
        }
        catch (COMException ex)
        {
            TraceSearchFailure(context, ex);
        }

        return result;
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
            catch (Exception ex) when (IsRecoverableSearchFailure(ex))
            {
                TraceSearchFailure("EnumerateOpenDocumentTargets", ex);
            }

            if (string.IsNullOrWhiteSpace(fullName))
            {
                continue;
            }

            string normalizedPath = PathNormalization.NormalizeFilePath(fullName);
            yield return (normalizedPath, document.ProjectItem?.ContainingProject?.UniqueName ?? string.Empty);
        }
    }

    private static (string Path, string ProjectUniqueName) TryGetActiveDocumentTarget(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        Document? activeDocument = dte.ActiveDocument;
        if (activeDocument is null || string.IsNullOrWhiteSpace(activeDocument.FullName))
        {
            return (string.Empty, string.Empty);
        }

        return (PathNormalization.NormalizeFilePath(activeDocument.FullName), activeDocument.ProjectItem?.ContainingProject?.UniqueName ?? string.Empty);
    }

    private static string? NormalizeSearchPathFilter(DTE2 dte, string? pathFilter)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (string.IsNullOrWhiteSpace(pathFilter))
        {
            return null;
        }

        string trimmedPath = pathFilter!.Trim();
        if (Path.IsPathRooted(trimmedPath))
        {
            return PathNormalization.NormalizeFilePath(trimmedPath);
        }

        string normalizedRelativePath = trimmedPath.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
        string? solutionDirectory = TryGetSolutionDirectory(dte);
        if (!string.IsNullOrWhiteSpace(solutionDirectory))
        {
            string rootedCandidate = Path.GetFullPath(Path.Combine(solutionDirectory, normalizedRelativePath));
            // Root only concrete FILE references (used to target a specific document). A relative
            // DIRECTORY fragment must stay a substring filter: for out-of-tree builds the solution
            // directory is the build tree, and a filter like "src/libslic3r" resolving to an
            // existing build-default\src\libslic3r would become an exclusive rooted prefix that
            // silently drops every match under the real source tree.
            if (File.Exists(rootedCandidate))
            {
                return PathNormalization.NormalizeFilePath(rootedCandidate);
            }
        }

        return normalizedRelativePath;
    }

    private static bool MatchesPathFilter(string path, string? pathFilter)
    {
        if (string.IsNullOrWhiteSpace(pathFilter))
        {
            return true;
        }

        string normalizedPath = PathNormalization.NormalizeFilePath(path);
        if (!Path.IsPathRooted(pathFilter))
        {
            string normalizedFilter = pathFilter!.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
            return normalizedPath.IndexOf(normalizedFilter, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        if (Directory.Exists(pathFilter))
        {
            string normalizedDirectory = pathFilter!.TrimEnd('\\');
            return string.Equals(normalizedPath, normalizedDirectory, StringComparison.OrdinalIgnoreCase)
                || normalizedPath.StartsWith(normalizedDirectory + "\\", StringComparison.OrdinalIgnoreCase);
        }

        return string.Equals(normalizedPath, pathFilter, StringComparison.OrdinalIgnoreCase);
    }

    private static string? TryGetSolutionDirectory(DTE2 dte)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        string? solutionFullName = dte.Solution?.FullName;
        return string.IsNullOrWhiteSpace(solutionFullName)
            ? null
            : Path.GetDirectoryName(solutionFullName);
    }

    private static (string Path, string ProjectUniqueName) TryResolveExplicitDocumentTarget(DTE2 dte, string? pathFilter)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (string.IsNullOrWhiteSpace(pathFilter) || !Path.IsPathRooted(pathFilter) || !File.Exists(pathFilter))
        {
            return (string.Empty, string.Empty);
        }

        foreach ((string Path, string ProjectUniqueName) solutionFile in EnumerateSolutionFiles(dte))
        {
            if (string.Equals(solutionFile.Path, pathFilter, StringComparison.OrdinalIgnoreCase))
            {
                return solutionFile;
            }
        }

        foreach ((string Path, string ProjectUniqueName) openFile in EnumerateOpenFiles(dte))
        {
            if (string.Equals(openFile.Path, pathFilter, StringComparison.OrdinalIgnoreCase))
            {
                return openFile;
            }
        }

        return (pathFilter!, string.Empty);
    }

    private static string[] ReadSearchLines(DTE2 dte, string path)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        string normalizedPath = PathNormalization.NormalizeFilePath(path);
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
                    EditPoint editPoint = textDocument.StartPoint.CreateEditPoint();
                    string text = editPoint.GetText(textDocument.EndPoint);
                    return text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
                }
            }
            catch (Exception ex) when (IsRecoverableSearchFailure(ex))
            {
                TraceSearchFailure("ReadSearchLines", ex);
            }
        }

        return File.ReadAllLines(normalizedPath);
    }

    private static bool IsRecoverableSearchFailure(Exception ex)
    {
        return ex is COMException
            || string.Equals(ex.GetType().FullName, "Microsoft.Assumes+InternalErrorException", StringComparison.Ordinal)
            || ex.Message.IndexOf("Failed to load the document", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static void TraceSearchFailure(string operation, Exception ex)
    {
        System.Diagnostics.Debug.WriteLine($"VsIdeBridge.SearchService {operation}: {ex}");
    }
}
