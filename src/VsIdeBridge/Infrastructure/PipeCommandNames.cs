using System;
using System.Collections.Generic;
using System.Text;

namespace VsIdeBridge.Infrastructure;

internal static class PipeCommandNames
{
    private static readonly Dictionary<string, string[]> PreferredAliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Tools.IdeHelp"] = new[] { "help" },
        ["Tools.IdeSmokeTest"] = new[] { "smoke-test" },
        ["Tools.IdeGetState"] = new[] { "state" },
        ["Tools.IdeWaitForReady"] = new[] { "ready" },
        ["Tools.IdeOpenSolution"] = new[] { "open-solution" },
        ["Tools.IdeCloseIde"] = new[] { "close-ide" },
        ["Tools.IdeBatchCommands"] = new[] { "batch" },
        ["Tools.IdeFindText"] = new[] { "find-text" },
        ["Tools.IdeFindFiles"] = new[] { "find-files" },
        ["Tools.IdeOpenDocument"] = new[] { "open-document" },
        ["Tools.IdeListDocuments"] = new[] { "list-documents" },
        ["Tools.IdeListOpenTabs"] = new[] { "list-tabs" },
        ["Tools.IdeActivateDocument"] = new[] { "activate-document" },
        ["Tools.IdeCloseDocument"] = new[] { "close-document" },
        ["Tools.IdeCloseFile"] = new[] { "close-file" },
        ["Tools.IdeCloseAllExceptCurrent"] = new[] { "close-others" },
        ["Tools.IdeActivateWindow"] = new[] { "activate-window" },
        ["Tools.IdeListWindows"] = new[] { "list-windows" },
        ["Tools.IdeExecuteVsCommand"] = new[] { "execute-command" },
        ["Tools.IdeFindAllReferences"] = new[] { "find-references" },
        ["Tools.IdeShowCallHierarchy"] = new[] { "call-hierarchy" },
        ["Tools.IdeGetDocumentSlice"] = new[] { "document-slice" },
        ["Tools.IdeGetSmartContextForQuery"] = new[] { "smart-context" },
        ["Tools.IdeApplyUnifiedDiff"] = new[] { "apply-diff" },
        ["Tools.IdeGoToDefinition"] = new[] { "goto-definition" },
        ["Tools.IdeGoToImplementation"] = new[] { "goto-implementation" },
        ["Tools.IdeGetFileOutline"] = new[] { "file-outline" },
        ["Tools.IdeSearchSymbols"] = new[] { "search-symbols" },
        ["Tools.IdeGetQuickInfo"] = new[] { "quick-info", "peek-definition" },
        ["Tools.IdeGetDocumentSlices"] = new[] { "document-slices" },
        ["Tools.IdeGetFileSymbols"] = new[] { "file-symbols" },
        ["Tools.IdeSetBreakpoint"] = new[] { "set-breakpoint" },
        ["Tools.IdeListBreakpoints"] = new[] { "list-breakpoints" },
        ["Tools.IdeRemoveBreakpoint"] = new[] { "remove-breakpoint" },
        ["Tools.IdeClearAllBreakpoints"] = new[] { "clear-breakpoints" },
        ["Tools.IdeEnableBreakpoint"] = new[] { "enable-breakpoint" },
        ["Tools.IdeDisableBreakpoint"] = new[] { "disable-breakpoint" },
        ["Tools.IdeEnableAllBreakpoints"] = new[] { "enable-all-breakpoints" },
        ["Tools.IdeDisableAllBreakpoints"] = new[] { "disable-all-breakpoints" },
        ["Tools.IdeDebugGetState"] = new[] { "debug-state" },
        ["Tools.IdeDebugStart"] = new[] { "debug-start" },
        ["Tools.IdeDebugStop"] = new[] { "debug-stop" },
        ["Tools.IdeDebugBreak"] = new[] { "debug-break" },
        ["Tools.IdeDebugContinue"] = new[] { "debug-continue" },
        ["Tools.IdeDebugStepOver"] = new[] { "debug-step-over" },
        ["Tools.IdeDebugStepInto"] = new[] { "debug-step-into" },
        ["Tools.IdeDebugStepOut"] = new[] { "debug-step-out" },
        ["Tools.IdeBuildSolution"] = new[] { "build" },
        ["Tools.IdeGetErrorList"] = new[] { "errors" },
        ["Tools.IdeGetWarnings"] = new[] { "warnings" },
        ["Tools.IdeBuildAndCaptureErrors"] = new[] { "build-errors" },
    };

    public static string GetPrimaryName(string canonicalName)
    {
        foreach (var alias in GetAliases(canonicalName))
        {
            return alias;
        }

        return canonicalName;
    }

    public static IReadOnlyList<string> GetAliases(string canonicalName)
    {
        var aliases = new List<string>();
        if (PreferredAliases.TryGetValue(canonicalName, out var preferred))
        {
            aliases.AddRange(preferred);
        }

        var generated = ToKebabAlias(canonicalName);
        if (!string.IsNullOrWhiteSpace(generated) &&
            !aliases.Exists(alias => string.Equals(alias, generated, StringComparison.OrdinalIgnoreCase)))
        {
            aliases.Add(generated);
        }

        return aliases;
    }

    private static string ToKebabAlias(string canonicalName)
    {
        var suffix = canonicalName;
        if (suffix.StartsWith("Tools.Ide", StringComparison.OrdinalIgnoreCase))
        {
            suffix = suffix.Substring("Tools.Ide".Length);
        }
        else if (suffix.StartsWith("Tools.VsIdeBridge", StringComparison.OrdinalIgnoreCase))
        {
            suffix = suffix.Substring("Tools.VsIdeBridge".Length);
        }
        else if (suffix.StartsWith("Tools.", StringComparison.OrdinalIgnoreCase))
        {
            suffix = suffix.Substring("Tools.".Length);
        }

        if (string.IsNullOrWhiteSpace(suffix))
        {
            return string.Empty;
        }

        var builder = new StringBuilder();
        for (var i = 0; i < suffix.Length; i++)
        {
            var ch = suffix[i];
            if (char.IsUpper(ch) && i > 0)
            {
                builder.Append('-');
            }

            builder.Append(char.ToLowerInvariant(ch));
        }

        return builder.ToString();
    }
}
