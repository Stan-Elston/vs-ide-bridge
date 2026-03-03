using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed class FailureContextService
{
    private const int MaxSymbolFiles = 3;
    private const int SymbolMaxDepth = 4;
    private const int MaxErrorSymbolRows = 8;
    private const int MaxRelevantSymbolsPerRow = 5;

    public async Task<JObject> CaptureAsync(IdeCommandContext? context)
    {
        if (context is null)
        {
            return new JObject();
        }

        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(context.CancellationToken);

        var data = new JObject();
        JObject? state = null;
        JObject? openTabs = null;
        JObject? errorList = null;

        try
        {
            state = await context.Runtime.IdeStateService.GetStateAsync(context.Dte).ConfigureAwait(true);
            data["state"] = state;
        }
        catch
        {
        }

        try
        {
            openTabs = await context.Runtime.DocumentService.ListOpenTabsAsync(context.Dte).ConfigureAwait(true);
            data["openTabs"] = openTabs;
        }
        catch
        {
        }

        try
        {
            errorList = await context.Runtime.ErrorListService
                .GetErrorListAsync(context, waitForIntellisense: false, timeoutMilliseconds: 1500, quickSnapshot: true)
                .ConfigureAwait(true);
            data["errorList"] = errorList;
        }
        catch
        {
        }

        var outlineCache = new Dictionary<string, JObject>(StringComparer.OrdinalIgnoreCase);
        var symbolFiles = CollectSymbolFiles(state, errorList);
        if (symbolFiles.Count > 0)
        {
            var symbolContext = new JArray();
            foreach (var file in symbolFiles.Take(MaxSymbolFiles))
            {
                try
                {
                    var outline = await GetOutlineAsync(context, outlineCache, file).ConfigureAwait(true);
                    symbolContext.Add(new JObject
                    {
                        ["path"] = file,
                        ["outline"] = outline,
                    });
                }
                catch
                {
                }
            }

            if (symbolContext.Count > 0)
            {
                data["symbolContext"] = symbolContext;
            }
        }

        var errorSymbolContext = BuildErrorSymbolContext(errorList, outlineCache);
        if (errorSymbolContext.Count > 0)
        {
            data["errorSymbolContext"] = errorSymbolContext;
        }

        return data;
    }

    private static JArray BuildErrorSymbolContext(JObject? errorList, IReadOnlyDictionary<string, JObject> outlineCache)
    {
        var items = new JArray();
        if (errorList?["rows"] is not JArray rows)
        {
            return items;
        }

        foreach (var row in rows.OfType<JObject>().Take(MaxErrorSymbolRows))
        {
            var file = row["file"]?.Value<string>();
            var line = row["line"]?.Value<int>() ?? 0;
            if (string.IsNullOrWhiteSpace(file) || line <= 0)
            {
                continue;
            }

            var normalizedFile = PathNormalization.NormalizeFilePath(file);
            if (!outlineCache.TryGetValue(normalizedFile, out var outline))
            {
                continue;
            }

            var relevantSymbols = SelectRelevantSymbols(outline, line);
            if (relevantSymbols.Count == 0)
            {
                continue;
            }

            items.Add(new JObject
            {
                ["file"] = normalizedFile,
                ["line"] = line,
                ["column"] = row["column"] ?? 0,
                ["severity"] = row["severity"] ?? string.Empty,
                ["code"] = row["code"] ?? string.Empty,
                ["message"] = row["message"] ?? string.Empty,
                ["relevantSymbols"] = relevantSymbols,
            });
        }

        return items;
    }

    private static JArray SelectRelevantSymbols(JObject outline, int line)
    {
        if (outline["symbols"] is not JArray symbols)
        {
            return new JArray();
        }

        var containing = symbols
            .OfType<JObject>()
            .Where(symbol => ContainsLine(symbol, line))
            .OrderBy(symbol => (symbol["endLine"]?.Value<int>() ?? int.MaxValue) - (symbol["startLine"]?.Value<int>() ?? 0))
            .ThenBy(symbol => symbol["depth"]?.Value<int>() ?? int.MaxValue)
            .Take(MaxRelevantSymbolsPerRow)
            .ToList();

        if (containing.Count > 0)
        {
            return new JArray(containing.Select(CloneSymbol));
        }

        var nearby = symbols
            .OfType<JObject>()
            .Select(symbol => new
            {
                Symbol = symbol,
                Distance = DistanceFromLine(symbol, line),
            })
            .OrderBy(item => item.Distance)
            .ThenBy(item => item.Symbol["depth"]?.Value<int>() ?? int.MaxValue)
            .Take(MaxRelevantSymbolsPerRow)
            .Select(item => CloneSymbol(item.Symbol));

        return new JArray(nearby);
    }

    private static bool ContainsLine(JObject symbol, int line)
    {
        var startLine = symbol["startLine"]?.Value<int>() ?? 0;
        var endLine = symbol["endLine"]?.Value<int>() ?? startLine;
        return startLine > 0 && line >= startLine && line <= Math.Max(startLine, endLine);
    }

    private static int DistanceFromLine(JObject symbol, int line)
    {
        var startLine = symbol["startLine"]?.Value<int>() ?? int.MaxValue;
        var endLine = symbol["endLine"]?.Value<int>() ?? startLine;
        if (line < startLine)
        {
            return startLine - line;
        }

        if (line > endLine)
        {
            return line - endLine;
        }

        return 0;
    }

    private static JObject CloneSymbol(JObject symbol)
    {
        return new JObject
        {
            ["name"] = symbol["name"] ?? string.Empty,
            ["kind"] = symbol["kind"] ?? string.Empty,
            ["startLine"] = symbol["startLine"] ?? 0,
            ["endLine"] = symbol["endLine"] ?? 0,
            ["depth"] = symbol["depth"] ?? 0,
        };
    }

    private static async Task<JObject> GetOutlineAsync(
        IdeCommandContext context,
        IDictionary<string, JObject> outlineCache,
        string file)
    {
        if (outlineCache.TryGetValue(file, out var cached))
        {
            return cached;
        }

        var outline = await context.Runtime.DocumentService
            .GetFileOutlineAsync(context.Dte, file, SymbolMaxDepth, kindFilter: null)
            .ConfigureAwait(true);
        outlineCache[file] = outline;
        return outline;
    }

    private static List<string> CollectSymbolFiles(JObject? state, JObject? errorList)
    {
        var files = new List<string>();

        var activeDocument = state?["activeDocument"]?.Value<string>();
        if (!string.IsNullOrWhiteSpace(activeDocument))
        {
            files.Add(PathNormalization.NormalizeFilePath(activeDocument));
        }

        if (errorList?["rows"] is JArray rows)
        {
            foreach (var row in rows.OfType<JObject>())
            {
                var file = row["file"]?.Value<string>();
                if (string.IsNullOrWhiteSpace(file))
                {
                    continue;
                }

                files.Add(PathNormalization.NormalizeFilePath(file));
            }
        }

        return files
            .Where(file => !string.IsNullOrWhiteSpace(file))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }
}
