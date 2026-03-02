using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using VsIdeBridge.Infrastructure;

namespace VsIdeBridge.Services;

internal sealed class VsCommandService
{
    public async Task<JObject> ExecuteCommandAsync(DTE2 dte, string commandName, string? commandArgs)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var command = ResolveCommand(dte, commandName);
        try
        {
            dte.ExecuteCommand(command.Name, commandArgs ?? string.Empty);
        }
        catch (COMException ex)
        {
            throw new CommandErrorException(
                "unsupported_operation",
                $"Visual Studio command failed: {command.Name}",
                new { command = command.Name, args = commandArgs ?? string.Empty, exception = ex.Message, hresult = ex.HResult });
        }

        return CreateCommandInfo(command, commandArgs);
    }

    public async Task<JObject> ExecuteSymbolCommandAsync(
        DTE2 dte,
        DocumentService documentService,
        WindowService windowService,
        string[] candidateCommands,
        string? filePath,
        string? documentQuery,
        int? line,
        int? column,
        bool selectWord,
        string resultWindowQuery,
        bool activateResultWindow,
        int timeoutMs)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var location = await documentService
            .PositionTextSelectionAsync(dte, filePath, documentQuery, line, column, selectWord)
            .ConfigureAwait(true);

        Command? executed = null;
        string? commandError = null;
        foreach (var candidate in candidateCommands)
        {
            var command = TryResolveCommand(dte, candidate);
            if (command is null)
            {
                continue;
            }

            try
            {
                dte.ExecuteCommand(command.Name, string.Empty);
                executed = command;
                break;
            }
            catch (COMException ex)
            {
                commandError = ex.Message;
            }
        }

        if (executed is null)
        {
            throw new CommandErrorException(
                "unsupported_operation",
                $"None of the Visual Studio commands could be executed: {string.Join(", ", candidateCommands)}",
                new { candidates = candidateCommands, error = commandError ?? string.Empty });
        }

        var data = CreateCommandInfo(executed, string.Empty);
        data["location"] = location;
        data["candidateCommands"] = new JArray(candidateCommands);

        var window = await windowService.WaitForWindowAsync(
                dte,
                resultWindowQuery,
                activateResultWindow,
                Math.Max(0, timeoutMs))
            .ConfigureAwait(true);
        data["resultWindowQuery"] = resultWindowQuery;
        data["resultWindowActivated"] = window is not null;
        data["resultWindow"] = window;
        return data;
    }

    private static JObject CreateCommandInfo(Command command, string? commandArgs)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        return new JObject
        {
            ["command"] = command.Name,
            ["args"] = commandArgs ?? string.Empty,
            ["guid"] = command.Guid ?? string.Empty,
            ["id"] = command.ID,
            ["bindings"] = new JArray(ToStringArray(command.Bindings)),
        };
    }

    private static Command ResolveCommand(DTE2 dte, string commandName)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        var command = TryResolveCommand(dte, commandName);
        if (command is not null)
        {
            return command;
        }

        throw new CommandErrorException("unsupported_operation", $"Visual Studio command not found: {commandName}");
    }

    private static Command? TryResolveCommand(DTE2 dte, string commandName)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            return dte.Commands.Item(commandName, 0);
        }
        catch
        {
        }

        return dte.Commands
            .Cast<Command>()
            .FirstOrDefault(command => string.Equals(command.Name, commandName, StringComparison.OrdinalIgnoreCase));
    }
    private static string[] ToStringArray(object bindings)
    {
        if (bindings is object[] items)
        {
            return items.Select(item => item?.ToString() ?? string.Empty)
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .ToArray();
        }

        return Array.Empty<string>();
    }
}
