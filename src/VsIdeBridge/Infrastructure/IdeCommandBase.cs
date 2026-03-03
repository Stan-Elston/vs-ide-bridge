using System;
using System.ComponentModel.Design;
using System.IO;
using System.Threading.Tasks;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Newtonsoft.Json.Linq;
using VsIdeBridge.Services;

namespace VsIdeBridge.Infrastructure;

internal abstract class IdeCommandBase
{
    private readonly VsIdeBridgePackage _package;
    private readonly IdeBridgeRuntime _runtime;
    protected OleMenuCommand MenuCommand { get; }

    protected IdeCommandBase(
        VsIdeBridgePackage package,
        IdeBridgeRuntime runtime,
        OleMenuCommandService commandService,
        int commandId,
        bool acceptsParameters = true)
    {
        _package = package;
        _runtime = runtime;

        var menuCommandId = new CommandID(CommandRegistrar.CommandSet, commandId);
        var menuCommand = new OleMenuCommand(Execute, menuCommandId);
        if (acceptsParameters)
        {
            menuCommand.ParametersDescription = "$";
        }

        commandService.AddCommand(menuCommand);
        MenuCommand = menuCommand;
    }

    protected abstract string CanonicalName { get; }

    protected VsIdeBridgePackage Package => _package;

    protected IdeBridgeRuntime Runtime => _runtime;

    internal string Name => CanonicalName;

    internal Task<CommandExecutionResult> ExecuteDirectAsync(IdeCommandContext ctx, CommandArguments args)
        => ExecuteAsync(ctx, args);

    protected abstract Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args);

    private void Execute(object sender, EventArgs e)
    {
        _ = _package.JoinableTaskFactory.RunAsync(() => ExecuteInternalAsync(e));
    }

    private async Task ExecuteInternalAsync(EventArgs e)
    {
        var startedAt = DateTimeOffset.UtcNow;
        var rawArguments = (e as OleMenuCmdEventArgs)?.InValue as string;
        var outputPath = string.Empty;
        string? requestId = null;

        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(_package.DisposalToken);
        var dte = await _package.GetServiceAsync(typeof(SDTE)).ConfigureAwait(true) as DTE2;
        Assumes.Present(dte);

        var context = new IdeCommandContext(_package, dte, _runtime.Logger, _runtime, _package.DisposalToken);

        try
        {
            var args = CommandArgumentParser.Parse(rawArguments);
            outputPath = ResolveOutputPath(args);
            requestId = args.GetString("request-id");
            var result = await ExecuteAsync(context, args).ConfigureAwait(true);
            var envelope = new CommandEnvelope
            {
                SchemaVersion = JsonSchemaVersioning.CurrentSchemaVersion,
                Command = CanonicalName,
                RequestId = requestId,
                Success = true,
                StartedAtUtc = startedAt.UtcDateTime.ToString("O"),
                FinishedAtUtc = DateTimeOffset.UtcNow.UtcDateTime.ToString("O"),
                Summary = result.Summary,
                Warnings = result.Warnings,
                Error = null,
                Data = result.Data,
            };

            await CommandResultWriter.WriteAsync(outputPath, envelope, _package.DisposalToken).ConfigureAwait(false);
            await context.Logger.LogAsync($"IDE Bridge: {CanonicalName} OK - {result.Summary} -> {outputPath}", _package.DisposalToken, activatePane: true).ConfigureAwait(true);
        }
        catch (CommandErrorException ex)
        {
            var failureData = await _runtime.FailureContextService.CaptureAsync(context).ConfigureAwait(true);
            var envelope = new CommandEnvelope
            {
                SchemaVersion = JsonSchemaVersioning.CurrentSchemaVersion,
                Command = CanonicalName,
                RequestId = requestId,
                Success = false,
                StartedAtUtc = startedAt.UtcDateTime.ToString("O"),
                FinishedAtUtc = DateTimeOffset.UtcNow.UtcDateTime.ToString("O"),
                Summary = ex.Message,
                Warnings = new JArray(),
                Error = new
                {
                    code = ex.Code,
                    message = ex.Message,
                    details = ex.Details,
                },
                Data = failureData,
            };

            if (!string.IsNullOrWhiteSpace(outputPath))
            {
                await CommandResultWriter.WriteAsync(outputPath, envelope, _package.DisposalToken).ConfigureAwait(false);
            }

            await context.Logger.LogAsync($"IDE Bridge: {CanonicalName} FAIL - {ex.Code}", _package.DisposalToken, activatePane: true).ConfigureAwait(true);
            ActivityLog.LogError(nameof(VsIdeBridgePackage), ex.ToString());
        }
        catch (Exception ex)
        {
            var failureData = await _runtime.FailureContextService.CaptureAsync(context).ConfigureAwait(true);
            var envelope = new CommandEnvelope
            {
                SchemaVersion = JsonSchemaVersioning.CurrentSchemaVersion,
                Command = CanonicalName,
                RequestId = requestId,
                Success = false,
                StartedAtUtc = startedAt.UtcDateTime.ToString("O"),
                FinishedAtUtc = DateTimeOffset.UtcNow.UtcDateTime.ToString("O"),
                Summary = ex.Message,
                Warnings = new JArray(),
                Error = new
                {
                    code = "internal_error",
                    message = ex.Message,
                    details = new { exception = ex.ToString() },
                },
                Data = failureData,
            };

            if (!string.IsNullOrWhiteSpace(outputPath))
            {
                await CommandResultWriter.WriteAsync(outputPath, envelope, _package.DisposalToken).ConfigureAwait(false);
            }

            await context.Logger.LogAsync($"IDE Bridge: {CanonicalName} FAIL - internal_error", _package.DisposalToken, activatePane: true).ConfigureAwait(true);
            ActivityLog.LogError(nameof(VsIdeBridgePackage), ex.ToString());
        }
    }

    private string ResolveOutputPath(CommandArguments args)
    {
        var explicitPath = args.GetString("out");
        if (!string.IsNullOrWhiteSpace(explicitPath))
        {
            return explicitPath!;
        }

        var fileName = CanonicalName.Replace("Tools.", string.Empty)
            .Replace('.', '-')
            .ToLowerInvariant() + ".json";
        return Path.Combine(Path.GetTempPath(), "vs-ide-bridge", fileName);
    }
}
