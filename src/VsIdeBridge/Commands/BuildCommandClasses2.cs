using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using System.ComponentModel.Design;
using System.Threading.Tasks;
using VsIdeBridge.Infrastructure;
using VsIdeBridge.Services;
using VsIdeBridge.Tooling.Handles;

namespace VsIdeBridge.Commands;

internal static partial class DebugBuildCommands
{
    private const string ConfigurationArgument = "configuration";
    private const string PlatformArgument = "platform";

    internal sealed class IdeBuildSolutionCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService) : IdeCommandBase(package, runtime, commandService, 0x0212)
    {
        protected override string CanonicalName => "Tools.IdeBuildSolution";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            int timeout = args.GetInt32(TimeoutMillisecondsArgument, DefaultBuildTimeoutMilliseconds);
            bool waitForCompletion = args.GetBoolean(WaitForCompletionArgument, false);
            await EnsureCleanDiagnosticsAsync(context, args, timeout, quickPreflight: !waitForCompletion).ConfigureAwait(true);

            string? project = args.GetString("project");
            JObject buildResult;
            if (!string.IsNullOrWhiteSpace(project))
            {
                if (!waitForCompletion)
                {
                    throw new CommandErrorException(
                        "invalid_arguments",
                        "wait_for_completion: false is not supported when a project is specified. " +
                        "To build a specific project, set wait_for_completion: true. " +
                        "To fire-and-forget a build, remove the project parameter to build the entire solution.");
                }

                await context.Runtime.BridgeApprovalService.RequestApprovalAsync(
                    context,
                    BridgeApprovalKind.Build,
                    subject: $"Build project '{project}'",
                    details: null).ConfigureAwait(true);

                buildResult = await context.Runtime.BuildService.BuildProjectAsync(
                    context,
                    timeout,
                    project!,
                    args.GetString(ConfigurationArgument),
                    args.GetString(PlatformArgument)).ConfigureAwait(true);
            }
            else
            {
                await context.Runtime.BridgeApprovalService.RequestApprovalAsync(
                    context,
                    BridgeApprovalKind.Build,
                    subject: "Build solution",
                    details: null).ConfigureAwait(true);

                buildResult = waitForCompletion
                    ? await context.Runtime.BuildService.BuildSolutionAsync(
                        context,
                        timeout,
                        args.GetString(ConfigurationArgument),
                        args.GetString(PlatformArgument)).ConfigureAwait(true)
                    : await context.Runtime.BuildService.StartBuildSolutionAsync(
                        context,
                        timeout,
                        args.GetString(ConfigurationArgument),
                        args.GetString(PlatformArgument)).ConfigureAwait(true);
            }

            if (!waitForCompletion)
            {
                return CreateStartedResult("Build", buildResult);
            }

            if (args.GetBoolean(RequireCleanDiagnosticsArgument, true))
            {
                JObject diagnostics = await GetDiagnosticsSnapshotAsync(context, args, timeout, waitForIntellisense: false).ConfigureAwait(true);
                ThrowIfBuildDiagnosticsPresent(diagnostics, args, buildResult, "Build completed but diagnostics remain");
            }

            return new CommandExecutionResult($"Build completed with LastBuildInfo={buildResult["lastBuildInfo"]}.", buildResult);
        }
    }

    internal sealed class IdeRebuildSolutionCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService) : IdeCommandBase(package, runtime, commandService, 0x0261)
    {
        protected override string CanonicalName => "Tools.IdeRebuildSolution";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            int timeout = args.GetInt32(TimeoutMillisecondsArgument, DefaultBuildTimeoutMilliseconds);
            bool waitForCompletion = args.GetBoolean(WaitForCompletionArgument, false);
            await EnsureCleanDiagnosticsAsync(context, args, timeout, quickPreflight: !waitForCompletion).ConfigureAwait(true);

            await context.Runtime.BridgeApprovalService.RequestApprovalAsync(
                context,
                BridgeApprovalKind.Build,
                subject: "Rebuild solution",
                details: null).ConfigureAwait(true);

            JObject rebuildResult = waitForCompletion
                ? await context.Runtime.BuildService.RebuildSolutionAsync(
                    context,
                    timeout,
                    args.GetString(ConfigurationArgument),
                    args.GetString(PlatformArgument)).ConfigureAwait(true)
                : await context.Runtime.BuildService.StartRebuildSolutionAsync(
                    context,
                    timeout,
                    args.GetString(ConfigurationArgument),
                    args.GetString(PlatformArgument)).ConfigureAwait(true);

            if (!waitForCompletion)
            {
                return CreateStartedResult("Rebuild", rebuildResult);
            }

            if (args.GetBoolean(RequireCleanDiagnosticsArgument, true))
            {
                JObject diagnostics = await GetDiagnosticsSnapshotAsync(context, args, timeout, waitForIntellisense: false).ConfigureAwait(true);
                ThrowIfBuildDiagnosticsPresent(diagnostics, args, rebuildResult, "Rebuild completed but diagnostics remain");
            }

            return new CommandExecutionResult($"Rebuild completed with LastBuildInfo={rebuildResult["lastBuildInfo"]}.", rebuildResult);
        }
    }

    internal sealed class IdeGetErrorListCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService) : IdeCommandBase(package, runtime, commandService, 0x0213)
    {
        protected override string CanonicalName => "Tools.IdeGetErrorList";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            bool quick = args.GetBoolean("quick", false);
            bool forceRefresh = GetDiagnosticsForceRefresh(args);
            JObject errorListResult = await GetDiagnosticsWithFallbackAsync(
                context,
                args.GetBoolean("wait-for-intellisense", !quick),
                args.GetInt32(TimeoutMillisecondsArgument, GetQuickDiagnosticsTimeout(quick)),
                quick,
                CreateErrorListQuery(context, args),
                forceRefresh).ConfigureAwait(true);

            context.Handles.RegisterDiagnosticRows(HandleKind.Error, (JArray)errorListResult["rows"]!);
            return new CommandExecutionResult(
                BuildDiagnosticsCountSummary(errorListResult),
                errorListResult,
                BuildDiagnosticsCrossReferenceAdvisories(errorListResult, "Warning", "Message"));
        }
    }

    internal sealed class IdeGetWarningsCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService) : IdeCommandBase(package, runtime, commandService, 0x0230)
    {
        protected override string CanonicalName => "Tools.IdeGetWarnings";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            bool quick = args.GetBoolean("quick", false);
            bool forceRefresh = GetDiagnosticsForceRefresh(args);
            JObject warningListResult = await GetDiagnosticsWithFallbackAsync(
                context,
                args.GetBoolean("wait-for-intellisense", !quick),
                args.GetInt32(TimeoutMillisecondsArgument, GetQuickDiagnosticsTimeout(quick)),
                quick,
                CreateErrorListQuery(context, args, "warning"),
                forceRefresh).ConfigureAwait(true);

            context.Handles.RegisterDiagnosticRows(HandleKind.Warning, (JArray)warningListResult["rows"]!);
            return new CommandExecutionResult(
                BuildDiagnosticsCountSummary(warningListResult),
                warningListResult,
                BuildDiagnosticsCrossReferenceAdvisories(warningListResult, "Error", "Message"));
        }
    }

    internal sealed class IdeGetMessagesCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService) : IdeCommandBase(package, runtime, commandService, 0x0263)
    {
        protected override string CanonicalName => "Tools.IdeGetMessages";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            bool quick = args.GetBoolean("quick", false);
            bool forceRefresh = GetDiagnosticsForceRefresh(args);
            JObject messageListResult = await GetDiagnosticsWithFallbackAsync(
                context,
                args.GetBoolean("wait-for-intellisense", !quick),
                args.GetInt32(TimeoutMillisecondsArgument, GetQuickDiagnosticsTimeout(quick)),
                quick,
                CreateErrorListQuery(context, args, "message"),
                forceRefresh).ConfigureAwait(true);

            context.Handles.RegisterDiagnosticRows(HandleKind.Message, (JArray)messageListResult["rows"]!);
            return new CommandExecutionResult(
                BuildDiagnosticsCountSummary(messageListResult),
                messageListResult,
                BuildDiagnosticsCrossReferenceAdvisories(messageListResult, "Error", "Warning"));
        }
    }

    internal sealed class IdeBuildAndCaptureErrorsCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService) : IdeCommandBase(package, runtime, commandService, 0x0214)
    {
        protected override string CanonicalName => "Tools.IdeBuildAndCaptureErrors";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            int timeout = GetBuildErrorsTimeout(args);
            await EnsureCleanDiagnosticsAsync(context, args, timeout).ConfigureAwait(true);
            JObject build = await context.Runtime.BuildService.BuildAndCaptureErrorsAsync(
                context,
                timeout,
                args.GetBoolean(WaitForIntellisenseArgument, true)).ConfigureAwait(true);
            JObject errors = await GetDiagnosticsWithFallbackAsync(
                context,
                waitForIntellisense: false,
                timeout,
                quickSnapshot: false,
                CreateErrorListQuery(context, args)).ConfigureAwait(true);

            JObject buildAndErrorsResult = new()
            {
                ["build"] = build,
                ["errors"] = errors,
            };

            if (args.GetBoolean(RequireCleanDiagnosticsArgument, true))
            {
                ThrowIfDiagnosticsPresent(errors, "Build completed but diagnostics remain", args, buildAndErrorsResult);
            }

            return new CommandExecutionResult($"Build finished and captured {errors["count"]} Error List row(s).", buildAndErrorsResult);
        }
    }

    private static void ThrowIfBuildDiagnosticsPresent(JObject diagnostics, CommandArguments args, JObject buildResult, string summaryPrefix)
    {
        JObject buildContext = new() { ["build"] = buildResult };
        ThrowIfDiagnosticsPresent(diagnostics, summaryPrefix, args, buildContext);
    }

    internal sealed class IdeRunCodeAnalysisCommand(VsIdeBridgePackage package, IdeBridgeRuntime runtime, OleMenuCommandService commandService) : IdeCommandBase(package, runtime, commandService, 0x0264)
    {
        protected override string CanonicalName => "Tools.IdeRunCodeAnalysis";

        protected override async Task<CommandExecutionResult> ExecuteAsync(IdeCommandContext context, CommandArguments args)
        {
            int timeout = args.GetInt32(TimeoutMillisecondsArgument, DefaultBuildTimeoutMilliseconds);
            bool waitForCompletion = args.GetBoolean(WaitForCompletionArgument, false);
            await context.Runtime.BridgeApprovalService.RequestApprovalAsync(
                context,
                BridgeApprovalKind.Build,
                subject: "Run code analysis on solution",
                details: null).ConfigureAwait(true);
            JObject analysisResult = waitForCompletion
                ? await context.Runtime.BuildService.RunCodeAnalysisAsync(context, timeout).ConfigureAwait(true)
                : await context.Runtime.BuildService.StartCodeAnalysisAsync(context, timeout).ConfigureAwait(true);

            if (!waitForCompletion)
            {
                return CreateStartedResult("Code analysis", analysisResult);
            }

            JObject errors = await GetDiagnosticsWithFallbackAsync(
                context,
                waitForIntellisense: false,
                timeout,
                quickSnapshot: false,
                CreateErrorListQuery(context, args)).ConfigureAwait(true);
            JObject result = new()
            {
                ["analysis"] = analysisResult,
                ["errors"] = errors,
            };
            return new CommandExecutionResult($"Code analysis finished and captured {errors["count"]} Error List row(s).", result);
        }
    }
}
