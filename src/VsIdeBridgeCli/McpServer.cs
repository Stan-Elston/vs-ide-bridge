using System.Diagnostics;
using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using VsIdeBridge.Shared;

namespace VsIdeBridgeCli;

internal static partial class CliApp
{
    private static readonly JsonSerializerOptions McpJsonOptions = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

    private static async Task<int> RunMcpServerAsync(CliOptions options)
    {
        await McpServer.RunAsync(options).ConfigureAwait(false);
        return 0;
    }

    private static class McpServer
    {
        private const string McpLog = @"C:\Temp\mcp-server.log";
        private const string DotNetExecutableName = "dotnet";
        private const string GitExecutableName = "git";
        private const string CondaExecutableName = "conda";
        private static readonly byte[] RawJsonTerminator = [(byte)'\n'];
        private const string ServiceControlPipeName = "VsIdeBridgeServiceControl";
        private static readonly byte[] HeaderTerminator = "\r\n\r\n"u8.ToArray();
        private static readonly string[] SupportedProtocolVersions = ["2025-03-26", "2024-11-05"];
        private static readonly string[] CondaExecutableExtensions = [".exe", ".cmd", ".bat", string.Empty];
        private static readonly string[] CondaRelativeCandidatePaths =
        [
            @"anaconda3\Scripts\conda.exe",
            @"miniconda3\Scripts\conda.exe",
            @"Miniconda3\Scripts\conda.exe",
        ];

        private enum McpWireFormat
        {
            HeaderFramed,
            RawJson,
        }

        private sealed class McpIncomingMessage
        {
            public JsonObject Request { get; set; } = null!;

            public McpWireFormat WireFormat { get; set; }
        }

        private sealed class BridgeBinding(CliOptions options)
        {
            private readonly CliOptions _options = options;
            private readonly bool _verbose = options.GetFlag("verbose");
            private BridgeInstanceSelector _selector = BridgeInstanceSelector.FromOptions(options);
            private PipeDiscovery? _cachedDiscovery;

            public async Task<JsonObject> SendAsync(JsonNode? id, string command, string args)
            {
                try
                {
                    return await SendCoreAsync(command, args).ConfigureAwait(false);
                }
                catch (CliException ex)
                {
                    throw new McpRequestException(id, -32001, ex.Message);
                }
                catch (TimeoutException ex)
                {
                    throw new McpRequestException(id, -32002, $"Timed out waiting for Visual Studio bridge response: {ex.Message}");
                }
                catch (IOException ex) when (_cachedDiscovery is not null)
                {
                    McpTrace($"cached pipe '{_cachedDiscovery.PipeName}' failed, refreshing binding: {ex.Message}");
                    _cachedDiscovery = null;

                    try
                    {
                        return await SendCoreAsync(command, args).ConfigureAwait(false);
                    }
                    catch (CliException retryEx)
                    {
                        throw new McpRequestException(id, -32001, retryEx.Message);
                    }
                    catch (TimeoutException retryEx)
                    {
                        throw new McpRequestException(id, -32002, $"Timed out waiting for Visual Studio bridge response: {retryEx.Message}");
                    }
                    catch (IOException retryEx)
                    {
                        throw new McpRequestException(id, -32003, $"Failed communicating with Visual Studio bridge pipe: {retryEx.Message}");
                    }
                }
                catch (IOException ex)
                {
                    throw new McpRequestException(id, -32003, $"Failed communicating with Visual Studio bridge pipe: {ex.Message}");
                }
            }

            public async Task<JsonObject> BindAsync(JsonNode? id, JsonObject? args)
            {
                _selector = CreateSelector(args);
                _cachedDiscovery = null;

                try
                {
                    var discovery = await GetDiscoveryAsync().ConfigureAwait(false);
                    return new JsonObject
                    {
                        ["success"] = true,
                        ["binding"] = DiscoveryToJson(discovery),
                        ["selector"] = SelectorToJson(_selector),
                    };
                }
                catch (CliException ex)
                {
                    throw new McpRequestException(id, -32001, ex.Message);
                }
            }

            public PipeDiscovery? CurrentDiscovery => _cachedDiscovery;
            public BridgeInstanceSelector CurrentSelector => _selector;
            public DiscoveryMode DiscoveryMode => ResolveDiscoveryMode(_options);

            public void PreferSolution(string? solutionHint)
            {
                _selector = new BridgeInstanceSelector
                {
                    InstanceId = _selector.InstanceId,
                    ProcessId = _selector.ProcessId,
                    PipeName = _selector.PipeName,
                    SolutionHint = solutionHint,
                };
                _cachedDiscovery = null;
                McpTrace($"selector updated to prefer solution hint '{solutionHint ?? string.Empty}'.");
            }

            private async Task<JsonObject> SendCoreAsync(string command, string args)
            {
                var discovery = await GetDiscoveryAsync().ConfigureAwait(false);
                await using var client = new PipeClient(discovery.PipeName, _options.GetInt32("timeout-ms", 130_000));
                var request = new JsonObject
                {
                    ["id"] = Guid.NewGuid().ToString("N")[..8],
                    ["command"] = command,
                    ["args"] = args,
                };

                return await client.SendAsync(request).ConfigureAwait(false);
            }

            private async Task<PipeDiscovery> GetDiscoveryAsync()
            {
                if (_cachedDiscovery is not null)
                {
                    return _cachedDiscovery;
                }

                var discovery = await PipeDiscovery
                    .SelectAsync(_selector, _verbose, ResolveDiscoveryMode(_options))
                    .ConfigureAwait(false);
                _cachedDiscovery = discovery;
                McpTrace($"bound instance={discovery.InstanceId} pipe={discovery.PipeName} source={discovery.Source} solution={discovery.SolutionPath}");
                return discovery;
            }

            private static BridgeInstanceSelector CreateSelector(JsonObject? args)
            {
                return new BridgeInstanceSelector
                {
                    InstanceId = GetString(args, "instance_id") ?? GetString(args, "instance"),
                    ProcessId = GetInt32(args, "pid"),
                    PipeName = GetString(args, "pipe_name") ?? GetString(args, "pipe"),
                    SolutionHint = GetString(args, "solution_hint") ?? GetString(args, "sln"),
                };
            }

            private static string? GetString(JsonObject? args, string name)
            {
                var value = args?[name]?.GetValue<string>();
                return string.IsNullOrWhiteSpace(value) ? null : value;
            }

            private static int? GetInt32(JsonObject? args, string name)
            {
                return args?[name]?.GetValue<int?>();
            }
        }

        private static void McpTrace(string msg)
        {
            try { File.AppendAllText(McpLog, $"{DateTime.Now:O} {msg}\n"); } catch { }
        }

        public static async Task RunAsync(CliOptions options)
        {
            try
            {
                File.WriteAllText(McpLog, $"{DateTime.Now:O} mcp-server started\n");
            }
            catch
            {
            }

            var input = Console.OpenStandardInput();
            var output = Console.OpenStandardOutput();
            var bridgeBinding = new BridgeBinding(options);
            var advertiseExtraCapabilities = !options.GetFlag("tools-only");
            var wireFormat = McpWireFormat.HeaderFramed;
            McpTrace("stdin/stdout opened");
            NotifyService("CLIENT_CONNECTED");
            while (true)
            {
                JsonObject? response;

                try
                {
                    McpTrace("waiting for next message...");
                    var incoming = await ReadMessageAsync(input).ConfigureAwait(false);
                    if (incoming is null)
                    {
                        McpTrace("null request (EOF) — exiting");
                        NotifyService("CLIENT_DISCONNECTED");
                        return;
                    }

                    wireFormat = incoming.WireFormat;
                    var request = incoming.Request;
                    var method = request["method"]?.GetValue<string>() ?? "(null)";
                    NotifyService("MCP_REQUEST");
                    McpTrace($"got request method={method}");

                    var trackInFlight = string.Equals(method, "tools/call", StringComparison.Ordinal);
                    if (trackInFlight)
                    {
                        NotifyService("COMMAND_START");
                    }

                    try
                    {
                        response = await HandleRequestAsync(request, bridgeBinding, advertiseExtraCapabilities).ConfigureAwait(false);
                        McpTrace($"handled method={method} response={(response is null ? "null" : "ok")}");
                    }
                    finally
                    {
                        if (trackInFlight)
                        {
                            NotifyService("COMMAND_END");
                        }
                    }
                }
                catch (McpRequestException ex)
                {
                    McpTrace($"McpRequestException: {ex.Message}");
                    response = CreateErrorResponse(ex.Id, ex.Code, ex.Message);
                }
                catch (JsonException ex)
                {
                    McpTrace($"JsonException: {ex.Message}");
                    response = CreateErrorResponse(null, -32700, $"Parse error: {ex.Message}");
                }
                catch (Exception ex)
                {
                    McpTrace($"Exception: {ex}");
                    response = CreateErrorResponse(null, -32603, $"Internal error: {ex.Message}");
                }

                if (response is not null)
                {
                    McpTrace("writing response...");
                    await WriteMessageAsync(output, response, wireFormat).ConfigureAwait(false);
                    McpTrace("response written");
                }
            }
        }

        private static void NotifyService(string evt)
        {
            if (string.IsNullOrWhiteSpace(evt))
            {
                return;
            }

            try
            {
                using var pipe = new NamedPipeClientStream(
                    ".",
                    ServiceControlPipeName,
                    PipeDirection.Out,
                    PipeOptions.Asynchronous);

                pipe.Connect(100);
                using var writer = new StreamWriter(pipe, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), leaveOpen: true)
                {
                    AutoFlush = true,
                };

                writer.WriteLine(evt);
            }
            catch
            {
                // Service host is optional. MCP server must continue even when service pipe is unavailable.
            }
        }

        private static async Task<JsonObject?> HandleRequestAsync(JsonObject request, BridgeBinding bridgeBinding, bool advertiseExtraCapabilities)
        {
            var id = request["id"]?.DeepClone();
            var method = request["method"]?.GetValue<string>() ?? string.Empty;
            var @params = request["params"] as JsonObject;

            if (method.StartsWith("notifications/", StringComparison.Ordinal))
                return null;

            JsonNode result = method switch
            {
                "initialize" => InitializeResult(@params, advertiseExtraCapabilities),
                "tools/list" => new JsonObject { ["tools"] = ListTools() },
                "tools/call" => await CallToolAsync(id, @params, bridgeBinding).ConfigureAwait(false),
                "resources/list" => new JsonObject { ["resources"] = ListResources() },
                "resources/templates/list" => new JsonObject { ["resourceTemplates"] = ListResourceTemplates() },
                "resources/read" => await ReadResourceAsync(id, @params, bridgeBinding).ConfigureAwait(false),
                "prompts/list" => new JsonObject { ["prompts"] = ListPrompts() },
                "prompts/get" => GetPrompt(id, @params),
                "ping" => new JsonObject(),
                _ => throw new McpRequestException(id, -32601, $"Unsupported MCP method: {method}"),
            };

            return new JsonObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = result };
        }

        private static JsonObject InitializeResult(JsonObject? @params, bool advertiseExtraCapabilities)
        {
            var capabilities = new JsonObject
            {
                ["tools"] = new JsonObject(),
            };

            if (advertiseExtraCapabilities)
            {
                capabilities["resources"] = new JsonObject { ["subscribe"] = false };
                capabilities["prompts"] = new JsonObject();
            }

            return new JsonObject
            {
                ["protocolVersion"] = SelectProtocolVersion(@params?["protocolVersion"]?.GetValue<string>()),
                ["capabilities"] = capabilities,
                ["serverInfo"] = new JsonObject
                {
                    ["name"] = "vs-ide-bridge-mcp",
                    ["version"] = "0.1.0",
                },
            };
        }

        private static string SelectProtocolVersion(string? clientProtocolVersion)
        {
            if (string.IsNullOrWhiteSpace(clientProtocolVersion))
            {
                return SupportedProtocolVersions[0];
            }

            if (SupportedProtocolVersions.Contains(clientProtocolVersion, StringComparer.Ordinal))
            {
                return clientProtocolVersion;
            }

            // Compatibility fallback for clients that require an exact echo.
            McpTrace($"initialize: client requested unsupported protocolVersion={clientProtocolVersion}; echoing for compatibility.");
            return clientProtocolVersion;
        }

        private static JsonArray ListTools() =>
        [
            Tool("state", "Capture current Visual Studio bridge state.", EmptySchema()),
            Tool("ready", "Wait for Visual Studio/IntelliSense readiness before semantic diagnostics.", EmptySchema()),
            Tool(
                "tool_help",
                "Return MCP tool help with descriptions, schemas, and examples. Pass name for one tool or omit for all.",
                ObjectSchema(("name", StringSchema("Optional tool name for focused help."), false))),
            Tool(
                "help",
                "Alias for tool_help. Return MCP tool help with descriptions, schemas, and examples.",
                ObjectSchema(("name", StringSchema("Optional tool name for focused help."), false))),
            Tool("bridge_health", "Get binding health, discovery source, and last round-trip metrics.", EmptySchema()),
            Tool("list_instances", "List live VS IDE Bridge instances visible to this MCP server.", EmptySchema()),
            Tool(
                "bind_instance",
                "Bind this MCP session to one Visual Studio bridge instance.",
                ObjectSchema(
                    ("instance_id", StringSchema("Optional exact bridge instance id."), false),
                    ("pid", IntegerSchema("Optional Visual Studio process id."), false),
                    ("pipe_name", StringSchema("Optional exact bridge pipe name."), false),
                    ("solution_hint", StringSchema("Optional solution path or name substring."), false))),
            Tool(
                "bind_solution",
                "Bind this MCP session to the Visual Studio bridge instance whose solution matches a name or path hint.",
                ObjectSchema(
                    ("solution", StringSchema("Solution name or path substring to match."), true))),
            Tool(
                "errors",
                "Get current errors.",
                ObjectSchema(
                    ("wait_for_intellisense", BooleanSchema("Wait for IntelliSense readiness first (default true)."), false),
                    ("quick", BooleanSchema("Read current snapshot immediately without stability polling (default false)."), false),
                    ("max", IntegerSchema("Optional max rows."), false),
                    ("code", StringSchema("Optional diagnostic code prefix filter."), false),
                    ("project", StringSchema("Optional project filter."), false),
                    ("path", StringSchema("Optional path filter."), false),
                    ("text", StringSchema("Optional message text filter."), false))),
            Tool(
                "warnings",
                "Get current warnings.",
                ObjectSchema(
                    ("wait_for_intellisense", BooleanSchema("Wait for IntelliSense readiness first (default true)."), false),
                    ("quick", BooleanSchema("Read current snapshot immediately without stability polling (default false)."), false),
                    ("max", IntegerSchema("Optional max rows."), false),
                    ("code", StringSchema("Optional diagnostic code prefix filter."), false),
                    ("project", StringSchema("Optional project filter."), false),
                    ("path", StringSchema("Optional path filter."), false),
                    ("text", StringSchema("Optional message text filter."), false))),
            Tool("list_tabs", "List open editor tabs.", EmptySchema()),
            Tool(
                "open_file",
                "Open an absolute path, solution-relative path, or solution item name and optional line/column.",
                ObjectSchema(
                    ("file", StringSchema("Absolute path, solution-relative path, or solution item name."), true),
                    ("line", IntegerSchema("Optional 1-based line number."), false),
                    ("column", IntegerSchema("Optional 1-based column number."), false),
                    ("allow_disk_fallback", BooleanSchema("Allow disk fallback under solution root when solution items do not match (default true)."), false))),
            Tool(
                "find_files",
                "Search solution explorer files by name or path fragment.",
                ObjectSchema(
                    ("query", StringSchema("File name or path fragment."), true),
                    ("path", StringSchema("Optional path fragment filter."), false),
                    ("extensions", ArrayOfStringsSchema("Optional extension filters like ['.cmake','.txt']."), false),
                    ("max_results", IntegerSchema("Optional max result count (default 200)."), false),
                    ("include_non_project", BooleanSchema("Include disk files under solution root that are not in projects (default true)."), false))),
            Tool(
                "search_symbols",
                "Search solution symbols by query.",
                ObjectSchema(
                    ("query", StringSchema("Symbol search text."), true),
                    ("kind", StringSchema("Optional symbol kind filter."), false))),
            Tool(
                "count_references",
                "Count symbol references at file/line/column with exact-or-explicit semantics.",
                ObjectSchema(
                    ("file", StringSchema("Absolute or solution-relative file path."), true),
                    ("line", IntegerSchema("1-based line number."), true),
                    ("column", IntegerSchema("1-based column number."), true),
                    ("activate_window", BooleanSchema("Activate references window while counting (default true)."), false),
                    ("timeout_ms", IntegerSchema("Optional window wait timeout in milliseconds."), false))),
            Tool(
                "quick_info",
                "Get quick info at file/line/column.",
                ObjectSchema(
                    ("file", StringSchema("Absolute or solution-relative file path."), true),
                    ("line", IntegerSchema("1-based line number."), true),
                    ("column", IntegerSchema("1-based column number."), true))),
            Tool(
                "apply_diff",
                "Apply unified diff through Visual Studio editor buffer. Changes are saved and opened automatically.",
                ObjectSchema(
                    ("patch", StringSchema("Unified diff text."), true),
                    ("post_check", BooleanSchema("If true, run ready and errors after applying diff."), false))),
            Tool("debug_threads", "Get debugger thread snapshot.", EmptySchema()),
            Tool(
                "debug_stack",
                "Get debugger stack frames for current or selected thread.",
                ObjectSchema(
                    ("thread_id", IntegerSchema("Optional debugger thread id."), false),
                    ("max_frames", IntegerSchema("Optional max frames (default 100)."), false))),
            Tool(
                "debug_locals",
                "Get local variables for the current stack frame.",
                ObjectSchema(
                    ("max", IntegerSchema("Optional max locals (default 200)."), false))),
            Tool("debug_modules", "Get debugger modules snapshot (best effort).", EmptySchema()),
            Tool(
                "debug_watch",
                "Evaluate one watch expression in break mode.",
                ObjectSchema(
                    ("expression", StringSchema("Debugger watch expression."), true),
                    ("timeout_ms", IntegerSchema("Optional evaluation timeout milliseconds (default 1000)."), false))),
            Tool("debug_exceptions", "Get debugger exception settings snapshot (best effort).", EmptySchema()),
            Tool(
                "diagnostics_snapshot",
                "Capture IDE/debug/build state and errors/warnings in one response.",
                ObjectSchema(
                    ("wait_for_intellisense", BooleanSchema("Wait for IntelliSense before diagnostics (default true)."), false),
                    ("quick", BooleanSchema("Use quick diagnostics snapshot mode (default false)."), false),
                    ("max", IntegerSchema("Optional max diagnostics rows for errors/warnings."), false))),
            Tool("build_configurations", "List solution build configurations/platforms.", EmptySchema()),
            Tool(
                "set_build_configuration",
                "Activate one build configuration/platform pair.",
                ObjectSchema(
                    ("configuration", StringSchema("Build configuration name (e.g. Debug)."), true),
                    ("platform", StringSchema("Optional platform (e.g. x64)."), false))),
            Tool("git_status", "Get repository status in porcelain mode.", EmptySchema()),
            Tool("git_current_branch", "Get current branch name (short).", EmptySchema()),
            Tool("git_remote_list", "List configured remotes with URLs.", EmptySchema()),
            Tool("git_tag_list", "List tags sorted by version-aware refname.", EmptySchema()),
            Tool("git_stash_list", "List stash entries.", EmptySchema()),
            Tool(
                "git_diff_unstaged",
                "Show unstaged diff with optional context lines.",
                ObjectSchema(("context", IntegerSchema("Optional context line count (default 3)."), false))),
            Tool(
                "git_diff_staged",
                "Show staged diff with optional context lines.",
                ObjectSchema(("context", IntegerSchema("Optional context line count (default 3)."), false))),
            Tool(
                "git_log",
                "Show recent commits in a compact machine-friendly format.",
                ObjectSchema(("max_count", IntegerSchema("Optional max commit count (default 20)."), false))),
            Tool(
                "git_show",
                "Show metadata and patch for a specific commit.",
                ObjectSchema(("revision", StringSchema("Commit-ish (hash, HEAD~1, tag)."), true))),
            Tool("git_branch_list", "List local and remote branches.", EmptySchema()),
            Tool(
                "git_checkout",
                "Checkout an existing branch or revision.",
                ObjectSchema(("target", StringSchema("Branch name or revision to checkout."), true))),
            Tool(
                "git_create_branch",
                "Create and switch to a new branch.",
                ObjectSchema(("name", StringSchema("New branch name."), true), ("start_point", StringSchema("Optional start point (default HEAD)."), false))),
            Tool(
                "git_add",
                "Stage files. Use ['.'] to stage all changes.",
                ObjectSchema(("paths", ArrayOfStringsSchema("File paths, globs, or '.' to stage all."), true))),
            Tool(
                "git_restore",
                "Restore paths from HEAD in the working tree.",
                ObjectSchema(("paths", ArrayOfStringsSchema("File paths or globs to restore."), true))),
            Tool(
                "git_commit",
                "Create a commit from staged changes.",
                ObjectSchema(("message", StringSchema("Commit message."), true))),
            Tool(
                "git_commit_amend",
                "Amend the previous commit. Optionally replace the message.",
                ObjectSchema(("message", StringSchema("Optional replacement commit message."), false), ("no_edit", BooleanSchema("If true, keep the current commit message."), false))),
            Tool(
                "git_reset",
                "Unstage paths while keeping working tree changes.",
                ObjectSchema(("paths", ArrayOfStringsSchema("File paths or globs to unstage."), false))),
            Tool(
                "git_fetch",
                "Fetch updates from remotes.",
                ObjectSchema(("remote", StringSchema("Optional remote name."), false), ("all", BooleanSchema("If true, fetch all remotes."), false), ("prune", BooleanSchema("If true, prune deleted remote refs (default true)."), false))),
            Tool(
                "git_stash_push",
                "Stash local changes. Optionally include untracked files and a message.",
                ObjectSchema(("message", StringSchema("Optional stash message."), false), ("include_untracked", BooleanSchema("If true, include untracked files."), false))),
            Tool("git_stash_pop", "Apply and drop the latest stash entry.", EmptySchema()),
            Tool(
                "git_pull",
                "Pull updates from a remote branch.",
                ObjectSchema(("remote", StringSchema("Optional remote name (default current tracking remote)."), false), ("branch", StringSchema("Optional branch name."), false))),
            Tool(
                "git_push",
                "Push current branch to remote.",
                ObjectSchema(("remote", StringSchema("Optional remote name (default current tracking remote)."), false), ("branch", StringSchema("Optional branch name."), false), ("set_upstream", BooleanSchema("If true, pass --set-upstream."), false))),
            Tool(
                "github_issue_search",
                "Search open or closed GitHub issues.",
                ObjectSchema(("query", StringSchema("Free-text search query."), false), ("state", StringSchema("open, closed, or all."), false), ("repo", StringSchema("Optional owner/repo. Defaults to git origin repo."), false), ("limit", IntegerSchema("Max results (default 20)."), false))),
            Tool(
                "github_issue_close",
                "Close a GitHub issue by number and optionally add a comment.",
                ObjectSchema(("issue_number", IntegerSchema("Issue number to close."), true), ("repo", StringSchema("Optional owner/repo. Defaults to git origin repo."), false), ("comment", StringSchema("Optional closing comment."), false))),
            Tool(
                "nuget_restore",
                "Restore NuGet packages with dotnet restore for the active solution or a specific path.",
                ObjectSchema(("path", StringSchema("Optional solution/project path. Defaults to the active bridge solution."), false))),
            Tool(
                "nuget_add_package",
                "Add a NuGet package reference to a project via dotnet add package.",
                ObjectSchema(
                    ("project", StringSchema("Project path (.csproj/.fsproj/.vbproj), absolute or solution-relative."), true),
                    ("package", StringSchema("NuGet package id to add."), true),
                    ("version", StringSchema("Optional package version."), false),
                    ("source", StringSchema("Optional package source (name or URL)."), false),
                    ("prerelease", BooleanSchema("If true, include prerelease versions."), false),
                    ("no_restore", BooleanSchema("If true, skip restore after adding the package."), false))),
            Tool(
                "nuget_remove_package",
                "Remove a NuGet package reference from a project via dotnet remove package.",
                ObjectSchema(
                    ("project", StringSchema("Project path (.csproj/.fsproj/.vbproj), absolute or solution-relative."), true),
                    ("package", StringSchema("NuGet package id to remove."), true))),
            Tool(
                "conda_install",
                "Install one or more packages into a conda environment.",
                ObjectSchema(
                    ("packages", ArrayOfStringsSchema("One or more conda package specs (for example ['numpy','cmake>=3.29'])."), true),
                    ("name", StringSchema("Optional environment name (-n/--name)."), false),
                    ("prefix", StringSchema("Optional environment prefix path (--prefix)."), false),
                    ("channels", ArrayOfStringsSchema("Optional channels to add with --channel."), false),
                    ("dry_run", BooleanSchema("If true, run with --dry-run."), false),
                    ("yes", BooleanSchema("Auto-confirm install (default true)."), false))),
            Tool(
                "conda_remove",
                "Remove one or more packages from a conda environment.",
                ObjectSchema(
                    ("packages", ArrayOfStringsSchema("One or more package names to remove."), true),
                    ("name", StringSchema("Optional environment name (-n/--name)."), false),
                    ("prefix", StringSchema("Optional environment prefix path (--prefix)."), false),
                    ("dry_run", BooleanSchema("If true, run with --dry-run."), false),
                    ("yes", BooleanSchema("Auto-confirm remove (default true)."), false))),
            Tool(
                "find_text",
                "Full-text search across the solution or a path subtree. Returns file paths, line numbers and preview text.",
                ObjectSchema(
                    ("query", StringSchema("Search text or regex pattern."), true),
                    ("path", StringSchema("Optional path or directory filter (solution-relative or absolute)."), false),
                    ("scope", StringSchema("Scope: solution (default), project, or document."), false),
                    ("match_case", BooleanSchema("Case-sensitive match (default false)."), false),
                    ("whole_word", BooleanSchema("Match whole word only (default false)."), false))),
            Tool(
                "read_file",
                "Read lines from a file. Use start_line/end_line for a range, or line with context_before/context_after centered on an anchor.",
                ObjectSchema(
                    ("file", StringSchema("Absolute or solution-relative file path."), true),
                    ("start_line", IntegerSchema("First 1-based line to read. Use with end_line for a range."), false),
                    ("end_line", IntegerSchema("Last 1-based line to read (inclusive). Use with start_line."), false),
                    ("line", IntegerSchema("Anchor 1-based line. Use with context_before/context_after."), false),
                    ("context_before", IntegerSchema("Lines before anchor (default 10)."), false),
                    ("context_after", IntegerSchema("Lines after anchor (default 30)."), false),
                    ("reveal_in_editor", BooleanSchema("Whether to reveal the slice in the editor (default true)."), false))),
            Tool(
                "find_references",
                "Find all references to the symbol at file/line/column using VS IntelliSense.",
                ObjectSchema(
                    ("file", StringSchema("Absolute or solution-relative file path."), true),
                    ("line", IntegerSchema("1-based line number."), true),
                    ("column", IntegerSchema("1-based column number."), true))),
            Tool(
                "peek_definition",
                "Return the definition source and surrounding context of the symbol at file/line/column.",
                ObjectSchema(
                    ("file", StringSchema("Absolute or solution-relative file path."), true),
                    ("line", IntegerSchema("1-based line number."), true),
                    ("column", IntegerSchema("1-based column number."), true))),
            Tool(
                "file_outline",
                "Get the symbol outline (namespaces, classes, methods, fields, etc.) of a file.",
                ObjectSchema(
                    ("file", StringSchema("Absolute or solution-relative file path."), true))),
            Tool(
                "build",
                "Trigger a solution build and return errors/warnings. Builds may take several minutes.",
                ObjectSchema(
                    ("configuration", StringSchema("Optional build configuration (e.g. Debug, Release)."), false),
                    ("platform", StringSchema("Optional build platform (e.g. x64)."), false))),
            Tool(
                "open_solution",
                "Open a solution file in the current Visual Studio instance without opening a new window.",
                ObjectSchema(
                    ("solution", StringSchema("Absolute path to the .sln file to open."), true),
                    ("wait_for_ready", BooleanSchema("Wait for readiness after opening the solution (default true)."), false)))
        ];

        private static JsonObject Tool(string name, string description, JsonObject inputSchema) => new()
        {
            ["name"] = name,
            ["description"] = ResolveToolDescription(name, description),
            ["inputSchema"] = inputSchema,
        };

        private static string ResolveToolDescription(string toolName, string fallback)
        {
            var bridgeCommand = ResolveBridgeCommandForTool(toolName);
            if (string.IsNullOrWhiteSpace(bridgeCommand))
            {
                return fallback;
            }

            return BridgeCommandCatalog.TryGetByPipeName(bridgeCommand, out var metadata)
                ? metadata.Description
                : fallback;
        }

        private static async Task<JsonNode> CallToolAsync(JsonNode? id, JsonObject? p, BridgeBinding bridgeBinding)
        {
            var toolName = p?["name"]?.GetValue<string>() ?? throw new McpRequestException(id, -32602, "tools/call missing name.");
            var args = p?["arguments"] as JsonObject;

            if (string.Equals(toolName, "tool_help", StringComparison.Ordinal) ||
                string.Equals(toolName, "help", StringComparison.Ordinal))
            {
                return ToolHelp(id, args?["name"]?.GetValue<string>());
            }

            if (string.Equals(toolName, "bridge_health", StringComparison.Ordinal))
            {
                return await BridgeHealthAsync(id, bridgeBinding).ConfigureAwait(false);
            }

            if (toolName.StartsWith("git_", StringComparison.Ordinal))
            {
                return await CallGitToolAsync(id, toolName, args, bridgeBinding).ConfigureAwait(false);
            }

            if (toolName.StartsWith("github_", StringComparison.Ordinal))
            {
                return await CallGitHubToolAsync(id, toolName, args, bridgeBinding).ConfigureAwait(false);
            }

            if (toolName.StartsWith("nuget_", StringComparison.Ordinal))
            {
                return await CallNuGetToolAsync(id, toolName, args, bridgeBinding).ConfigureAwait(false);
            }

            if (toolName.StartsWith("conda_", StringComparison.Ordinal))
            {
                return await CallCondaToolAsync(id, toolName, args, bridgeBinding).ConfigureAwait(false);
            }

            if (string.Equals(toolName, "list_instances", StringComparison.Ordinal))
            {
                return await ListInstancesAsync(bridgeBinding).ConfigureAwait(false);
            }

            if (string.Equals(toolName, "bind_instance", StringComparison.Ordinal))
            {
                var result = await bridgeBinding.BindAsync(id, args).ConfigureAwait(false);
                return new JsonObject
                {
                    ["content"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["type"] = "text",
                            ["text"] = result.ToJsonString(JsonOptions),
                        },
                    },
                    ["isError"] = !(result["success"]?.GetValue<bool>() ?? false),
                    ["structuredContent"] = result,
                };
            }

            if (string.Equals(toolName, "bind_solution", StringComparison.Ordinal))
            {
                var bindArgs = new JsonObject
                {
                    ["solution_hint"] = args?["solution"]?.DeepClone() ?? args?["solution_hint"]?.DeepClone(),
                };

                var result = await bridgeBinding.BindAsync(id, bindArgs).ConfigureAwait(false);
                return new JsonObject
                {
                    ["content"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["type"] = "text",
                            ["text"] = result.ToJsonString(JsonOptions),
                        },
                    },
                    ["isError"] = !(result["success"]?.GetValue<bool>() ?? false),
                    ["structuredContent"] = result,
                };
            }

            if (string.Equals(toolName, "open_solution", StringComparison.Ordinal))
            {
                return await OpenSolutionAsync(id, args, bridgeBinding).ConfigureAwait(false);
            }

            var (command, commandArgs) = toolName switch
            {
                "state" => ("state", string.Empty),
                "ready" => ("ready", string.Empty),
                "errors" => ("errors", BuildArgs(
                    ("wait-for-intellisense", GetBoolean(args, "wait_for_intellisense", true) ? "true" : "false"),
                    ("quick", GetBoolean(args, "quick", false) ? "true" : "false"),
                    ("max", args?["max"]?.ToString()),
                    ("code", args?["code"]?.GetValue<string>()),
                    ("project", args?["project"]?.GetValue<string>()),
                    ("path", args?["path"]?.GetValue<string>()),
                    ("text", args?["text"]?.GetValue<string>()))),
                "warnings" => ("warnings", BuildArgs(
                    ("wait-for-intellisense", GetBoolean(args, "wait_for_intellisense", true) ? "true" : "false"),
                    ("quick", GetBoolean(args, "quick", false) ? "true" : "false"),
                    ("max", args?["max"]?.ToString()),
                    ("code", args?["code"]?.GetValue<string>()),
                    ("project", args?["project"]?.GetValue<string>()),
                    ("path", args?["path"]?.GetValue<string>()),
                    ("text", args?["text"]?.GetValue<string>()))),
                "list_tabs" => ("list-tabs", string.Empty),
                "open_file" => ("open-document", BuildArgs(
                    ("file", args?["file"]?.GetValue<string>()),
                    ("line", args?["line"]?.ToString()),
                    ("column", args?["column"]?.ToString()),
                    ("allow-disk-fallback", GetBoolean(args, "allow_disk_fallback", true) ? "true" : "false"))),
                "find_files" => ("find-files", BuildArgs(
                    ("query", args?["query"]?.GetValue<string>()),
                    ("path", args?["path"]?.GetValue<string>()),
                    ("extensions", GetCsv(args?["extensions"] as JsonArray)),
                    ("max-results", args?["max_results"]?.ToString()),
                    ("include-non-project", GetBoolean(args, "include_non_project", true) ? "true" : "false"))),
                "search_symbols" => ("search-symbols", BuildArgs(("query", args?["query"]?.GetValue<string>()), ("kind", args?["kind"]?.GetValue<string>()))),
                "count_references" => ("count-references", BuildArgs(
                    ("file", args?["file"]?.GetValue<string>()),
                    ("line", args?["line"]?.ToString()),
                    ("column", args?["column"]?.ToString()),
                    ("activate-window", GetBoolean(args, "activate_window", true) ? "true" : "false"),
                    ("timeout-ms", args?["timeout_ms"]?.ToString()))),
                "quick_info" => ("quick-info", BuildArgs(("file", args?["file"]?.GetValue<string>()), ("line", args?["line"]?.ToString()), ("column", args?["column"]?.ToString()))),
                "apply_diff" => ("apply-diff", BuildArgs(("patch-text-base64", Convert.ToBase64String(Encoding.UTF8.GetBytes(args?["patch"]?.GetValue<string>() ?? string.Empty))), ("open-changed-files", "true"), ("save-changed-files", "true"))),
                "debug_threads" => ("debug-threads", string.Empty),
                "debug_stack" => ("debug-stack", BuildArgs(("thread-id", args?["thread_id"]?.ToString()), ("max-frames", args?["max_frames"]?.ToString()))),
                "debug_locals" => ("debug-locals", BuildArgs(("max", args?["max"]?.ToString()))),
                "debug_modules" => ("debug-modules", string.Empty),
                "debug_watch" => ("debug-watch", BuildArgs(("expression", args?["expression"]?.GetValue<string>()), ("timeout-ms", args?["timeout_ms"]?.ToString()))),
                "debug_exceptions" => ("debug-exceptions", string.Empty),
                "diagnostics_snapshot" => ("diagnostics-snapshot", BuildArgs(
                    ("wait-for-intellisense", GetBoolean(args, "wait_for_intellisense", true) ? "true" : "false"),
                    ("quick", GetBoolean(args, "quick", false) ? "true" : "false"),
                    ("max", args?["max"]?.ToString()))),
                "build_configurations" => ("build-configurations", string.Empty),
                "set_build_configuration" => ("set-build-configuration", BuildArgs(("configuration", args?["configuration"]?.GetValue<string>()), ("platform", args?["platform"]?.GetValue<string>()))),
                "find_text" => ("find-text", BuildArgs(("query", args?["query"]?.GetValue<string>()), ("path", args?["path"]?.GetValue<string>()), ("scope", args?["scope"]?.GetValue<string>()), ("match-case", args?["match_case"]?.GetValue<bool>() == true ? "true" : null), ("whole-word", args?["whole_word"]?.GetValue<bool>() == true ? "true" : null))),
                "read_file" => ("document-slice", BuildArgs(
                    ("file", args?["file"]?.GetValue<string>()),
                    ("start-line", args?["start_line"]?.ToString()),
                    ("end-line", args?["end_line"]?.ToString()),
                    ("line", args?["line"]?.ToString()),
                    ("context-before", args?["context_before"]?.ToString()),
                    ("context-after", args?["context_after"]?.ToString()),
                    ("reveal-in-editor", GetBoolean(args, "reveal_in_editor", true) ? "true" : "false"))),
                "find_references" => ("find-references", BuildArgs(("file", args?["file"]?.GetValue<string>()), ("line", args?["line"]?.ToString()), ("column", args?["column"]?.ToString()))),
                "peek_definition" => ("peek-definition", BuildArgs(("file", args?["file"]?.GetValue<string>()), ("line", args?["line"]?.ToString()), ("column", args?["column"]?.ToString()))),
                "file_outline" => ("file-outline", BuildArgs(("file", args?["file"]?.GetValue<string>()))),
                "build" => ("build", BuildArgs(("configuration", args?["configuration"]?.GetValue<string>()), ("platform", args?["platform"]?.GetValue<string>()))),
                _ => throw new McpRequestException(id, -32602, $"Unknown MCP tool: {toolName}"),
            };

            var response = await SendBridgeAsync(id, bridgeBinding, command, commandArgs).ConfigureAwait(false);

            if (string.Equals(toolName, "apply_diff", StringComparison.Ordinal) && GetBoolean(args, "post_check", false))
            {
                var ready = await SendBridgeAsync(id, bridgeBinding, "ready", string.Empty).ConfigureAwait(false);
                var errors = await SendBridgeAsync(id, bridgeBinding, "errors", "--wait-for-intellisense true").ConfigureAwait(false);
                response["postCheck"] = new JsonObject
                {
                    ["ready"] = ready,
                    ["errors"] = errors,
                };
            }

            return new JsonObject
            {
                ["content"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["type"] = "text",
                        ["text"] = response.ToJsonString(JsonOptions),
                    },
                },
                ["isError"] = !ResponseFormatter.IsSuccess(response),
                ["structuredContent"] = response.DeepClone(),
            };
        }

        private static JsonObject ToolHelp(JsonNode? id, string? toolName)
        {
            var tools = ListTools()
                .OfType<JsonObject>()
                .ToArray();

            if (!string.IsNullOrWhiteSpace(toolName))
            {
                var match = tools.FirstOrDefault(tool =>
                    string.Equals(tool["name"]?.GetValue<string>(), toolName, StringComparison.Ordinal))
                    ?? throw new McpRequestException(id, -32602, $"Unknown MCP tool: {toolName}");

                var item = BuildToolHelpEntry(match);
                var result = new JsonObject
                {
                    ["count"] = 1,
                    ["items"] = new JsonArray { item },
                };
                return WrapToolResult(result, isError: false);
            }

            var entries = new JsonArray();
            foreach (var tool in tools.OrderBy(item => item["name"]?.GetValue<string>(), StringComparer.Ordinal))
            {
                entries.Add(BuildToolHelpEntry(tool));
            }

            var catalog = new JsonObject
            {
                ["count"] = entries.Count,
                ["items"] = entries,
            };

            return WrapToolResult(catalog, isError: false);
        }

        private static JsonObject BuildToolHelpEntry(JsonObject tool)
        {
            var name = tool["name"]?.GetValue<string>() ?? string.Empty;
            var inputSchema = tool["inputSchema"] as JsonObject ?? EmptySchema();
            var bridgeCommand = ResolveBridgeCommandForTool(name);
            BridgeCommandMetadata? bridgeMetadata = null;
            if (!string.IsNullOrWhiteSpace(bridgeCommand))
            {
                if (BridgeCommandCatalog.TryGetByPipeName(bridgeCommand, out var commandMetadata))
                {
                    bridgeMetadata = commandMetadata;
                }
            }

            var hasBridgeMetadata = bridgeMetadata is not null;
            var description = hasBridgeMetadata
                ? bridgeMetadata!.Description
                : tool["description"]?.GetValue<string>() ?? string.Empty;

            string? bridgeCommandValue = hasBridgeMetadata ? bridgeMetadata!.PipeName : null;
            string? bridgeExampleValue = hasBridgeMetadata ? bridgeMetadata!.Example : null;
            return new JsonObject
            {
                ["name"] = name,
                ["description"] = description,
                ["inputSchema"] = inputSchema.DeepClone(),
                ["example"] = GetToolExample(name, inputSchema),
                ["bridgeCommand"] = bridgeCommandValue,
                ["bridgeExample"] = bridgeExampleValue,
            };
        }

        private static JsonObject WrapToolResult(JsonObject structuredContent, bool isError)
        {
            return new JsonObject
            {
                ["content"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["type"] = "text",
                        ["text"] = structuredContent.ToJsonString(JsonOptions),
                    },
                },
                ["isError"] = isError,
                ["structuredContent"] = structuredContent.DeepClone(),
            };
        }

        private static string GetToolExample(string name, JsonObject inputSchema)
        {
            var overrideExample = name switch
            {
                "bind_solution" => "{ \"solution\": \"VsIdeBridge.sln\" }",
                "help" => "{ \"name\": \"open_file\" }",
                "tool_help" => "{ \"name\": \"open_file\" }",
                "open_solution" => "{ \"solution\": \"C:\\\\repo\\\\VsIdeBridge.sln\", \"wait_for_ready\": true }",
                "open_file" => "{ \"file\": \"src\\\\VsIdeBridgeCli\\\\Program.cs\", \"line\": 1 }",
                "find_files" => "{ \"query\": \"CMakeLists.txt\", \"include_non_project\": true }",
                "errors" => "{ \"wait_for_intellisense\": true, \"quick\": false }",
                "warnings" => "{ \"wait_for_intellisense\": true, \"quick\": false }",
                "apply_diff" => "{ \"patch\": \"*** Begin Patch\\n*** End Patch\\n\", \"post_check\": true }",
                "debug_watch" => "{ \"expression\": \"count\", \"timeout_ms\": 1000 }",
                "set_build_configuration" => "{ \"configuration\": \"Debug\", \"platform\": \"x64\" }",
                "count_references" => "{ \"file\": \"src\\\\foo.cpp\", \"line\": 42, \"column\": 13 }",
                "nuget_restore" => "{ \"path\": \"VsIdeBridge.sln\" }",
                "nuget_add_package" => "{ \"project\": \"src\\\\VsIdeBridgeCli\\\\VsIdeBridgeCli.csproj\", \"package\": \"Newtonsoft.Json\", \"version\": \"13.0.3\" }",
                "nuget_remove_package" => "{ \"project\": \"src\\\\VsIdeBridgeCli\\\\VsIdeBridgeCli.csproj\", \"package\": \"Newtonsoft.Json\" }",
                "conda_install" => "{ \"packages\": [\"cmake\", \"ninja\"], \"name\": \"superslicer\", \"yes\": true }",
                "conda_remove" => "{ \"packages\": [\"ninja\"], \"name\": \"superslicer\", \"yes\": true }",
                _ => string.Empty,
            };

            if (!string.IsNullOrWhiteSpace(overrideExample))
            {
                return overrideExample;
            }

            var example = new JsonObject();
            var requiredNames = inputSchema["required"] as JsonArray ?? [];
            var properties = inputSchema["properties"] as JsonObject ?? [];
            foreach (var required in requiredNames.OfType<JsonNode>())
            {
                var nameToken = required.GetValue<string>();
                var propertySchema = properties[nameToken] as JsonObject;
                var type = propertySchema?["type"]?.GetValue<string>() ?? "string";
                example[nameToken] = type switch
                {
                    "integer" => 1,
                    "boolean" => true,
                    "array" => new JsonArray("value"),
                    _ => "value",
                };
            }

            return example.ToJsonString(JsonOptions);
        }

        private static string? ResolveBridgeCommandForTool(string toolName)
        {
            return toolName switch
            {
                "help" => "help",
                "state" => "state",
                "ready" => "ready",
                "errors" => "errors",
                "warnings" => "warnings",
                "list_tabs" => "list-tabs",
                "open_file" => "open-document",
                "find_files" => "find-files",
                "search_symbols" => "search-symbols",
                "count_references" => "count-references",
                "quick_info" => "quick-info",
                "apply_diff" => "apply-diff",
                "debug_threads" => "debug-threads",
                "debug_stack" => "debug-stack",
                "debug_locals" => "debug-locals",
                "debug_modules" => "debug-modules",
                "debug_watch" => "debug-watch",
                "debug_exceptions" => "debug-exceptions",
                "diagnostics_snapshot" => "diagnostics-snapshot",
                "build_configurations" => "build-configurations",
                "set_build_configuration" => "set-build-configuration",
                "find_text" => "find-text",
                "read_file" => "document-slice",
                "find_references" => "find-references",
                "peek_definition" => "quick-info",
                "file_outline" => "file-outline",
                "build" => "build",
                "open_solution" => "open-solution",
                _ => null,
            };
        }

        private static async Task<JsonNode> BridgeHealthAsync(JsonNode? id, BridgeBinding bridgeBinding)
        {
            var sw = Stopwatch.StartNew();
            var state = await SendBridgeAsync(id, bridgeBinding, "state", string.Empty).ConfigureAwait(false);
            var ready = await SendBridgeAsync(id, bridgeBinding, "ready", string.Empty).ConfigureAwait(false);
            sw.Stop();

            var discovery = bridgeBinding.CurrentDiscovery;
            var stateData = state["Data"] as JsonObject;
            var watchdog = stateData?["watchdog"] as JsonObject;
            var isDegraded = GetBoolean(watchdog, "isDegraded", false);
            var readySuccess = ResponseFormatter.IsSuccess(ready);
            var stateSuccess = ResponseFormatter.IsSuccess(state);
            var result = new JsonObject
            {
                ["success"] = stateSuccess && readySuccess && !isDegraded,
                ["status"] = stateSuccess && readySuccess && !isDegraded ? "healthy" : "degraded",
                ["binding"] = discovery is null ? null : DiscoveryToJson(discovery),
                ["selector"] = SelectorToJson(bridgeBinding.CurrentSelector),
                ["roundTripMs"] = Math.Round(sw.Elapsed.TotalMilliseconds, 1),
                ["state"] = state,
                ["lastReady"] = ready,
                ["watchdog"] = watchdog,
            };

            return WrapToolResult(result, isError: false);
        }

        private static async Task<JsonNode> OpenSolutionAsync(JsonNode? id, JsonObject? args, BridgeBinding bridgeBinding)
        {
            var solution = args?["solution"]?.GetValue<string>();
            if (string.IsNullOrWhiteSpace(solution))
            {
                throw new McpRequestException(id, -32602, "open_solution requires a non-empty solution path.");
            }

            var waitForReady = GetBoolean(args, "wait_for_ready", true);
            // Clear solution hint before sending so instance lookup succeeds even when VS has a different solution open.
            bridgeBinding.PreferSolution(null);
            var open = await SendBridgeAsync(id, bridgeBinding, "open-solution", BuildArgs(("solution", solution))).ConfigureAwait(false);

            JsonObject? ready = null;
            JsonObject? state = null;
            if (ResponseFormatter.IsSuccess(open))
            {
                bridgeBinding.PreferSolution(solution);
                if (waitForReady)
                {
                    ready = await SendBridgeAsync(id, bridgeBinding, "ready", string.Empty).ConfigureAwait(false);
                }

                state = await SendBridgeAsync(id, bridgeBinding, "state", string.Empty).ConfigureAwait(false);
            }

            var result = new JsonObject
            {
                ["open"] = open,
                ["ready"] = ready,
                ["state"] = state,
            };

            return WrapToolResult(result, isError: !ResponseFormatter.IsSuccess(open));
        }

        private static async Task<JsonNode> ListInstancesAsync(BridgeBinding bridgeBinding)
        {
            var instances = await PipeDiscovery.ListAsync(verbose: false, bridgeBinding.DiscoveryMode).ConfigureAwait(false);
            var boundInstanceId = bridgeBinding.CurrentDiscovery?.InstanceId;
            var items = new JsonArray();
            foreach (var instance in instances.OrderByDescending(item => item.LastWriteTimeUtc))
            {
                var json = DiscoveryToJson(instance);
                json["is_bound"] = string.Equals(instance.InstanceId, boundInstanceId, StringComparison.OrdinalIgnoreCase);
                items.Add(json);
            }

            var result = new JsonObject
            {
                ["success"] = true,
                ["count"] = instances.Count,
                ["items"] = items,
            };

            return new JsonObject
            {
                ["content"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["type"] = "text",
                        ["text"] = result.ToJsonString(JsonOptions),
                    },
                },
                ["isError"] = false,
                ["structuredContent"] = result,
            };
        }

        private static JsonArray ListResources() =>
        [
            Resource("bridge://current-solution", "Current solution"),
            Resource("bridge://active-document", "Active document"),
            Resource("bridge://open-tabs", "Open tabs"),
            Resource("bridge://error-list-snapshot", "Error list snapshot"),
        ];

        private static JsonArray ListResourceTemplates() => [];

        private static JsonObject Resource(string uri, string name) => new()
        {
            ["uri"] = uri,
            ["name"] = name,
            ["mimeType"] = "application/json",
        };

        private static async Task<JsonNode> ReadResourceAsync(JsonNode? id, JsonObject? p, BridgeBinding bridgeBinding)
        {
            var uri = p?["uri"]?.GetValue<string>() ?? throw new McpRequestException(id, -32602, "resources/read missing uri.");
            JsonObject data = uri switch
            {
                "bridge://current-solution" => await SendBridgeAsync(id, bridgeBinding, "state", string.Empty).ConfigureAwait(false),
                "bridge://active-document" => await SendBridgeAsync(id, bridgeBinding, "state", string.Empty).ConfigureAwait(false),
                "bridge://open-tabs" => await SendBridgeAsync(id, bridgeBinding, "list-tabs", string.Empty).ConfigureAwait(false),
                "bridge://error-list-snapshot" => await SendBridgeAsync(id, bridgeBinding, "errors", "--quick --wait-for-intellisense false").ConfigureAwait(false),
                _ => throw new McpRequestException(id, -32602, $"Unknown resource uri: {uri}"),
            };

            return new JsonObject
            {
                ["contents"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["uri"] = uri,
                        ["mimeType"] = "application/json",
                        ["text"] = data.ToJsonString(JsonOptions),
                    },
                },
            };
        }

        private static JsonArray ListPrompts() =>
        [
            Prompt("help", "Show bridge and MCP usage guidance."),
            Prompt("fix_current_errors", "Gather errors and propose patch flow."),
            Prompt("open_solution_and_wait_ready", "Run ensure then ready flow."),
            Prompt("git_review_before_commit", "Review status, diff, and log before committing."),
            Prompt("git_sync_with_remote", "Fetch, inspect divergence, then pull or push safely."),
            Prompt("github_issue_triage", "Search open issues, inspect details, and close resolved items."),
        ];

        private static JsonObject Prompt(string name, string description) => new()
        {
            ["name"] = name,
            ["description"] = description,
            ["arguments"] = new JsonArray(),
        };

        private static JsonObject GetPrompt(JsonNode? id, JsonObject? p)
        {
            var name = p?["name"]?.GetValue<string>() ?? throw new McpRequestException(id, -32602, "prompts/get missing name.");
            var text = name switch
            {
                "help" => "Key tools: bind_solution or bind_instance to connect, open_solution to load a .sln, state/ready/bridge_health for status. Navigation: find_files, find_text, open_file, search_symbols, count_references, quick_info, read_file, find_references, peek_definition, file_outline. Editing: apply_diff (optionally post_check). Diagnostics: errors, warnings, diagnostics_snapshot, build, build_configurations, set_build_configuration. Dependencies: nuget_restore, nuget_add_package, nuget_remove_package, conda_install, conda_remove. Debug: debug_threads, debug_stack, debug_locals, debug_modules, debug_watch, debug_exceptions. Use tool_help for per-tool schemas and examples.",
                "fix_current_errors" => "Bind to the right solution first, call errors to list problems. Use read_file or find_text to inspect code, quick_info and find_references for context, then apply_diff to fix.",
                "open_solution_and_wait_ready" => "Call open_solution with the absolute .sln path and wait_for_ready=true (default). Then call state or bridge_health.",
                "git_review_before_commit" => "Call git_status, git_diff_unstaged, git_diff_staged, git_log, then git_add and git_commit when ready.",
                "git_sync_with_remote" => "Call git_fetch, git_status, and git_log first. Then use git_pull when behind or git_push when ahead.",
                "github_issue_triage" => "Use github_issue_search with state=open, review candidates, then use github_issue_close with issue_number and optional comment.",
                _ => throw new McpRequestException(id, -32602, $"Unknown prompt: {name}"),
            };

            return new JsonObject
            {
                ["description"] = $"Bridge prompt: {name}",
                ["messages"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["role"] = "user",
                        ["content"] = new JsonObject { ["type"] = "text", ["text"] = text },
                    },
                },
            };
        }

        private static Task<JsonObject> SendBridgeAsync(JsonNode? id, BridgeBinding bridgeBinding, string command, string args)
        {
            return bridgeBinding.SendAsync(id, command, args);
        }

        private static JsonObject DiscoveryToJson(PipeDiscovery discovery)
        {
            return new JsonObject
            {
                ["instanceId"] = discovery.InstanceId,
                ["pid"] = discovery.ProcessId,
                ["pipeName"] = discovery.PipeName,
                ["solutionPath"] = discovery.SolutionPath,
                ["solutionName"] = discovery.SolutionName,
                ["source"] = discovery.Source,
                ["startedAtUtc"] = discovery.StartedAtUtc,
                ["discoveryFile"] = discovery.DiscoveryFile,
                ["lastWriteTimeUtc"] = discovery.LastWriteTimeUtc.ToString("O"),
            };
        }

        private static JsonObject SelectorToJson(BridgeInstanceSelector selector)
        {
            return new JsonObject
            {
                ["instanceId"] = selector.InstanceId,
                ["pid"] = selector.ProcessId,
                ["pipeName"] = selector.PipeName,
                ["solutionHint"] = selector.SolutionHint,
            };
        }

        private static string BuildArgs(params (string Name, string? Value)[] items)
        {
            var builder = new PipeArgsBuilder();
            foreach (var (name, value) in items)
            {
                builder.Add(name, value);
            }

            return builder.Build();
        }

        private static bool GetBoolean(JsonObject? args, string name, bool defaultValue)
        {
            return args?[name]?.GetValue<bool?>() ?? defaultValue;
        }

        private static string? GetCsv(JsonArray? values)
        {
            if (values is null || values.Count == 0)
            {
                return null;
            }

            var items = values
                .OfType<JsonNode>()
                .Select(item => item.GetValue<string>())
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Select(item => item.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (items.Length == 0)
            {
                return null;
            }

            return string.Join(",", items);
        }

        private static JsonObject EmptySchema() => new()
        {
            ["type"] = "object",
            ["properties"] = new JsonObject(),
            ["additionalProperties"] = false,
        };

        private static JsonObject ObjectSchema(params (string Name, JsonObject Schema, bool Required)[] properties)
        {
            var schema = new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject(),
                ["additionalProperties"] = false,
            };

            var propertyBag = (JsonObject)schema["properties"]!;
            var required = new JsonArray();
            foreach (var (name, propertySchema, isRequired) in properties)
            {
                propertyBag[name] = propertySchema;
                if (isRequired)
                {
                    required.Add(name);
                }
            }

            if (required.Count > 0)
            {
                schema["required"] = required;
            }

            return schema;
        }

        private static JsonObject StringSchema(string description) => new()
        {
            ["type"] = "string",
            ["description"] = description,
        };

        private static JsonObject IntegerSchema(string description) => new()
        {
            ["type"] = "integer",
            ["description"] = description,
        };

        private static JsonObject BooleanSchema(string description) => new()
        {
            ["type"] = "boolean",
            ["description"] = description,
        };

        private static JsonObject ArrayOfStringsSchema(string description) => new()
        {
            ["type"] = "array",
            ["description"] = description,
            ["items"] = new JsonObject
            {
                ["type"] = "string",
            },
        };

        private static async Task<JsonNode> CallGitToolAsync(JsonNode? id, string toolName, JsonObject? args, BridgeBinding bridgeBinding)
        {
            var workingDirectory = await ResolveSolutionWorkingDirectoryAsync(id, bridgeBinding).ConfigureAwait(false);

            var gitArgs = toolName switch
            {
                "git_status" => "status --porcelain=v1 --branch",
                "git_current_branch" => "branch --show-current",
                "git_remote_list" => "remote --verbose",
                "git_tag_list" => "tag --list --sort=version:refname",
                "git_stash_list" => "stash list",
                "git_diff_unstaged" => $"diff --no-color --unified={GetIntOrDefault(args, "context", 3)}",
                "git_diff_staged" => $"diff --cached --no-color --unified={GetIntOrDefault(args, "context", 3)}",
                "git_log" => $"log --max-count={GetIntOrDefault(args, "max_count", 20)} --date=iso-strict --pretty=format:%H%x09%ad%x09%an%x09%s",
                "git_show" => $"show --no-color {QuoteForGit(GetRequiredString(args, id, "revision"))}",
                "git_branch_list" => "branch --all --verbose --no-abbrev",
                "git_checkout" => $"checkout {QuoteForGit(GetRequiredString(args, id, "target"))}",
                "git_create_branch" => BuildGitCreateBranchArgs(args, id),
                "git_add" => $"add -- {JoinGitPaths(GetRequiredPaths(args, id, "paths"))}",
                "git_restore" => $"restore --source=HEAD --worktree -- {JoinGitPaths(GetRequiredPaths(args, id, "paths"))}",
                "git_commit" => $"commit -m {QuoteForGit(GetRequiredString(args, id, "message"))}",
                "git_commit_amend" => BuildGitCommitAmendArgs(args),
                "git_reset" => BuildGitResetArgs(args),
                "git_fetch" => BuildGitFetchArgs(args),
                "git_stash_push" => BuildGitStashPushArgs(args),
                "git_stash_pop" => "stash pop",
                "git_pull" => BuildGitPullPushArgs("pull", args),
                "git_push" => BuildGitPullPushArgs("push", args),
                _ => throw new McpRequestException(id, -32602, $"Unknown MCP tool: {toolName}"),
            };

            var gitResult = await RunGitAsync(workingDirectory, gitArgs).ConfigureAwait(false);
            return new JsonObject
            {
                ["content"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["type"] = "text",
                        ["text"] = gitResult["stdout"]?.GetValue<string>() ?? string.Empty,
                    },
                },
                ["isError"] = !(gitResult["success"]?.GetValue<bool>() ?? false),
                ["structuredContent"] = gitResult,
            };
        }

        private static async Task<JsonNode> CallNuGetToolAsync(JsonNode? id, string toolName, JsonObject? args, BridgeBinding bridgeBinding)
        {
            var workingDirectory = await ResolveSolutionWorkingDirectoryAsync(id, bridgeBinding).ConfigureAwait(false);
            var nugetArgs = toolName switch
            {
                "nuget_restore" => BuildNuGetRestoreArgs(args),
                "nuget_add_package" => BuildNuGetAddPackageArgs(args, id),
                "nuget_remove_package" => BuildNuGetRemovePackageArgs(args, id),
                _ => throw new McpRequestException(id, -32602, $"Unknown MCP tool: {toolName}"),
            };

            var nugetResult = await RunProcessAsync(DotNetExecutableName, nugetArgs, workingDirectory).ConfigureAwait(false);
            return new JsonObject
            {
                ["content"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["type"] = "text",
                        ["text"] = nugetResult["stdout"]?.GetValue<string>() ?? string.Empty,
                    },
                },
                ["isError"] = !(nugetResult["success"]?.GetValue<bool>() ?? false),
                ["structuredContent"] = nugetResult,
            };
        }

        private static async Task<JsonNode> CallCondaToolAsync(JsonNode? id, string toolName, JsonObject? args, BridgeBinding bridgeBinding)
        {
            var condaExecutable = ResolveCondaExecutable(id);
            var workingDirectory = await ResolveSolutionWorkingDirectoryAsync(id, bridgeBinding).ConfigureAwait(false);
            var condaArgs = toolName switch
            {
                "conda_install" => BuildCondaInstallArgs(args, id),
                "conda_remove" => BuildCondaRemoveArgs(args, id),
                _ => throw new McpRequestException(id, -32602, $"Unknown MCP tool: {toolName}"),
            };
            var condaResult = await RunProcessAsync(condaExecutable, condaArgs, workingDirectory).ConfigureAwait(false);

            return new JsonObject
            {
                ["content"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["type"] = "text",
                        ["text"] = condaResult["stdout"]?.GetValue<string>() ?? string.Empty,
                    },
                },
                ["isError"] = !(condaResult["success"]?.GetValue<bool>() ?? false),
                ["structuredContent"] = condaResult,
            };
        }

        private static string BuildNuGetRestoreArgs(JsonObject? args)
        {
            var path = args?["path"]?.GetValue<string>();
            return string.IsNullOrWhiteSpace(path)
                ? "restore"
                : $"restore {QuoteForProcess(path)}";
        }

        private static string BuildNuGetAddPackageArgs(JsonObject? args, JsonNode? id)
        {
            var project = GetRequiredString(args, id, "project");
            var package = GetRequiredString(args, id, "package");
            var version = args?["version"]?.GetValue<string>();
            var source = args?["source"]?.GetValue<string>();
            var prerelease = GetBoolean(args, "prerelease", false);
            var noRestore = GetBoolean(args, "no_restore", false);

            var segments = new List<string>
            {
                "add",
                QuoteForProcess(project),
                "package",
                QuoteForProcess(package),
            };

            if (!string.IsNullOrWhiteSpace(version))
            {
                segments.Add("--version");
                segments.Add(QuoteForProcess(version));
            }

            if (!string.IsNullOrWhiteSpace(source))
            {
                segments.Add("--source");
                segments.Add(QuoteForProcess(source));
            }

            if (prerelease)
            {
                segments.Add("--prerelease");
            }

            if (noRestore)
            {
                segments.Add("--no-restore");
            }

            return string.Join(" ", segments);
        }

        private static string BuildNuGetRemovePackageArgs(JsonObject? args, JsonNode? id)
        {
            var project = GetRequiredString(args, id, "project");
            var package = GetRequiredString(args, id, "package");

            return $"remove {QuoteForProcess(project)} package {QuoteForProcess(package)}";
        }

        private static string BuildCondaInstallArgs(JsonObject? args, JsonNode? id)
        {
            var packages = GetRequiredStringArray(args, id, "packages");
            var channels = GetOptionalStringArray(args, "channels");
            var environmentName = args?["name"]?.GetValue<string>();
            var environmentPrefix = args?["prefix"]?.GetValue<string>();
            var dryRun = GetBoolean(args, "dry_run", false);
            var autoYes = GetBoolean(args, "yes", true);

            var segments = new List<string> { "install" };
            AppendCondaEnvironmentSelector(segments, environmentName, environmentPrefix, id, "conda_install");

            foreach (var channel in channels)
            {
                segments.Add("--channel");
                segments.Add(QuoteForProcess(channel));
            }

            foreach (var package in packages)
            {
                segments.Add(QuoteForProcess(package));
            }

            if (autoYes)
            {
                segments.Add("--yes");
            }

            if (dryRun)
            {
                segments.Add("--dry-run");
            }

            return string.Join(" ", segments);
        }

        private static string BuildCondaRemoveArgs(JsonObject? args, JsonNode? id)
        {
            var packages = GetRequiredStringArray(args, id, "packages");
            var environmentName = args?["name"]?.GetValue<string>();
            var environmentPrefix = args?["prefix"]?.GetValue<string>();
            var dryRun = GetBoolean(args, "dry_run", false);
            var autoYes = GetBoolean(args, "yes", true);

            var segments = new List<string> { "remove" };
            AppendCondaEnvironmentSelector(segments, environmentName, environmentPrefix, id, "conda_remove");

            foreach (var package in packages)
            {
                segments.Add(QuoteForProcess(package));
            }

            if (autoYes)
            {
                segments.Add("--yes");
            }

            if (dryRun)
            {
                segments.Add("--dry-run");
            }

            return string.Join(" ", segments);
        }

        private static void AppendCondaEnvironmentSelector(
            List<string> segments,
            string? environmentName,
            string? environmentPrefix,
            JsonNode? id,
            string toolName)
        {
            if (!string.IsNullOrWhiteSpace(environmentName) && !string.IsNullOrWhiteSpace(environmentPrefix))
            {
                throw new McpRequestException(id, -32602, $"{toolName} accepts either name or prefix, not both.");
            }

            if (!string.IsNullOrWhiteSpace(environmentName))
            {
                segments.Add("--name");
                segments.Add(QuoteForProcess(environmentName));
            }

            if (!string.IsNullOrWhiteSpace(environmentPrefix))
            {
                segments.Add("--prefix");
                segments.Add(QuoteForProcess(environmentPrefix));
            }
        }

        private static string ResolveCondaExecutable(JsonNode? id)
        {
            var condaFromEnvironment = Environment.GetEnvironmentVariable("CONDA_EXE");
            if (!string.IsNullOrWhiteSpace(condaFromEnvironment))
            {
                if (File.Exists(condaFromEnvironment))
                {
                    return condaFromEnvironment;
                }

                throw new McpRequestException(id, -32007, $"CONDA_EXE points to '{condaFromEnvironment}', but the file does not exist.");
            }

            var condaFromPath = ResolveCondaFromPath();
            if (!string.IsNullOrWhiteSpace(condaFromPath))
            {
                return condaFromPath;
            }

            var condaFromKnownLocations = ResolveCondaFromKnownLocations();
            if (!string.IsNullOrWhiteSpace(condaFromKnownLocations))
            {
                return condaFromKnownLocations;
            }

            throw new McpRequestException(id, -32007, "Conda executable not found. Install Miniconda/Anaconda or set CONDA_EXE.");
        }

        private static string? ResolveCondaFromPath()
        {
            var pathValue = Environment.GetEnvironmentVariable("PATH");
            if (string.IsNullOrWhiteSpace(pathValue))
            {
                return null;
            }

            var directories = pathValue.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            foreach (var directory in directories)
            {
                foreach (var extension in CondaExecutableExtensions)
                {
                    var candidate = Path.Combine(directory, $"{CondaExecutableName}{extension}");
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
            }

            return null;
        }

        private static string? ResolveCondaFromKnownLocations()
        {
            var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var roots = new List<string>();
            if (!string.IsNullOrWhiteSpace(userProfile))
            {
                roots.Add(userProfile);
            }

            if (!string.IsNullOrWhiteSpace(localAppData))
            {
                roots.Add(localAppData);
            }

            foreach (var root in roots)
            {
                foreach (var relativePath in CondaRelativeCandidatePaths)
                {
                    var candidate = Path.Combine(root, relativePath);
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
            }

            return null;
        }


        private static string BuildGitCommitAmendArgs(JsonObject? args)
        {
            var message = args?["message"]?.GetValue<string>();
            var noEdit = args?["no_edit"]?.GetValue<bool>() == true;
            if (!string.IsNullOrWhiteSpace(message))
            {
                return $"commit --amend -m {QuoteForGit(message)}";
            }

            return noEdit ? "commit --amend --no-edit" : "commit --amend";
        }

        private static string BuildGitFetchArgs(JsonObject? args)
        {
            var remote = args?["remote"]?.GetValue<string>();
            var fetchAll = args?["all"]?.GetValue<bool>() == true;
            var prune = args?["prune"]?.GetValue<bool>();

            var segments = new List<string> { "fetch" };
            if (fetchAll)
            {
                segments.Add("--all");
            }

            if (prune != false)
            {
                segments.Add("--prune");
            }

            if (!string.IsNullOrWhiteSpace(remote))
            {
                segments.Add(QuoteForGit(remote));
            }

            return string.Join(" ", segments);
        }

        private static string BuildGitStashPushArgs(JsonObject? args)
        {
            var includeUntracked = args?["include_untracked"]?.GetValue<bool>() == true;
            var message = args?["message"]?.GetValue<string>();

            var segments = new List<string> { "stash", "push" };
            if (includeUntracked)
            {
                segments.Add("--include-untracked");
            }

            if (!string.IsNullOrWhiteSpace(message))
            {
                segments.Add("-m");
                segments.Add(QuoteForGit(message));
            }

            return string.Join(" ", segments);
        }
        private static string BuildGitResetArgs(JsonObject? args)
        {
            var paths = GetOptionalPaths(args, "paths");
            return paths.Count == 0
                ? "reset"
                : $"reset -- {JoinGitPaths(paths)}";
        }

        private static string BuildGitPullPushArgs(string verb, JsonObject? args)
        {
            var remote = args?["remote"]?.GetValue<string>();
            var branch = args?["branch"]?.GetValue<string>();
            var setUpstream = args?["set_upstream"]?.GetValue<bool>() == true;

            var segments = new List<string> { verb };
            if (setUpstream && string.Equals(verb, "push", StringComparison.Ordinal))
            {
                segments.Add("--set-upstream");
            }

            if (!string.IsNullOrWhiteSpace(remote))
            {
                segments.Add(QuoteForGit(remote));
            }

            if (!string.IsNullOrWhiteSpace(branch))
            {
                segments.Add(QuoteForGit(branch));
            }

            return string.Join(" ", segments);
        }

        private static string BuildGitCreateBranchArgs(JsonObject? args, JsonNode? id)
        {
            var name = GetRequiredString(args, id, "name");
            var startPoint = args?["start_point"]?.GetValue<string>();
            return string.IsNullOrWhiteSpace(startPoint)
                ? $"checkout -b {QuoteForGit(name)}"
                : $"checkout -b {QuoteForGit(name)} {QuoteForGit(startPoint)}";
        }

        private static async Task<string> ResolveSolutionWorkingDirectoryAsync(JsonNode? id, BridgeBinding bridgeBinding)
        {
            var state = await SendBridgeAsync(id, bridgeBinding, "state", string.Empty).ConfigureAwait(false);
            var data = state["Data"] as JsonObject;
            var solutionPath = data?["solutionPath"]?.GetValue<string>() ?? string.Empty;
            var directory = Path.GetDirectoryName(solutionPath);
            if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            {
                throw new McpRequestException(id, -32004, "Could not determine solution directory for local package/git operations. Ensure a solution is open.");
            }

            return directory;
        }

        private static async Task<JsonObject> RunGitAsync(string workingDirectory, string arguments)
        {
            return await RunProcessAsync(GitExecutableName, arguments, workingDirectory).ConfigureAwait(false);
        }

        private static async Task<JsonObject> RunProcessAsync(string command, string arguments, string workingDirectory)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = command,
                Arguments = arguments,
                WorkingDirectory = workingDirectory,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };

            using var process = new Process { StartInfo = startInfo };
            process.Start();

            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync().ConfigureAwait(false);

            var stdout = await stdoutTask.ConfigureAwait(false);
            var stderr = await stderrTask.ConfigureAwait(false);
            return new JsonObject
            {
                ["success"] = process.ExitCode == 0,
                ["exitCode"] = process.ExitCode,
                ["command"] = command,
                ["workingDirectory"] = workingDirectory,
                ["args"] = arguments,
                ["stdout"] = stdout,
                ["stderr"] = stderr,
            };
        }

        private static string GetRequiredString(JsonObject? args, JsonNode? id, string name)
        {
            var value = args?[name]?.GetValue<string>();
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new McpRequestException(id, -32602, $"tools/call arguments missing required field '{name}'.");
            }

            return value;
        }

        private static int GetIntOrDefault(JsonObject? args, string name, int defaultValue)
        {
            return args?[name]?.GetValue<int?>() ?? defaultValue;
        }

        private static List<string> GetRequiredPaths(JsonObject? args, JsonNode? id, string name)
        {
            var paths = GetOptionalPaths(args, name);
            if (paths.Count == 0)
            {
                throw new McpRequestException(id, -32602, $"tools/call arguments missing required array '{name}'.");
            }

            return paths;
        }

        private static List<string> GetOptionalPaths(JsonObject? args, string name)
        {
            return GetOptionalStringArray(args, name);
        }

        private static List<string> GetRequiredStringArray(JsonObject? args, JsonNode? id, string name)
        {
            var values = GetOptionalStringArray(args, name);
            if (values.Count == 0)
            {
                throw new McpRequestException(id, -32602, $"tools/call arguments missing required array '{name}'.");
            }

            return values;
        }

        private static List<string> GetOptionalStringArray(JsonObject? args, string name)
        {
            return args?[name] is JsonArray array
                ? [.. array.OfType<JsonNode>().Select(node => node.GetValue<string>()).Where(value => !string.IsNullOrWhiteSpace(value))]
                : [];
        }

        private static string JoinGitPaths(IEnumerable<string> paths)
        {
            return string.Join(" ", paths.Select(QuoteForGit));
        }

        private static string QuoteForGit(string input)
        {
            return QuoteForProcess(input);
        }

        private static string QuoteForProcess(string input)
        {
            var escaped = input.Replace("\\", "\\\\", StringComparison.Ordinal).Replace("\"", "\\\"", StringComparison.Ordinal);
            return $"\"{escaped}\"";
        }


        private static async Task<JsonNode> CallGitHubToolAsync(JsonNode? id, string toolName, JsonObject? args, BridgeBinding bridgeBinding)
        {
            var workingDirectory = await ResolveSolutionWorkingDirectoryAsync(id, bridgeBinding).ConfigureAwait(false);
            var repo = args?["repo"]?.GetValue<string>();
            if (string.IsNullOrWhiteSpace(repo))
            {
                repo = await ResolveGitHubRepoFromOriginAsync(workingDirectory).ConfigureAwait(false);
            }

            if (string.IsNullOrWhiteSpace(repo))
            {
                throw new McpRequestException(id, -32005, "Could not determine GitHub repository. Pass -- repo as owner/repo or set origin remote.");
            }

            var token = Environment.GetEnvironmentVariable("GITHUB_TOKEN") ?? Environment.GetEnvironmentVariable("GH_TOKEN");
            if (string.IsNullOrWhiteSpace(token))
            {
                throw new McpRequestException(id, -32006, "Missing GITHUB_TOKEN (or GH_TOKEN) for GitHub issue operations.");
            }

            var result = toolName switch
            {
                "github_issue_search" => await GitHubIssueSearchAsync(repo, args, token).ConfigureAwait(false),
                "github_issue_close" => await GitHubIssueCloseAsync(repo, args, token, id).ConfigureAwait(false),
                _ => throw new McpRequestException(id, -32602, $"Unknown MCP tool: {toolName}"),
            };

            return new JsonObject
            {
                ["content"] = new JsonArray { new JsonObject { ["type"] = "text", ["text"] = result.ToJsonString(JsonOptions) } },
                ["isError"] = !(result["success"]?.GetValue<bool>() ?? false),
                ["structuredContent"] = result,
            };
        }

        private static async Task<JsonObject> GitHubIssueSearchAsync(string repo, JsonObject? args, string token)
        {
            var query = args?["query"]?.GetValue<string>()?.Trim();
            var state = (args?["state"]?.GetValue<string>() ?? "open").Trim().ToLowerInvariant();
            if (state is not ("open" or "closed" or "all"))
            {
                state = "open";
            }

            var limit = Math.Clamp(GetIntOrDefault(args, "limit", 20), 1, 100);
            var q = $"repo:{repo} is:issue";
            if (state != "all")
            {
                q += $" is:{state}";
            }

            if (!string.IsNullOrWhiteSpace(query))
            {
                q += $" {query}";
            }

            var uri = $"https://api.github.com/search/issues?q={Uri.EscapeDataString(q)}&per_page={limit}";
            return await SendGitHubRequestAsync(HttpMethod.Get, uri, token).ConfigureAwait(false);
        }

        private static async Task<JsonObject> GitHubIssueCloseAsync(string repo, JsonObject? args, string token, JsonNode? id)
        {
            var issueNumber = args?["issue_number"]?.GetValue<int?>();
            if (issueNumber is null || issueNumber <= 0)
            {
                throw new McpRequestException(id, -32602, "github_issue_close requires a positive issue_number.");
            }

            var comment = args?["comment"]?.GetValue<string>();
            if (!string.IsNullOrWhiteSpace(comment))
            {
                var commentUri = $"https://api.github.com/repos/{repo}/issues/{issueNumber}/comments";
                _ = await SendGitHubRequestAsync(HttpMethod.Post, commentUri, token, new JsonObject { ["body"] = comment }).ConfigureAwait(false);
            }

            var closeUri = $"https://api.github.com/repos/{repo}/issues/{issueNumber}";
            return await SendGitHubRequestAsync(HttpMethod.Patch, closeUri, token, new JsonObject { ["state"] = "closed" }).ConfigureAwait(false);
        }

        private static async Task<string?> ResolveGitHubRepoFromOriginAsync(string workingDirectory)
        {
            var result = await RunGitAsync(workingDirectory, "remote get-url origin").ConfigureAwait(false);
            if (!(result["success"]?.GetValue<bool>() ?? false))
            {
                return null;
            }

            var url = (result["stdout"]?.GetValue<string>() ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(url))
            {
                return null;
            }

            var normalized = url.Replace(".git", string.Empty, StringComparison.OrdinalIgnoreCase);
            const string sshPrefix = "git@github.com:";
            const string httpsPrefix = "https://github.com/";
            if (normalized.StartsWith(sshPrefix, StringComparison.OrdinalIgnoreCase))
            {
                return normalized[sshPrefix.Length..];
            }

            if (normalized.StartsWith(httpsPrefix, StringComparison.OrdinalIgnoreCase))
            {
                return normalized[httpsPrefix.Length..];
            }

            return null;
        }

        private static async Task<JsonObject> SendGitHubRequestAsync(HttpMethod method, string uri, string token, JsonObject? body = null)
        {
            using var client = new HttpClient();
            using var request = new HttpRequestMessage(method, uri);
            request.Headers.TryAddWithoutValidation("User-Agent", "vs-ide-bridge-mcp");
            request.Headers.TryAddWithoutValidation("Authorization", $"Bearer {token}");
            request.Headers.TryAddWithoutValidation("X-GitHub-Api-Version", "2022-11-28");
            request.Headers.TryAddWithoutValidation("Accept", "application/vnd.github+json");
            if (body is not null)
            {
                request.Content = new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json");
            }

            using var response = await client.SendAsync(request).ConfigureAwait(false);
            var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            JsonNode? json = null;
            if (!string.IsNullOrWhiteSpace(payload))
            {
                try { json = JsonNode.Parse(payload); } catch { }
            }

            return new JsonObject
            {
                ["success"] = response.IsSuccessStatusCode,
                ["statusCode"] = (int)response.StatusCode,
                ["uri"] = uri,
                ["method"] = method.Method,
                ["data"] = json,
                ["raw"] = json is null ? payload : null,
            };
        }

        private static JsonObject CreateErrorResponse(JsonNode? id, int code, string message) => new()
        {
            ["jsonrpc"] = "2.0",
            ["id"] = id,
            ["error"] = new JsonObject
            {
                ["code"] = code,
                ["message"] = message,
            },
        };

        private static async Task<McpIncomingMessage?> ReadMessageAsync(Stream input)
        {
            var firstByte = await ReadNextNonWhitespaceByteAsync(input).ConfigureAwait(false);
            if (firstByte is null)
            {
                return null;
            }

            if (LooksLikeRawJson(firstByte.Value))
            {
                McpTrace($"ReadMessageAsync: detected raw JSON transport starting with 0x{firstByte.Value:X2}");
                var rawJson = await ReadRawJsonMessageAsync(input, firstByte.Value).ConfigureAwait(false);
                return new McpIncomingMessage
                {
                    Request = ParseJsonObject(rawJson),
                    WireFormat = McpWireFormat.RawJson,
                };
            }

            var header = await ReadHeaderAsync(input, firstByte.Value).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(header))
            {
                return null;
            }

            var lengthLine = header.Split('\n').FirstOrDefault(line => line.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase))
                ?? throw new McpRequestException(null, -32600, "MCP request missing Content-Length header.");
            if (!int.TryParse(lengthLine.Split(':', 2)[1].Trim(), out var length) || length < 0)
            {
                throw new McpRequestException(null, -32600, "MCP request has invalid Content-Length.");
            }

            var payloadBytes = new byte[length];
            var offset = 0;
            while (offset < length)
            {
                var read = await input.ReadAsync(payloadBytes.AsMemory(offset, length - offset)).ConfigureAwait(false);
                if (read == 0)
                {
                    throw new McpRequestException(null, -32600, "Unexpected EOF while reading MCP payload.");
                }

                offset += read;
            }

            var json = Encoding.UTF8.GetString(payloadBytes);
            return new McpIncomingMessage
            {
                Request = ParseJsonObject(json),
                WireFormat = McpWireFormat.HeaderFramed,
            };
        }

        private static JsonObject ParseJsonObject(string json)
        {
            return JsonNode.Parse(json) as JsonObject
                ?? throw new McpRequestException(null, -32600, "MCP request must be a JSON object.");
        }

        private static async Task<byte?> ReadNextNonWhitespaceByteAsync(Stream input)
        {
            while (true)
            {
                var buffer = new byte[1];
                var read = await input.ReadAsync(buffer).ConfigureAwait(false);
                if (read == 0)
                {
                    return null;
                }

                if (!char.IsWhiteSpace((char)buffer[0]))
                {
                    return buffer[0];
                }
            }
        }

        private static bool LooksLikeRawJson(byte firstByte)
        {
            return firstByte == (byte)'{' || firstByte == (byte)'[';
        }

        private static async Task<string> ReadRawJsonMessageAsync(Stream input, byte firstByte)
        {
            List<byte> bytes = [firstByte];
            var depth = 1;
            var inString = false;
            var isEscaped = false;

            while (depth > 0)
            {
                var buffer = new byte[1];
                var read = await input.ReadAsync(buffer).ConfigureAwait(false);
                if (read == 0)
                {
                    throw new McpRequestException(null, -32600, "Unexpected EOF while reading raw JSON MCP payload.");
                }

                var current = buffer[0];
                bytes.Add(current);

                if (isEscaped)
                {
                    isEscaped = false;
                    continue;
                }

                if (current == (byte)'\\')
                {
                    if (inString)
                    {
                        isEscaped = true;
                    }

                    continue;
                }

                if (current == (byte)'"')
                {
                    inString = !inString;
                    continue;
                }

                if (inString)
                {
                    continue;
                }

                if (current == (byte)'{' || current == (byte)'[')
                {
                    depth++;
                }
                else if (current == (byte)'}' || current == (byte)']')
                {
                    depth--;
                }
            }

            return Encoding.UTF8.GetString([.. bytes]);
        }

        private static async Task<string> ReadHeaderAsync(Stream input, byte firstByte)
        {
            List<byte> bytes = [firstByte];
            var lastFour = new Queue<byte>(4);
            List<byte> firstBytes = []; // log first 64 bytes for diagnostics
            firstBytes.Add(firstByte);
            lastFour.Enqueue(firstByte);
            while (true)
            {
                if (lastFour.Count == 4 && lastFour.SequenceEqual(HeaderTerminator))
                {
                    McpTrace($"ReadHeaderAsync: got CRLF header after {bytes.Count} bytes");
                    return Encoding.ASCII.GetString([.. bytes]);
                }

                var arr = lastFour.ToArray();
                if (arr.Length >= 2 && arr[^1] == (byte)'\n' && arr[^2] == (byte)'\n'
                    && !(arr.Length >= 4 && arr[^4] == (byte)'\r'))
                {
                    McpTrace($"ReadHeaderAsync: got LF-only header after {bytes.Count} bytes");
                    return Encoding.ASCII.GetString([.. bytes]);
                }

                var b = new byte[1];
                var read = await input.ReadAsync(b).ConfigureAwait(false);
                if (read == 0)
                {
                    McpTrace($"ReadHeaderAsync: EOF after {bytes.Count} bytes. First bytes: {BitConverter.ToString([.. firstBytes])}");
                    return string.Empty;
                }

                bytes.Add(b[0]);
                if (firstBytes.Count < 64) firstBytes.Add(b[0]);
                lastFour.Enqueue(b[0]);
                if (lastFour.Count > 4)
                {
                    lastFour.Dequeue();
                }
            }
        }

        private static async Task WriteMessageAsync(Stream output, JsonObject response, McpWireFormat wireFormat)
        {
            var bytes = Encoding.UTF8.GetBytes(response.ToJsonString());
            if (wireFormat == McpWireFormat.RawJson)
            {
                await output.WriteAsync(bytes).ConfigureAwait(false);
                await output.WriteAsync(RawJsonTerminator).ConfigureAwait(false);
                await output.FlushAsync().ConfigureAwait(false);
                return;
            }

            var header = Encoding.ASCII.GetBytes($"Content-Length: {bytes.Length}\r\n\r\n");
            await output.WriteAsync(header).ConfigureAwait(false);
            await output.WriteAsync(bytes).ConfigureAwait(false);
            await output.FlushAsync().ConfigureAwait(false);
        }

        private sealed class McpRequestException : Exception
        {
            public McpRequestException()
            {
            }

            public McpRequestException(string message)
                : base(message)
            {
            }

            public McpRequestException(string message, Exception innerException)
                : base(message, innerException)
            {
            }

            public McpRequestException(JsonNode? id, int code, string message)
                : base(message)
            {
                Id = id;
                Code = code;
            }

            public JsonNode? Id { get; }
            public int Code { get; }
        }
    }
}
