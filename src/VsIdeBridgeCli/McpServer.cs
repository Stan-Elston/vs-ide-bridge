using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

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
        public static async Task RunAsync(CliOptions options)
        {
            var input = Console.OpenStandardInput();
            var output = Console.OpenStandardOutput();
            while (true)
            {
                JsonObject? response;

                try
                {
                    var request = await ReadMessageAsync(input).ConfigureAwait(false);
                    if (request is null)
                    {
                        return;
                    }

                    response = await HandleRequestAsync(request, options).ConfigureAwait(false);
                }
                catch (McpRequestException ex)
                {
                    response = CreateErrorResponse(ex.Id, ex.Code, ex.Message);
                }
                catch (JsonException ex)
                {
                    response = CreateErrorResponse(null, -32700, $"Parse error: {ex.Message}");
                }
                catch (Exception ex)
                {
                    response = CreateErrorResponse(null, -32603, $"Internal error: {ex.Message}");
                }

                if (response is not null)
                {
                    await WriteMessageAsync(output, response).ConfigureAwait(false);
                }
            }
        }

        private static async Task<JsonObject?> HandleRequestAsync(JsonObject request, CliOptions options)
        {
            var id = request["id"]?.DeepClone();
            var method = request["method"]?.GetValue<string>() ?? string.Empty;
            var @params = request["params"] as JsonObject;

            JsonNode result = method switch
            {
                "initialize" => InitializeResult(),
                "tools/list" => new JsonObject { ["tools"] = ListTools() },
                "tools/call" => await CallToolAsync(id, @params, options).ConfigureAwait(false),
                "resources/list" => new JsonObject { ["resources"] = ListResources() },
                "resources/read" => await ReadResourceAsync(id, @params, options).ConfigureAwait(false),
                "prompts/list" => new JsonObject { ["prompts"] = ListPrompts() },
                "prompts/get" => GetPrompt(id, @params),
                "notifications/initialized" => null!,
                _ => throw new McpRequestException(id, -32601, $"Unsupported MCP method: {method}"),
            };

            if (method == "notifications/initialized")
            {
                return null;
            }

            return new JsonObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = result };
        }

        private static JsonObject InitializeResult() => new()
        {
            ["protocolVersion"] = "2024-11-05",
            ["capabilities"] = new JsonObject
            {
                ["tools"] = new JsonObject(),
            },
            ["serverInfo"] = new JsonObject
            {
                ["name"] = "vs-ide-bridge-mcp",
                ["version"] = "0.1.0",
            },
        };

        private static JsonArray ListTools() => new()
        {
            Tool("state", "Capture current Visual Studio bridge state.", EmptySchema()),
            Tool("errors", "Get current errors.", EmptySchema()),
            Tool("warnings", "Get current warnings.", EmptySchema()),
            Tool("list_tabs", "List open editor tabs.", EmptySchema()),
            Tool(
                "open_file",
                "Open a file path and optional line/column.",
                ObjectSchema(
                    ("file", StringSchema("Absolute or solution-relative file path."), true),
                    ("line", IntegerSchema("Optional 1-based line number."), false),
                    ("column", IntegerSchema("Optional 1-based column number."), false))),
            Tool(
                "search_symbols",
                "Search solution symbols by query.",
                ObjectSchema(
                    ("query", StringSchema("Symbol search text."), true),
                    ("kind", StringSchema("Optional symbol kind filter."), false))),
            Tool(
                "quick_info",
                "Get quick info at file/line/column.",
                ObjectSchema(
                    ("file", StringSchema("Absolute or solution-relative file path."), true),
                    ("line", IntegerSchema("1-based line number."), true),
                    ("column", IntegerSchema("1-based column number."), true))),
            Tool(
                "apply_diff",
                "Apply unified diff through Visual Studio editor buffer.",
                ObjectSchema(
                    ("patch", StringSchema("Unified diff text."), true))),
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
        };

        private static JsonObject Tool(string name, string description, JsonObject inputSchema) => new()
        {
            ["name"] = name,
            ["description"] = description,
            ["inputSchema"] = inputSchema,
        };

        private static async Task<JsonNode> CallToolAsync(JsonNode? id, JsonObject? p, CliOptions options)
        {
            var toolName = p?["name"]?.GetValue<string>() ?? throw new McpRequestException(id, -32602, "tools/call missing name.");
            var args = p?["arguments"] as JsonObject;

            if (toolName.StartsWith("git_", StringComparison.Ordinal))
            {
                return await CallGitToolAsync(id, toolName, args, options).ConfigureAwait(false);
            }

            if (toolName.StartsWith("github_", StringComparison.Ordinal))
            {
                return await CallGitHubToolAsync(id, toolName, args, options).ConfigureAwait(false);
            }

            var (command, commandArgs) = toolName switch
            {
                "state" => ("state", string.Empty),
                "errors" => ("errors", "--quick --wait-for-intellisense false"),
                "warnings" => ("warnings", "--quick --wait-for-intellisense false"),
                "list_tabs" => ("list-tabs", string.Empty),
                "open_file" => ("open-document", BuildArgs(("file", args?["file"]?.GetValue<string>()), ("line", args?["line"]?.ToString()), ("column", args?["column"]?.ToString()))),
                "search_symbols" => ("search-symbols", BuildArgs(("query", args?["query"]?.GetValue<string>()), ("kind", args?["kind"]?.GetValue<string>()))),
                "quick_info" => ("quick-info", BuildArgs(("file", args?["file"]?.GetValue<string>()), ("line", args?["line"]?.ToString()), ("column", args?["column"]?.ToString()))),
                "apply_diff" => ("apply-diff", BuildArgs(("patch-text-base64", Convert.ToBase64String(Encoding.UTF8.GetBytes(args?["patch"]?.GetValue<string>() ?? string.Empty))), ("open-changed-files", "true"))),
                _ => throw new McpRequestException(id, -32602, $"Unknown MCP tool: {toolName}"),
            };

            var response = await SendBridgeAsync(id, options, command, commandArgs).ConfigureAwait(false);
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

        private static JsonArray ListResources() => new()
        {
            Resource("bridge://current-solution", "Current solution"),
            Resource("bridge://active-document", "Active document"),
            Resource("bridge://open-tabs", "Open tabs"),
            Resource("bridge://error-list-snapshot", "Error list snapshot"),
        };

        private static JsonObject Resource(string uri, string name) => new()
        {
            ["uri"] = uri,
            ["name"] = name,
            ["mimeType"] = "application/json",
        };

        private static async Task<JsonNode> ReadResourceAsync(JsonNode? id, JsonObject? p, CliOptions options)
        {
            var uri = p?["uri"]?.GetValue<string>() ?? throw new McpRequestException(id, -32602, "resources/read missing uri.");
            JsonObject data = uri switch
            {
                "bridge://current-solution" => await SendBridgeAsync(id, options, "state", string.Empty).ConfigureAwait(false),
                "bridge://active-document" => await SendBridgeAsync(id, options, "state", string.Empty).ConfigureAwait(false),
                "bridge://open-tabs" => await SendBridgeAsync(id, options, "list-tabs", string.Empty).ConfigureAwait(false),
                "bridge://error-list-snapshot" => await SendBridgeAsync(id, options, "errors", "--quick --wait-for-intellisense false").ConfigureAwait(false),
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

        private static JsonArray ListPrompts() => new()
        {
            Prompt("help", "Show bridge and MCP usage guidance."),
            Prompt("fix_current_errors", "Gather errors and propose patch flow."),
            Prompt("open_solution_and_wait_ready", "Run ensure then ready flow."),
            Prompt("git_review_before_commit", "Review status, diff, and log before committing."),
            Prompt("git_sync_with_remote", "Fetch, inspect divergence, then pull or push safely."),
            Prompt("github_issue_triage", "Search open issues, inspect details, and close resolved items."),
        };

        private static JsonObject Prompt(string name, string description) => new()
        {
            ["name"] = name,
            ["description"] = description,
            ["arguments"] = new JsonArray(),
        };

        private static JsonNode GetPrompt(JsonNode? id, JsonObject? p)
        {
            var name = p?["name"]?.GetValue<string>() ?? throw new McpRequestException(id, -32602, "prompts/get missing name.");
            var text = name switch
            {
                "help" => "Use tools state, errors, warnings, list_tabs, open_file, search_symbols, quick_info, and apply_diff.",
                "fix_current_errors" => "Call errors, inspect rows, then use open_file, quick_info, search_symbols, and apply_diff.",
                "open_solution_and_wait_ready" => "Outside MCP, run: vs-ide-bridge ensure --solution <path>; then call state until ready.",
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

        private static async Task<JsonObject> SendBridgeAsync(JsonNode? id, CliOptions options, string command, string args)
        {
            try
            {
                var selector = BridgeInstanceSelector.FromOptions(options);
                var discovery = await PipeDiscovery.SelectAsync(selector, options.GetFlag("verbose")).ConfigureAwait(false);
                await using var client = new PipeClient(discovery.PipeName, options.GetInt32("timeout-ms", 10_000));
                var request = new JsonObject
                {
                    ["id"] = Guid.NewGuid().ToString("N")[..8],
                    ["command"] = command,
                    ["args"] = args,
                };

                return await client.SendAsync(request).ConfigureAwait(false);
            }
            catch (CliException ex)
            {
                throw new McpRequestException(id, -32001, ex.Message);
            }
            catch (TimeoutException ex)
            {
                throw new McpRequestException(id, -32002, $"Timed out waiting for Visual Studio bridge response: {ex.Message}");
            }
            catch (IOException ex)
            {
                throw new McpRequestException(id, -32003, $"Failed communicating with Visual Studio bridge pipe: {ex.Message}");
            }
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

        private static async Task<JsonNode> CallGitToolAsync(JsonNode? id, string toolName, JsonObject? args, CliOptions options)
        {
            var workingDirectory = await ResolveGitWorkingDirectoryAsync(id, options).ConfigureAwait(false);

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

        private static async Task<string> ResolveGitWorkingDirectoryAsync(JsonNode? id, CliOptions options)
        {
            var state = await SendBridgeAsync(id, options, "state", string.Empty).ConfigureAwait(false);
            var data = state["Data"] as JsonObject;
            var solutionPath = data?["solutionPath"]?.GetValue<string>() ?? string.Empty;
            var directory = Path.GetDirectoryName(solutionPath);
            if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            {
                throw new McpRequestException(id, -32004, "Could not determine solution directory for git operations. Ensure a solution is open.");
            }

            return directory;
        }

        private static async Task<JsonObject> RunGitAsync(string workingDirectory, string arguments)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "git",
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
            var value = args?[name]?.GetValue<int?>();
            return value.GetValueOrDefault(defaultValue);
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
            var array = args?[name] as JsonArray;
            return array is null
                ? new List<string>()
                : array.OfType<JsonNode>().Select(node => node.GetValue<string>()).Where(path => !string.IsNullOrWhiteSpace(path)).ToList();
        }

        private static string JoinGitPaths(IEnumerable<string> paths)
        {
            return string.Join(" ", paths.Select(QuoteForGit));
        }

        private static string QuoteForGit(string input)
        {
            var escaped = input.Replace("\\", "\\\\", StringComparison.Ordinal).Replace("\"", "\\\"", StringComparison.Ordinal);
            return $"\"{escaped}\"";
        }


        private static async Task<JsonNode> CallGitHubToolAsync(JsonNode? id, string toolName, JsonObject? args, CliOptions options)
        {
            var workingDirectory = await ResolveGitWorkingDirectoryAsync(id, options).ConfigureAwait(false);
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

        private static async Task<JsonObject?> ReadMessageAsync(Stream input)
        {
            var header = await ReadHeaderAsync(input).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(header))
            {
                return null;
            }

            var lengthLine = header.Split('\n').FirstOrDefault(line => line.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase));
            if (lengthLine is null)
            {
                throw new McpRequestException(null, -32600, "MCP request missing Content-Length header.");
            }

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
            return JsonNode.Parse(json) as JsonObject
                ?? throw new McpRequestException(null, -32600, "MCP request must be a JSON object.");
        }

        private static async Task<string> ReadHeaderAsync(Stream input)
        {
            var bytes = new List<byte>();
            var lastFour = new Queue<byte>(4);
            while (true)
            {
                var b = new byte[1];
                var read = await input.ReadAsync(b, 0, 1).ConfigureAwait(false);
                if (read == 0)
                {
                    return string.Empty;
                }

                bytes.Add(b[0]);
                lastFour.Enqueue(b[0]);
                if (lastFour.Count > 4)
                {
                    lastFour.Dequeue();
                }

                if (lastFour.Count == 4 && lastFour.SequenceEqual(new byte[] { (byte)'\r', (byte)'\n', (byte)'\r', (byte)'\n' }))
                {
                    return Encoding.ASCII.GetString(bytes.ToArray());
                }
            }
        }

        private static async Task WriteMessageAsync(Stream output, JsonObject response)
        {
            var bytes = Encoding.UTF8.GetBytes(response.ToJsonString(McpJsonOptions));
            var header = Encoding.ASCII.GetBytes($"Content-Length: {bytes.Length}\r\n\r\n");
            await output.WriteAsync(header, 0, header.Length).ConfigureAwait(false);
            await output.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
            await output.FlushAsync().ConfigureAwait(false);
        }

        private sealed class McpRequestException : Exception
        {
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
