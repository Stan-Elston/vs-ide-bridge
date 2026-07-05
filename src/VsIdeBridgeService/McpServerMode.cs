using System.Diagnostics;
using System.Text.Json.Nodes;
using System.Text.Json;
using System.Net;
using System.Text;
using System.IO;
using VsIdeBridge.Shared;

namespace VsIdeBridgeService;

// Runs the MCP server JSON-RPC stdio loop.
// Launched when the service binary is invoked with the "mcp-server" argument.
internal static class McpServerMode
{
    private static readonly ToolExecutionRegistry Registry = ToolCatalog.Registry;
    private static readonly TimeSpan StdoutWriteTimeout = TimeSpan.FromSeconds(30);
    private const int BadRequestStatusCode = 400;
    private const int StdoutWriteTimeoutExitCode = 2;

    public static async Task RunAsync(string[] args)
    {
        using Stream input = Console.OpenStandardInput();
        using Stream output = Console.OpenStandardOutput();
        BridgeConnection bridge = new(args);
        McpToolSurface toolSurface = McpToolSurface.FromArgs(args);
        using StdioHostLease? hostLease = StdioHostLease.TryCreate();
        ServiceControlClient? controlClient = null;
        using SemaphoreSlim writeLock = new(1, 1);

        controlClient = await TryConnectControlPipeAsync().ConfigureAwait(false);

        McpServerLog.Write("stdio loop started");

        try
        {
            while (true)
            {
                McpProtocol.IncomingMessage? incoming;
                try
                {
                    incoming = await McpProtocol.ReadAsync(input).ConfigureAwait(false);
                }
                catch (McpRequestException ex)
                {
                    McpServerLog.Write($"read error code={ex.Code} message={ex.Message}");
                    await WriteLockedAsync(output, McpProtocol.ErrorResponse(ex.Id, ex.Code, ex.Message),
                        McpProtocol.WireFormat.HeaderFramed, writeLock).ConfigureAwait(false);
                    continue;
                }

                if (incoming is null)
                {
                    McpServerLog.Write("stdin closed; exiting stdio loop");
                    break;
                }

                StdioHostLease.MarkActivity();
                string incomingMethod = incoming.Request["method"]?.GetValue<string>() ?? string.Empty;

                // Handle notifications immediately without blocking the read loop.
                if (incomingMethod.StartsWith("notifications/", StringComparison.Ordinal))
                {
                    if (incomingMethod == "notifications/cancelled")
                    {
                        // Cancel the matching in-flight request's CTS so its polling loop exits.
                        JsonNode? cancelId = incoming.Request["params"]?["requestId"]?.DeepClone();
                        McpServerLog.Write($"cancel notification requestId={cancelId?.ToJsonString()}");
                        RequestCancellationRegistry.Cancel(cancelId);
                    }

                    continue;
                }

                if (controlClient is not null)
                {
                    try
                    {
                        await controlClient.NotifyRequestAsync().ConfigureAwait(false);
                    }
                    catch (IOException ex)
                    {
                        McpServerLog.Write($"control pipe request notify failed: {ex.Message}");
                        await controlClient.DisposeAsync().ConfigureAwait(false);
                        controlClient = null;
                    }
                    catch (ObjectDisposedException ex)
                    {
                        McpServerLog.Write($"control pipe request notify failed: {ex.Message}");
                        await controlClient.DisposeAsync().ConfigureAwait(false);
                        controlClient = null;
                    }
                    catch (InvalidOperationException ex)
                    {
                        McpServerLog.Write($"control pipe request notify failed: {ex.Message}");
                        await controlClient.DisposeAsync().ConfigureAwait(false);
                        controlClient = null;
                    }
                }

                McpServerLog.WriteRequest(incoming.Request, incoming.Format);

                // Dispatch on a background task so this read loop stays live for
                // notifications/cancelled while a long-running tool (e.g. build with
                // wait_for_completion=true) is in flight.
                _ = Task.Run(() => DispatchRequestAsync(incoming, output, bridge, controlClient, writeLock, toolSurface));
            }
        }
        finally
        {
            if (controlClient is not null)
            {
                await controlClient.DisposeAsync().ConfigureAwait(false);
            }
        }
    }

    // Isolated per-request dispatch method so RunAsync stays within line-length budget and
    // nesting depth stays shallow. Owns the CTS lifetime and the write-lock acquire.
    private static async Task DispatchRequestAsync(
        McpProtocol.IncomingMessage incoming,
        Stream output,
        BridgeConnection bridge,
        ServiceControlClient? controlClient,
        SemaphoreSlim writeLock,
        McpToolSurface toolSurface)
    {
        JsonNode? requestId = incoming.Request["id"]?.DeepClone();
        Stopwatch stopwatch = Stopwatch.StartNew();
        using IDisposable leaseActivity = StdioHostLease.BeginActivity();
        using CancellationTokenSource cts = RequestCancellationRegistry.Register(requestId);
        try
        {
            JsonObject? response = await HandleRequestAsync(incoming.Request, bridge, controlClient, toolSurface)
                .ConfigureAwait(false);

            if (response is not null)
            {
                McpServerLog.WriteResponse(response);
                await WriteLockedAsync(output, response, incoming.Format, writeLock)
                    .ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException)
        {
            McpServerLog.Write($"tool cancelled id={requestId?.ToJsonString()}");
            // No response sent for cancelled requests.
        }
        catch (Exception ex) when (ex is not null) // top-level dispatch boundary — fire-and-forget task
        {
            McpServerLog.Write($"dispatch fatal id={requestId?.ToJsonString()} ex={ex}");
        }
        finally
        {
            McpServerLog.Write($"dispatch finished id={FormatId(requestId)} ms={stopwatch.ElapsedMilliseconds}");
            RequestCancellationRegistry.Unregister(requestId);
        }
    }

    internal static async Task<JsonObject?> HandleRequestAsync(
        JsonObject request,
        BridgeConnection bridge,
        ServiceControlClient? controlClient,
        McpToolSurface? toolSurface = null)
    {
        JsonNode? id = request["id"]?.DeepClone();
        string method = request["method"]?.GetValue<string>() ?? string.Empty;
        JsonObject? @params = request["params"] as JsonObject;

        if (method.StartsWith("notifications/", StringComparison.Ordinal))
            return null;

        try
        {
            JsonNode result = method switch
            {
                "initialize"                => InitializeResult(@params, toolSurface ?? McpToolSurface.Lazy),
                "tools/list"                => new JsonObject { ["tools"] = Registry.BuildToolsList(toolSurface ?? McpToolSurface.Lazy) },
                "tools/call"                => await DispatchToolAsync(id, @params, bridge, controlClient).ConfigureAwait(false),
                "resources/list"            => EmptyResourcesList(),
                "resources/templates/list"  => EmptyResourceTemplatesList(),
                "ping"                      => new JsonObject(),
                _                           => throw new McpRequestException(
                    id, McpErrorCodes.MethodNotFound, $"Unsupported method: {method}"),
            };

            return new JsonObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = result };
        }
        catch (McpRequestException ex)
        {
            McpServerLog.Write($"request error method={method} code={ex.Code} message={ex.Message}");
            return McpProtocol.ErrorResponse(ex.Id ?? id, ex.Code, ex.Message);
        }
        catch (OperationCanceledException)
        {
            throw; // Propagate to the concurrent dispatch handler in RunAsync, which suppresses the response.
        }
        catch (Exception ex) when (ex is not null) // top-level MCP request boundary
        {
            McpServerLog.Write($"request fatal method={method} message={ex}");
            return McpProtocol.ErrorResponse(id, McpErrorCodes.MethodNotFound, ex.Message);
        }
    }

    private static JsonObject InitializeResult(JsonObject? @params, McpToolSurface toolSurface)
    {
        string clientVersion = @params?["protocolVersion"]?.GetValue<string>() ?? string.Empty;
        string negotiated = McpProtocol.SelectProtocolVersion(clientVersion);

        return new JsonObject
        {
            ["protocolVersion"] = negotiated,
            ["capabilities"] = new JsonObject
            {
                ["tools"] = new JsonObject(),
                ["resources"] = new JsonObject(),
            },
            ["serverInfo"] = new JsonObject
            {
                ["name"]    = "vs_ide_bridge",
                ["version"] = "0.1.0",
            },
            ["toolSurface"] = toolSurface.ToJsonObject(),
            ["instructions"] =
                BuildInstructions(toolSurface),
        };
    }

    private static string BuildInstructions(McpToolSurface toolSurface)
    {
        const string CallToolPrefix = "call_tool({";

        string surfaceText = toolSurface.IsFull
            ? "The full tool surface is exposed in tools/list. "
            : "A compact lazy tool surface is exposed in tools/list. " +
              "Use recommend_tools for task-based discovery, list_tools_by_category to browse a specific category, " +
              "or " + CallToolPrefix + "\"name\":\"list_tools\",...}) as a last resort to see every available tool name. ";

        return
            "VS IDE Bridge MCP server. " +
            "These tools are for YOU (the AI assistant) to call autonomously. " +
            "Do not ask the user which tools to use, do not present tool names or tool lists to the user, " +
            "and do not wait for user approval before calling a tool. " +
            "Interpret tool results and respond to the user in plain language. " +
            "When the user tells you to call a tool (e.g. 'call list_tools', 'use glob to find X', 'run errors') " +
            "they are directing your investigation — not asking you to display raw tool output. " +
            "Call the tool, use the results internally, then respond in plain language describing what you found or did. " +
            "Never dump raw JSON, tool lists, file paths, or unformatted output at the user. " +
            "Never output JSON in your response to the user unless the task itself is about writing or editing JSON — " +
            "tool results arrive as JSON internally but must always be translated into plain language before replying. " +
            surfaceText +
            "Visual Studio tab pressure rule: call list_tabs when the active instance may be sluggish or after opening several files. " +
            "If list_tabs reports more than 7 open tabs, close inactive tabs you no longer need with close_file, close_document, or close_others. " +
            "In lazy mode, ALL bridge tools must be called through call_tool with an arguments object — " +
            "including apply_diff, read_file, read_file_batch, write_file, list_tabs, close_file, close_document, close_others, " +
            "find_text, find_text_batch, search_symbols, file_outline, errors, build, build_solution, rebuild_solution, " +
            "build_errors, and git_status. " +
            "Use the wrapper shape " + CallToolPrefix + "\"name\":\"read_file\",\"arguments\":{\"file\":\"h:2\",\"start_line\":260,\"end_line\":360}}); " +
            "never pass raw strings as arguments. " +
            "Exact parameter names for the most-used tools (do NOT guess — use tool_help if unsure): " +
            "find_text: required=query, optional=path,scope,regex,whole_word,match_case,chunk_size; " +
            "find_files/glob: required=query, optional=path; " +
            "search_symbols: required=query, optional=kind,scope,path,project; " +
            "read_file: required=file, optional=start_line,end_line; " +
            "apply_diff (single edit): required=file,old_content,new_content, optional=replace_all(bool — replaces every occurrence atomically, one undo reverts all); " +
            "errors/warnings/messages: all optional: code,file,path,text,project,quick,refresh,group_by,sort_by,chunk_size. " +
            "When results include handle fields such as h:2, f:1, e:1, w:1, or m:1, " +
            "use the handle as the next file/path value instead of copying full paths. " +
            "Handles are in the JSON data rows (match.handle, row.handle) — they do NOT appear in the Summary string. " +
            "Always read the structured data fields, not just the Summary text. " +
            "For single-file edits: call read_file first (the response Data.handle field is an f:N handle), " +
            "then call apply_diff with that handle as file, the exact current text as old_content, and your replacement as new_content. " +
            "reserve diff for multi-file or structural patches. " +
            "Examples: " +
            CallToolPrefix + "\"name\":\"apply_diff\",\"arguments\":{\"file\":\"h:2\",\"old_content\":\"exact old text\",\"new_content\":\"replacement\"}}) " +
            "or " + CallToolPrefix + "\"name\":\"read_file\",\"arguments\":{\"file\":\"h:2\",\"start_line\":260,\"end_line\":360}}) " +
            "or " + CallToolPrefix + "\"name\":\"rebuild_solution\",\"arguments\":{}}) " +
            "or " + CallToolPrefix + "\"name\":\"errors\",\"arguments\":{}}) " +
            "or " + CallToolPrefix + "\"name\":\"git_status\",\"arguments\":{}}). " +
            "CRITICAL: For any file that belongs to the open VS solution, ALWAYS use bridge tools " +
            "(read_file, find_text, file_outline, search_symbols, find_references, goto_definition, apply_diff, write_file) — " +
            "NEVER use the host environment's own file-reading, grep, glob, shell, Edit, or Write tools on solution files. " +
            "The bridge has the live editor buffer, IntelliSense state, and semantic navigation that host tools cannot provide. " +
            "Search tool priority: " +
            "(1) When you know a symbol name, use search_symbols — not find_text — it returns h: handles and uses IntelliSense. " +
            "(2) Before reading an unfamiliar file, call file_outline first to get member line numbers, then read_file with start_line+end_line for just the relevant slice. " +
            "(3) Use glob (bridge) not the host Glob tool to find files by pattern. " +
            "(4) Use read_file_batch to read multiple slices across files in one call instead of repeated read_file calls. " +
            "(5) Use find_text_batch to search multiple patterns at once instead of repeated find_text calls. " +
            "Handle prefixes — always pass the handle as the file: value, never copy the full path: " +
            "h: is produced by find_text, find_text_batch, smart_context, search_symbols, find_references, goto_definition, goto_implementation, peek_definition, symbol_info; " +
            "f: is produced by find_files, glob, read_file, file_outline, file_symbols; " +
            "e:/w:/m: are produced by errors, warnings, messages, diagnostics_snapshot. " +
            "Use recommend_tools for task-based discovery and tool_help with name=<tool> for the full schema. " +
            "If you get an unknown-tool error, use recommend_tools or list_tools_by_category to find the right tool — do not guess tool names. " +
            "apply_diff failure rule: If apply_diff returns a content-mismatch error, NEVER fall back to write_file — that overwrites the entire file and destroys unrelated content. " +
            "Instead: call read_file on the same file (handle or path), use the f:N handle from the response Data.handle field, " +
            "copy the exact current text verbatim into old_content, and retry apply_diff with that handle. " +
            "Multi-instance: If more than one Visual Studio window is open and you get a binding error or are asked which solution to use, " +
            "call list_instances to see all running VS instances, then call bind_instance with the chosen instance_id (or bind_solution with a solution name hint). " +
            "You can switch to a different VS instance at any point mid-session by calling bind_instance again — " +
            "all tool calls apply only to the currently bound instance. " +
            "Save rule: before switching solutions, always ask the user if they want to save any unsaved work in the current solution first. " +
            "If the target solution is not in list_instances at all, it is not open in VS yet — " +
            "use glob with pattern **/*.sln and set the path to one level above the current solution root (the repos root) to search sibling repositories on disk. " +
            "Tell the user which .sln you found and offer two options: " +
            "(a) open it in a new VS window — both solutions stay open and you can switch between them at any time using bind_instance, or " +
            "(b) swap it into the current VS window using open_solution — this closes the current solution in that window. " +
            "Once the target solution is open in VS, call list_instances to confirm it appeared, then bind_instance with the new instance_id. " +
            "Diagnostics rule: After making any code edits, call warnings({ refresh=true, file=<edited file> }) and errors({ refresh=true, file=<edited file> }) " +
            "to get a fresh live read — the default cache is stale after edits, so always pass refresh=true when checking post-edit diagnostics. " +
            "Errors you introduced must always be fixed immediately — do not ask the user, just fix them. " +
            "Warnings and messages you introduced (in files you touched): fix them immediately without asking — they are your responsibility. " +
            "Warnings and messages in files you did not edit (pre-existing): show a summary to the user and ask if they want them addressed — do not auto-fix. " +
            "To distinguish new from pre-existing: any diagnostic in a file you edited that appears after your change is yours; if you cannot tell, note it and ask. " +
            "Math rule: NEVER compute non-trivial arithmetic mentally or guess at a numeric result. " +
            "Always call python_eval with the expression — it runs real Python and returns the exact answer. " +
            "Both python_eval and python_exec expose full builtins and allow import; pre-imported for convenience: math, statistics, decimal, fractions, re, itertools, collections. " +
            "Example: " + CallToolPrefix + "\"name\":\"python_eval\",\"arguments\":{\"expression\":\"math.sqrt(2)\"}}).";
    }

    private static JsonObject EmptyResourcesList()
    {
        return new JsonObject
        {
            ["resources"] = new JsonArray(),
        };
    }

    private static JsonObject EmptyResourceTemplatesList()
    {
        return new JsonObject
        {
            ["resourceTemplates"] = new JsonArray(),
        };
    }

    private static async Task<JsonNode> DispatchToolAsync(
        JsonNode? id, JsonObject? @params, BridgeConnection bridge, ServiceControlClient? controlClient)
    {
        string toolName = @params?["name"]?.GetValue<string>() ?? string.Empty;
        JsonObject? args = @params?["arguments"] as JsonObject;

        if (string.IsNullOrWhiteSpace(toolName))
            throw new McpRequestException(id, McpErrorCodes.InvalidParams, "Missing 'name' in tools/call params.");

        McpServerLog.Write($"dispatch tool={toolName}");

        if (controlClient is null)
        {
            return await Registry.DispatchAsync(id, toolName, args, bridge).ConfigureAwait(false);
        }

        await controlClient.NotifyCommandStartAsync().ConfigureAwait(false);
        try
        {
            return await Registry.DispatchAsync(id, toolName, args, bridge).ConfigureAwait(false);
        }
        finally
        {
            await controlClient.NotifyCommandEndAsync().ConfigureAwait(false);
        }
    }

    private static async Task<ServiceControlClient?> TryConnectControlPipeAsync()
    {
        try
        {
            return await ServiceControlClient.ConnectAsync().ConfigureAwait(false);
        }
        catch
        {
            return null;
        }
    }

    // Serialize stdout writes across concurrent request tasks.
    private static async Task WriteLockedAsync(
        Stream output, JsonObject response, McpProtocol.WireFormat format, SemaphoreSlim writeLock)
    {
        string responseId = FormatId(response["id"]);
        Stopwatch stopwatch = Stopwatch.StartNew();
        McpServerLog.Write($"response write queued id={responseId} format={format}");

        bool lockTaken = await writeLock.WaitAsync(StdoutWriteTimeout).ConfigureAwait(false);
        if (!lockTaken)
        {
            McpServerLog.Write(
                $"response write lock timed out id={responseId} after {StdoutWriteTimeout.TotalSeconds:F0}s; exiting stdio host");
            Environment.Exit(StdoutWriteTimeoutExitCode);
            return;
        }

        try
        {
            McpServerLog.Write($"response write start id={responseId}");
            Task writeTask = McpProtocol.WriteAsync(output, response, format);
            Task timeoutTask = Task.Delay(StdoutWriteTimeout);
            Task completedTask = await Task.WhenAny(writeTask, timeoutTask).ConfigureAwait(false);
            if (!ReferenceEquals(completedTask, writeTask))
            {
                McpServerLog.Write(
                    $"response write timed out id={responseId} after {StdoutWriteTimeout.TotalSeconds:F0}s; exiting stdio host because stdout did not drain");
                Environment.Exit(StdoutWriteTimeoutExitCode);
                return;
            }

            await writeTask.ConfigureAwait(false);
            McpServerLog.Write($"response write complete id={responseId} ms={stopwatch.ElapsedMilliseconds}");
        }
        catch (IOException ex)
        {
            McpServerLog.WriteException($"response write failed id={responseId}; exiting stdio host", ex);
            Environment.Exit(StdoutWriteTimeoutExitCode);
        }
        catch (ObjectDisposedException ex)
        {
            McpServerLog.WriteException($"response write failed id={responseId}; exiting stdio host", ex);
            Environment.Exit(StdoutWriteTimeoutExitCode);
        }
        catch (InvalidOperationException ex)
        {
            McpServerLog.WriteException($"response write failed id={responseId}; exiting stdio host", ex);
            Environment.Exit(StdoutWriteTimeoutExitCode);
        }
        finally
        {
            writeLock.Release();
        }
    }

    private static string FormatId(JsonNode? id)
    {
        return id is null ? "<null>" : id.ToJsonString();
    }

    public static async Task RunHttpAsync(string[] args, CancellationToken cancellationToken = default)
    {
        int port = HttpServerDefaults.HttpPort;
        for (int i = 0; i < args.Length - 1; i++)
        {
            if (string.Equals(args[i], "--port", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(args[i + 1], out int parsedPort) && parsedPort > 0 && parsedPort < 65536)
                    port = parsedPort;
                break;
            }
        }

        string prefix = $"http://localhost:{port}/";
        BridgeConnection bridge = new(args);
        McpToolSurface toolSurface = McpToolSurface.FromArgs(args);

        McpServerLog.Write($"MCP HTTP server starting on {prefix}");

        using HttpListener listener = new();
        listener.Prefixes.Add(prefix);
        try
        {
            listener.Start();
            McpServerLog.Write("HTTP listener started successfully");
        }
        catch (HttpListenerException ex)
        {
            McpServerLog.Write($"Failed to start listener: {ex}");
            throw;
        }
        catch (ObjectDisposedException ex)
        {
            McpServerLog.Write($"Failed to start listener: {ex}");
            throw;
        }

        // When the cancellationToken fires, stop the listener so GetContextAsync throws
        // HttpListenerException with IsListening == false and the accept loop exits cleanly.
        using CancellationTokenRegistration stopReg = cancellationToken.Register(
            static state => ((HttpListener)state!).Stop(), listener);

        while (true)
        {
            HttpListenerContext? context = null;
            try
            {
                context = await listener.GetContextAsync().ConfigureAwait(false);
                await HandleHttpRequestAsync(context, bridge, toolSurface).ConfigureAwait(false);
            }
            catch (HttpListenerException) when (listener.IsListening == false)
            {
                break;
            }
            catch (HttpListenerException ex)
            {
                McpServerLog.Write($"HTTP accept error: {ex}");
            }
            catch (IOException ex)
            {
                McpServerLog.Write($"HTTP accept error: {ex}");
            }
            catch (ObjectDisposedException ex)
            {
                McpServerLog.Write($"HTTP accept error: {ex}");
            }
        }
    }

    private static async Task HandleHttpRequestAsync(HttpListenerContext context, BridgeConnection bridge, McpToolSurface toolSurface)
    {
        try
        {
            string path = context.Request.Url?.AbsolutePath ?? "/";
            string method = context.Request.HttpMethod;

            McpServerLog.Write($"HTTP {method} {path}");

            if (method == "GET")
            {
                await WriteHealthResponseAsync(context).ConfigureAwait(false);
                return;
            }

            if (method != "POST")
            {
                context.Response.StatusCode = 405;
                context.Response.ContentType = "text/plain";
                await context.Response.OutputStream.WriteAsync("Method not allowed"u8.ToArray()).ConfigureAwait(false);
                return;
            }

            JsonObject? request = await ReadHttpRequestAsync(context).ConfigureAwait(false);
            if (request is null)
            {
                return;
            }

            McpServerLog.WriteRequest(request, McpProtocol.WireFormat.RawJson);

            JsonObject? response = await HandleRequestAsync(request, bridge, controlClient: null, toolSurface).ConfigureAwait(false);

            if (response is not null)
            {
                McpServerLog.WriteResponse(response);
                string jsonResponse = response.ToJsonString();
                byte[] bytes = Encoding.UTF8.GetBytes(jsonResponse);
                context.Response.ContentType = "application/json; charset=utf-8";
                context.Response.ContentLength64 = bytes.Length;
                await context.Response.OutputStream.WriteAsync(bytes).ConfigureAwait(false);
            }
            else
            {
                context.Response.StatusCode = 204; // no content for notifications
            }
        }
        catch (McpRequestException ex)
        {
            McpServerLog.Write($"HTTP MCP error code={ex.Code}: {ex.Message}");
            await WriteErrorResponse(context, ex.Message, ex.Code);
        }
        catch (Exception ex) when (ex is not null) // top-level HTTP request boundary
        {
            McpServerLog.Write($"HTTP request fatal: {ex}");
            await WriteErrorResponse(context, ex.Message);
        }
        finally
        {
            try
            {
                context.Response.Close();
            }
            catch (ObjectDisposedException ex) { McpServerLog.WriteException("HTTP response close skipped because the response was already disposed", ex); }
            catch (HttpListenerException ex) { McpServerLog.WriteException("HTTP response close skipped because the listener was already stopping", ex); }
        }
    }

    private static async Task WriteHealthResponseAsync(HttpListenerContext context)
    {
        JsonObject info = new()
        {
            ["name"] = "vs-ide-bridge",
            ["version"] = "0.1.0",
            ["protocolVersions"] = new JsonArray { "2025-03-26", "2024-11-05" },
            ["capabilities"] = new JsonObject { ["tools"] = new JsonObject() },
            ["status"] = "ok"
        };

        await WriteJsonResponseAsync(context, info).ConfigureAwait(false);
    }

    private static async Task<JsonObject?> ReadHttpRequestAsync(HttpListenerContext context)
    {
        string body = await ReadRequestBodyAsync(context.Request).ConfigureAwait(false);
        if (string.IsNullOrWhiteSpace(body))
        {
            context.Response.StatusCode = BadRequestStatusCode;
            return null;
        }

        try
        {
            JsonNode? node = JsonNode.Parse(body);
            return node as JsonObject ?? throw new InvalidOperationException("Not an object");
        }
        catch (InvalidOperationException)
        {
            context.Response.StatusCode = BadRequestStatusCode;
            await WriteErrorResponse(context, "Invalid JSON").ConfigureAwait(false);
            return null;
        }
        catch (FormatException)
        {
            context.Response.StatusCode = BadRequestStatusCode;
            await WriteErrorResponse(context, "Invalid JSON").ConfigureAwait(false);
            return null;
        }
        catch (JsonException)
        {
            context.Response.StatusCode = BadRequestStatusCode;
            await WriteErrorResponse(context, "Invalid JSON").ConfigureAwait(false);
            return null;
        }
    }

    private static async Task<string> ReadRequestBodyAsync(HttpListenerRequest request)
    {
        using StreamReader reader = new(request.InputStream, request.ContentEncoding ?? Encoding.UTF8);
        return await reader.ReadToEndAsync().ConfigureAwait(false);
    }

    private static async Task WriteJsonResponseAsync(HttpListenerContext context, JsonObject payload)
    {
        string json = payload.ToJsonString();
        byte[] bytes = Encoding.UTF8.GetBytes(json);
        context.Response.ContentType = "application/json; charset=utf-8";
        context.Response.ContentLength64 = bytes.Length;
        await context.Response.OutputStream.WriteAsync(bytes).ConfigureAwait(false);
    }

    private static async Task WriteErrorResponse(HttpListenerContext context, string message, int code = -32603)
    {
        JsonObject errorResponse = new()
        {
            ["jsonrpc"] = "2.0",
            ["id"] = null,
            ["error"] = new JsonObject { ["code"] = code, ["message"] = message }
        };
        context.Response.StatusCode = 200;
        await WriteJsonResponseAsync(context, errorResponse).ConfigureAwait(false);
    }
}
