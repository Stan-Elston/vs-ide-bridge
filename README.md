# VS IDE Bridge

VS IDE Bridge connects your AI assistant to Visual Studio. Once connected, your assistant can read and edit code, check errors, run builds, inspect debugger state, and work with Git — all against the live project that is already open on your machine.

## Requirements

- Windows 10 or Windows 11
- Visual Studio 2022 (17.14 or later) or Visual Studio 2026 (18.x) — Community, Professional, or Enterprise
- An AI assistant with MCP support (Claude Code, Grok, Cursor, Codex, LM Studio, or any MCP-compatible client)

## Installation

1. **Download the installer.** Get the latest `vs-ide-bridge-setup-<version>.exe` from the [Releases page](https://github.com/RenegadeRiff86/Visual-Studio-MCP/releases/latest).
2. **Close Visual Studio.** The installer replaces the VS extension, which can't be updated while VS is running.
3. **Run the installer** (it asks for administrator rights). It installs three things: the Visual Studio extension, the background MCP service (**VS IDE Bridge Service**, a Windows service), and a bundled Python runtime — no separate Python install needed.
4. **Open Visual Studio** and load your solution. The extension activates automatically. To confirm it's installed, check **Extensions → Manage Extensions → Installed**, or look for the **Tools → VS IDE Bridge** menu.
5. **Connect your AI assistant** using the setup below.

> **Upgrading** is the same steps: close Visual Studio, run the newer installer, reopen. Your assistant's connection settings carry over — if the tools stop appearing after an upgrade, re-run the registration command for your client and restart it.

## Connecting Your AI Assistant

Pick the connection method that matches your client.

### Claude Code (stdio)

```bash
# Register for the current project
claude mcp add --transport stdio --scope project vs-ide-bridge -- "C:\Program Files\VsIdeBridge\service\VsIdeBridgeService.exe" mcp-server

# Or register globally (available from any project)
claude mcp add --transport stdio --scope user vs-ide-bridge -- "C:\Program Files\VsIdeBridge\service\VsIdeBridgeService.exe" mcp-server
```

Restart Claude Code after adding. Visual Studio must be running before you start a session.

If the server stops appearing after a reinstall, re-run the command above and restart Claude Code.

### Grok / xAI (stdio)

Grok uses the same stdio transport as Claude Code. Visual Studio must be running before you start a session.

**Grok CLI** — register globally via command or config:

```bash
grok mcp add vs-ide-bridge --command "C:\Program Files\VsIdeBridge\service\VsIdeBridgeService.exe" --args mcp-server
```

```toml
# ~/.grok/config.toml
[mcp_servers.vs-ide-bridge]
command = 'C:\Program Files\VsIdeBridge\service\VsIdeBridgeService.exe'
args = ["mcp-server"]
enabled = true
```

Verify the connection with `grok mcp doctor vs-ide-bridge`, then restart Grok.

**Grok in Cursor** — use the Cursor stdio config in the section below. Restart Cursor after adding the server.

### Cursor (stdio)

Cursor can also connect directly over stdio using the same service executable:

```json
// ~/.cursor/mcp.json
{
  "mcpServers": {
    "vs-ide-bridge": {
      "command": "C:\\Program Files\\VsIdeBridge\\service\\VsIdeBridgeService.exe",
      "args": ["mcp-server"]
    }
  }
}
```

Restart Cursor after adding. Visual Studio must be running before you start a session.

### HTTP (Codex and most web-based clients)

The bridge listens on `http://localhost:43117/` by default.

```toml
# Codex — ~/.codex/config.toml
[mcp_servers.vs-ide-bridge]
url = "http://localhost:43117/"
enabled = true
```

```json
// Generic HTTP MCP config
{
  "mcpServers": {
    "vs-ide-bridge": {
      "transport": { "type": "http", "url": "http://localhost:43117/" }
    }
  }
}
```

If the endpoint is off, enable it from Visual Studio: **Tools → VS IDE Bridge → Toggle HTTP MCP Server**

### Streamable HTTP (LM Studio and local models)

```json
{
  "mcpServers": {
    "vs-ide-bridge": {
      "transport": { "type": "streamable-http", "url": "http://localhost:43118/mcp" }
    }
  }
}
```

Enable from Visual Studio: **Tools → VS IDE Bridge → Toggle Streamable HTTP MCP Server**

## What You Can Do

Once connected, describe what you want in plain language. The assistant figures out how to use the bridge.

**Navigate code**
- "Find all the places where `ProcessOrder` is called."
- "What does this method do? Show me its definition."
- "Show me everything that calls into this class."

**Make changes**
- "Rename this variable everywhere it's used."
- "Extract this block into a new helper method."
- "Add null checks to all the public methods in this file."

**Diagnose and build**
- "What errors are in the project right now?"
- "Fix all the warnings in this file."
- "Build the solution and tell me what broke."

**Debug**
- "Set a breakpoint on line 42 and start the debugger."
- "Break when execution reaches the function `ToolResultFormatter.StructuredToolResult`."
- "Log `requestId` every time that line runs, but don't stop."
- "What are the local variables at the current line?"
- "Show me the call stack."
- "Expand `m_buffer` and show me every element and its fields."
- "Show me the exception the debugger just hit."

Breakpoints can target a function or symbol name instead of a file and line — these survive source edits and line shifts — and can log a message and keep running instead of pausing (tracepoints/logpoints). `debug_locals` and `debug_watch` can expand a container or struct into its elements and fields; use `chunk_lines` on `debug_watch` to page large expansions in memory without flooding the reply. Debugger run and step tools include `lastException` when Visual Studio reports a thrown or unhandled exception; for native C++ exceptions, the VSIX also listens for lower-level debugger exception events. Ask for `debug_exceptions` when you need the exception settings snapshot and the latest captured exception details.

**Git**
- "What files have I changed since the last commit?"
- "Commit everything with a sensible message."
- "Show me the diff for this file."

**GitHub** (requires the `gh` CLI, authenticated)
- "List the open issues and show me issue 42."
- "Comment on issue 42, add the `bug` label, and close it as completed."
- "Open an issue titled 'Fix flaky cooling test' with the `bug` label."
- "Show me the open pull requests and the diff for PR 17."

**Project and packages**
- "What NuGet packages does this project reference?"
- "Add the latest version of Newtonsoft.Json."

## Tips

- **Open Visual Studio first.** The bridge connects to a running VS instance. If VS is not open, the assistant cannot see your code.
- **Multiple VS windows?** The bridge handles this — see [Working with Multiple VS Instances](#working-with-multiple-vs-instances) below.
- **After a build error?** Ask *"what errors came up?"* instead of reading the Output window yourself — the assistant gets a structured list it can act on immediately.
- **Slow start?** IntelliSense takes a moment to load after VS opens. If the assistant reports incomplete symbol results, wait a few seconds and try again.

## Working with Multiple VS Instances

You can have any number of Visual Studio windows open at the same time. The bridge discovers all of them automatically — each VS instance runs its own copy of the extension.

**When only one instance is open** the assistant binds to it automatically at the start of a session. No extra steps needed.

**When more than one instance is open** the bridge rebinds to the solution it was using last time if that solution is still open; otherwise it auto-binds to the first instance it discovers and includes a warning listing the other open instances. If it picked the wrong one — or you want to be explicit up front — just tell the assistant:

- *"Bind to MySolution.sln."*
- *"Switch to the CodeMaid project."*
- *"Connect to the VS instance with OrderService open."*

The assistant stays bound to that instance for the rest of the session. All tool calls — reads, edits, builds, errors — apply to the currently bound instance only.

**Switching mid-session** is supported. Ask the assistant to bind to a different solution at any point and it will switch immediately. This is useful when a change in one project requires a follow-up in another.

**Typical setup** — two VS windows side by side:

| Window | Purpose |
|---|---|
| `MySolution.sln` | The project you are working on |
| `VsIdeBridge.sln` | Bridge source — open only if you are developing the bridge itself |

The assistant works against whichever one it is bound to. Binding is session-scoped and has no effect on the other open instances.

## Seeing What the Model Is Doing

The bridge streams a trace of every tool call into the Visual Studio **Output** window in real time. To watch it:

1. In Visual Studio, open **View → Output** (or press **Ctrl+Alt+O**).
2. Click the **Show output from** dropdown at the top of the Output pane.
3. Select **IDE Bridge**.

Each line shows the command name, key arguments, and result status as the model works — useful for understanding what the assistant is doing or for diagnosing unexpected behaviour.

## Best-Practice Warnings

VS IDE Bridge runs a lightweight best-practice analyzer on your code and reports issues (prefixed `BP`) in the Visual Studio Error List alongside normal compiler diagnostics. This is on by default.

To turn it off (or back on), use **Tools → VS IDE Bridge → Toggle Best-Practice Warnings**. The setting persists across Visual Studio restarts — no configuration file to edit.

## Logs

All logs are written to the same directory. The location is resolved in this order:

| Scenario | Path |
|---|---|
| Installed (default) | `C:\Program Files\VsIdeBridge\logs\` |
| No installer, `%COMMONAPPDATA%` present | `C:\ProgramData\VsIdeBridge\logs\` |
| Fallback | `%TEMP%\VsIdeBridge\logs\` |

The installed path is read from the registry key `HKLM\SOFTWARE\VsIdeBridge\InstallPath` written by the installer. Both the Windows service and the Visual Studio extension use `BridgeLogPaths.GetSharedLogDirectory()` (in `src/Shared/BridgeLogPaths.cs`) so they always land in the same folder.

Two files are written:

- `mcp-server.log` — MCP request/response traffic (written by the Windows service)
- `vs-ide-bridge-yyyy-MM-dd.log` — extension warnings and errors (written by the VS extension, one file per day)

**Retention** — logs are managed automatically. `mcp-server.log` is rotated to `mcp-server.log.old` when it exceeds 5 MB (one backup kept, ~10 MB cap). Daily extension logs older than 7 days are deleted on the first write of each VS session.

> **Developer note** — when running from source without running the installer first, the registry key is absent and logs go to `C:\ProgramData\VsIdeBridge\logs\`. If you need to change the retention limits or add new log files, all path and cleanup logic is centralised in `src/Shared/BridgeLogPaths.cs`.

## Troubleshooting

**Assistant says it can't find Visual Studio or the solution**
Visual Studio must be running with your solution open before you start the AI session. Try telling your assistant: *"Connect to Visual Studio"* or *"Bind to MySolution.sln."*

**Tools stopped working after a reinstall**
Re-run the registration command for your client (see [Connecting Your AI Assistant](#connecting-your-ai-assistant)), then restart the client.

**Bridge is not responding**
Check that the `VsIdeBridgeService` Windows service is running. Open Services (`services.msc`) and look for **VS IDE Bridge Service**.

## Shell Safety

The bridge can run arbitrary shell commands on your machine when asked. This is intentional for advanced tasks, but treat it as a power tool:

- Prefer asking for typed actions (edit this file, run a build, commit this change) before asking for raw shell commands.
- Do not use unattended automation that includes shell commands without review.
- If your MCP client supports tool approval rules, add `shell_exec` to the requires-approval list.

## Building from Source

Prerequisites:

- Visual Studio 2026 (18.x) with the **Visual Studio extension development** workload (the solution builds against the VS 2026 SDK; the produced VSIX installs on both VS 2022 17.14+ and VS 2026)
- [Inno Setup 6](https://jrsoftware.org/isinfo.php) — download and install it separately if you want to produce the installer; it is not bundled and not needed for code-only builds

```bash
git clone https://github.com/RenegadeRiff86/Visual-Studio-MCP
cd Visual-Studio-MCP
scripts\build.bat Release
```

`build.bat` locates MSBuild through vswhere and rebuilds the full solution including the VSIX. (Plain `dotnet build` does not work here: the solution contains a C++ probe project that only builds under Visual Studio's MSBuild.)

Open `VsIdeBridge.sln` in Visual Studio to work on the extension. To produce the installer, compile the Inno Setup script after a Release build:

```bash
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\inno\vs-ide-bridge.iss
```

which writes:

```
installer\output\vs-ide-bridge-setup-<version>.exe
```

When developing without the installer, logs are written to `C:\ProgramData\VsIdeBridge\logs\` instead of the default install directory. See [Logs](#logs) for the full path resolution and retention details.

## Contributors

See [CONTRIBUTORS.md](CONTRIBUTORS.md) for a list of people who have contributed code to this project.

## Third-Party Notices

See [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) for license information covering third-party components used in this project.
