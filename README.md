# VS IDE Bridge

Visual Studio extension that exposes scriptable IDE commands for external automation over a named pipe. The bridge starts automatically when Visual Studio starts, and the native CLI can enumerate, launch, and target specific live IDE instances.

## Disclaimer

This project is experimental. Use it at your own risk.

## Purpose

Lets you drive a running Visual Studio instance from outside the IDE: search code, navigate to symbols, slice documents, apply diffs, control the debugger, and capture build output — all without touching the keyboard.

Commands are invoked through simple pipe names like `state`, `search-symbols`, and `quick-info`. The legacy `Tools.Ide*` names still work for compatibility. The native CLI can print `json`, `summary`, or `keyvalue` to stdout and can also write envelopes to a caller-specified output file.

The CLI also includes a Windows-side stdio MCP facade (`vs-ide-bridge mcp-server`) that forwards MCP tool/resource/prompt requests to the existing bridge pipe server. This keeps the VSIX bridge as the source of truth while exposing an LLM-friendly command surface.

## LLM Workflow

Use this five-step pattern:

1. `vs-ide-bridge help`
2. `vs-ide-bridge ensure --solution C:\path\to\Your.sln`
3. copy the returned `instanceId`
4. `vs-ide-bridge catalog --instance <instanceId>`
5. `vs-ide-bridge search-symbols --instance <instanceId> --query RunAsync`

If Visual Studio is already open on the right solution, `ensure` reuses it. If more than one Visual Studio bridge instance is live, run `vs-ide-bridge instances` and then use `--instance`.

If you need task-oriented examples, run `vs-ide-bridge prompts`.
If you need to extract one field from a saved bridge result, run `vs-ide-bridge parse`.

## Requirements

- Windows
- Visual Studio 2026 / 18

The extension targets VS 18. The build script probes the `Community` install path first and then falls back to the VS 2022 Community path. If you use Professional or Enterprise, adjust that path in `scripts\build.bat`.

## Build And Install

Close all Visual Studio instances before updating the extension. If `devenv.exe` is still running, `VsIdeBridge.dll` can stay locked and the install step will fail even when the build succeeds.

### Preferred Full Build

```bat
scripts\build.bat
```

This is the only build script. It builds the solution, including:

- the VSIX package in `src\VsIdeBridge\bin\<Configuration>\net472\VsIdeBridge.vsix`
- the native pipe client in `src\VsIdeBridgeCli\bin\<Configuration>\net8.0\vs-ide-bridge.exe`

Debug is the default configuration. Pass a configuration name as the first argument:

```bat
scripts\build.bat Release
```

### Managed-Only Fallback Build

If the full solution build fails in an unrelated native project, such as `src\IdeBridgeJsonProbe\IdeBridgeJsonProbe.vcxproj`, build just the bridge extension and CLI:

```bat
"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" src\VsIdeBridge\VsIdeBridge.csproj /restore /p:Configuration=Debug
dotnet build src\VsIdeBridgeCli\VsIdeBridgeCli.csproj -c Debug
```

This is enough to produce:

- `src\VsIdeBridge\bin\Debug\net472\VsIdeBridge.vsix`
- `src\VsIdeBridge\bin\Debug\net472\VsIdeBridge.dll`
- `src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe`

### Install Or Update The Extension

**Prerequisites — before running the installer:**

1. Close Visual Studio (`send --command close-ide` via bridge, or File → Exit).
2. Kill any lingering helper processes — `clangd.exe`, `ServiceHub.RoslynCodeAnalysisService.exe`, `DevHub.exe` — they block the installer even after VS exits.
3. If the version in `source.extension.vsixmanifest` has not changed since the last install, bump it (e.g. `0.1.0` → `0.1.1`); the installer exits 209 ("already installed") otherwise and copies nothing.

**Install command — use PowerShell, not `cmd.exe /c`:**

`cmd.exe /c` silently swallows errors and returns exit 0 without running the installer. Use PowerShell `Start-Process` with an argument array instead:

```powershell
powershell.exe -Command "
  \$p = Start-Process \`
    -FilePath 'C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\VSIXInstaller.exe' \`
    -ArgumentList @('/quiet', '/instanceIds:8fd42dc7', 'C:\path\to\VsIdeBridge.vsix') \`
    -PassThru -Wait
  Write-Host \"Exit: \$(\$p.ExitCode)\"
"
```

Exit 0 = success. A new log appears in `%TEMP%\dd_VSIXInstaller_*.log` and a new folder under `%LOCALAPPDATA%\Microsoft\VisualStudio\18.0_8fd42dc7\Extensions\` contains `VsIdeBridge.dll`.

**Direct-copy fallback** (when the installer is not cooperating):

1. Find the installed folder under `%LOCALAPPDATA%\Microsoft\VisualStudio\18.0_*\Extensions\`.
2. Copy these files from `src\VsIdeBridge\bin\Debug\net472\` into that folder while VS is closed:
   - `VsIdeBridge.dll`
   - `VsIdeBridge.pkgdef`
   - `extension.vsixmanifest`

Start Visual Studio after either method; it picks up the updated extension on the next load.

### Validate The Install

After install, verify the bridge with:

```bat
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe ensure --solution "C:\path\to\Your.sln"
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe current --format keyvalue
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe catalog --instance <instanceId>
```

## Start The Bridge

Preferred:

```bat
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe ensure --solution "C:\path\to\Your.sln"
```

Optional PowerShell wrapper:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\start_bridge.ps1 `
  -SolutionPath "C:\path\to\Your.sln"
```

After the VSIX is installed, the bridge starts automatically with Visual Studio. The native `ensure` command:

- reuses an already-running bridge if it matches the solution
- otherwise starts Visual Studio with the requested solution
- waits until the bridge responds over the named pipe
- exits with code `0` on success and `1` on failure

The PowerShell script is now just a thin wrapper over `vs-ide-bridge ensure`.

Bridge UI defaults:

- pipe activity writes to the `IDE Bridge` Output pane and updates the Visual Studio status bar
- `IDE Bridge > Allow Bridge Edits` is off by default
- `IDE Bridge > Go To Edited Parts` is on by default

That means reads and diagnostics are visible in VS immediately, but `apply-diff` will refuse to change files until you explicitly allow bridge edits from the menu.

Bridge edits now apply through the live editor buffer with temporary line markers: added/modified regions are color-highlighted, and deletions are shown as deletion markers so review can happen directly in Visual Studio before save.

Optional parameters:

- `-Configuration Debug|Release`
- `-StartupTimeoutSeconds 180`
- `-PollIntervalMilliseconds 1000`
- `-OutputPath C:\temp\ide-state.json`
- `-SkipWaitForReady`

## Quick Start

1. Close Visual Studio.
2. Build with `scripts\build.bat`.
3. If that fails in a non-bridge native project, use the managed-only fallback build above.
4. Install `src\VsIdeBridge\bin\Debug\net472\VsIdeBridge.vsix`.
5. Run `src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe ensure --solution "C:\path\to\Your.sln"`.
6. Run `src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe current`.
7. Run `src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe help` if you need a refresher.
8. Send commands with `src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe`.

## Command Surface

The tables below list the preferred simple pipe names. The live `catalog` command also returns the legacy `Tools.Ide*` names for compatibility.

### Core

| Command | Description |
|---------|-------------|
| `help` | List all registered commands |
| `state` | Snapshot of IDE state (solution path, active document, etc.) |
| `ready` | Block until IntelliSense is available |
| `open-solution` | Open a solution file |
| `close` | Close the targeted Visual Studio instance gracefully |
| `batch` | Execute multiple commands in one pipe round-trip |

### Search and Navigation

| Command | Description |
|---------|-------------|
| `find-text` | Text search across solution or project |
| `find-files` | Find files by name |
| `parse` | Parse saved JSON locally with slash-path selection |
| `document-slice` | Fetch lines around a location |
| `document-slices` | Fetch multiple document ranges from a JSON file or inline JSON |
| `file-symbols` | Symbol list for a file, with optional kind filtering |
| `file-outline` | Alias for `file-symbols` |
| `search-symbols` | Search symbols by name across the current scope |
| `quick-info` | Resolve symbol/definition info at a location |
| `smart-context` | Multi-context gather for agent queries |
| `find-references` | Semantic find-all-references |
| `call-hierarchy` | Callers and callees |
| `goto-definition` | Navigate to definition (F12) |
| `peek-definition` | Peek definition inline |
| `goto-implementation` | Navigate to implementation |
| `open-document` | Open a file at a line/column |
| `list-documents` | List open documents |
| `list-tabs` | List open editor tabs |
| `activate-document` | Activate a document by name |
| `close-document` | Close a document by name |
| `close-file` | Close a document by full path |
| `close-others` | Close all tabs except the active one |
| `activate-window` | Activate a tool window |
| `list-windows` | List tool windows |
| `execute-command` | Execute any native VS command by name |
| `apply-diff` | Apply a unified diff through the live VS editor |

### Breakpoints

> These are pipe commands, not first-class CLI verbs. Use `send --command <name>` to invoke them.

| Command | Description |
|---------|-------------|
| `set-breakpoint` | Set a breakpoint at file/line |
| `list-breakpoints` | List all breakpoints |
| `remove-breakpoint` | Remove a breakpoint |
| `clear-breakpoints` | Remove all breakpoints |
| `enable-breakpoint` | Enable a breakpoint at file/line |
| `disable-breakpoint` | Disable a breakpoint at file/line |
| `enable-all-breakpoints` | Enable every breakpoint |
| `disable-all-breakpoints` | Disable every breakpoint |

### Debugger

> These are pipe commands, not first-class CLI verbs. Use `send --command <name>` to invoke them.

| Command | Description |
|---------|-------------|
| `debug-state` | Current debugger state |
| `debug-start` | Start debugging |
| `debug-stop` | Stop debugging |
| `debug-break` | Break into the debugger |
| `debug-continue` | Continue execution |
| `debug-step-over` | Step over |
| `debug-step-into` | Step into |
| `debug-step-out` | Step out |

### Build and Diagnostics

| Command | Description |
|---------|-------------|
| `build` | Build the solution |
| `warnings` | Capture warnings from the Error List, or Build Output fallback, as JSON |
| `errors` | Capture errors from the Error List, or Build Output fallback, as JSON |
| `build-errors` | Build then capture diagnostics in one call |

## Argument Contract

- `--out "C:\path\result.json"` — output path (falls back to `%TEMP%\vs-ide-bridge\<command>.json`)
- `--request-id "abc123"` — optional correlation id echoed in the envelope
- `--timeout-ms 120000` — timeout on wait/build commands
- boolean flags: `--flag` (bare, implies `true`) or `--flag true` / `--flag false`
- enum values: lowercase kebab-case

## Command Examples

```text
ready --timeout-ms 120000 --out "C:\temp\ready.json"
state --out "C:\temp\state.json"
open-solution --solution "C:\path\to\your.sln" --out "C:\temp\open.json"

find-files --query "GUI_App.cpp" --out "C:\temp\files.json"
find-text --query "OnInit" --scope solution --path "src\libslic3r" --out "C:\temp\find.json"
parse --json-file "C:\temp\errors.json" --select "/Data/errors/rows/0/message"
parse --json-file "C:\temp\errors.json" --select "/Data/errors/rows/*/file" --format lines
document-slice --file "C:\repo\src\foo.cpp" --line 42 --context-before 8 --context-after 24 --out "C:\temp\slice.json"
file-outline --file "C:\repo\src\foo.cpp" --max-depth 2 --out "C:\temp\outline.json"
file-symbols --file "C:\repo\src\foo.cpp" --kind function --out "C:\temp\file-symbols.json"
search-symbols --query "propose_export_file_name_and_path" --kind function --out "C:\temp\symbols.json"
quick-info --file "C:\repo\src\foo.cpp" --line 42 --column 13 --out "C:\temp\quick-info.json"
document-slices --ranges-file "C:\temp\ranges.json" --out "C:\temp\slices.json"
smart-context --query "where is OnInit called" --out "C:\temp\smart-context.json"
find-references --file "C:\repo\src\foo.cpp" --line 42 --column 13 --out "C:\temp\refs.json"
call-hierarchy --file "C:\repo\src\foo.cpp" --line 42 --column 13 --out "C:\temp\hierarchy.json"
goto-definition --file "C:\repo\src\foo.cpp" --line 42 --column 13 --out "C:\temp\goto.json"

open-document --file "C:\repo\src\foo.cpp" --line 42 --column 1 --out "C:\temp\open-doc.json"
list-documents --out "C:\temp\documents.json"
list-tabs --out "C:\temp\tabs.json"
activate-document --query "foo.cpp" --out "C:\temp\activate.json"
close-document --query "foo.cpp" --out "C:\temp\close.json"
close-file --file "C:\repo\src\foo.cpp" --out "C:\temp\close-file.json"
close-others --out "C:\temp\close-others.json"
list-windows --query "Error List" --out "C:\temp\windows.json"
execute-command --command "View.SolutionExplorer" --out "C:\temp\exec.json"
apply-diff --patch-file "C:\temp\change.diff" --out "C:\temp\apply.json"

send --command set-breakpoint --args "--file \"C:\repo\src\foo.cpp\" --line 42 --condition \"count == 12\"" --out "C:\temp\bp.json"
send --command list-breakpoints --out "C:\temp\breakpoints.json"
send --command debug-start --args "--wait-for-break true --timeout-ms 120000" --out "C:\temp\debug-start.json"
send --command debug-continue --args "--wait-for-break true --timeout-ms 30000" --out "C:\temp\continue.json"

errors --wait-for-intellisense false --quick --out "C:\temp\errors.json"
build --configuration Debug --platform x64 --out "C:\temp\build.json"
build-errors --timeout-ms 600000 --out "C:\temp\build-errors.json"
```

### `errors` flags

- `--wait-for-intellisense true` (default) — waits for IntelliSense to finish loading before reading
- `--quick` — reads the current diagnostics snapshot immediately; skips the stability polling loop (use after a build has finished)

On large C++ solutions, `errors` and `warnings` may return diagnostics from Build Output when the Error List is empty or too slow to enumerate.

### `batch`

Execute multiple commands in a single bridge request:

```json
[
  { "command": "state" },
  { "command": "find-files", "args": "--query \"GUI_App.cpp\"" },
  { "command": "document-slice", "args": "--file \"C:\\repo\\src\\GUI_App.cpp\" --line 1384 --context-before 10 --context-after 30" }
]
```

```text
batch --file "C:\temp\batch.json" --out "C:\temp\batch-result.json"
```

Add `--stop-on-error` to halt on first failure. The result envelope contains a `results[]` array with per-step `success`, `summary`, and `data`.

### `IdeGoToDefinition`

Positions the cursor at `--file`/`--line`/`--column`, posts `Edit.GoToDefinition` through the VS shell dispatcher (same path as F12), and returns both the source location and the resolved definition location. Works on any language with a VS language service. `definitionFound` is `true` when the definition is at a different file or line.

### `IdeGetFileOutline`

Returns symbols in a file (functions, classes, structs, enums, namespaces) using VS's FileCodeModel. C# and VB support is complete; C++ support is partial and depends on VS having a code model for the file.

### `IdeGetFileSymbols`

Compatibility alias over `IdeGetFileOutline` for agents that naturally ask for file symbols. Supports the same `--file`, `--kind`, and `--max-depth` arguments.

### `IdeApplyUnifiedDiff`

Accepts `--patch-file` (path) or `--patch-text-base64` (inline). Existing-file edits are applied through the live Visual Studio editor buffer first, so the change is visible immediately and VS can re-evaluate syntax. By default the edited document stays unsaved; pass `--save-changed-files` if you want the bridge to save it.

### `IdeExecuteVsCommand`

Supports optional `--file`, `--document`, `--line`, `--column` args to position the editor before dispatching the native command. Useful for VS commands that act on the caret position.

## JSON Output

Every command writes a JSON envelope:

```json
{
  "SchemaVersion": 1,
  "Command": "Tools.IdeGetState",
  "RequestId": null,
  "Success": true,
  "StartedAtUtc": "2026-01-01T12:00:00.0000000Z",
  "FinishedAtUtc": "2026-01-01T12:00:00.0100000Z",
  "Summary": "IDE state captured.",
  "Warnings": [],
  "Error": null,
  "Data": {}
}
```

Failures use the same envelope shape with `Success: false` and a populated `Error` object containing `code` and `message`.

Bridge failures also include IDE context in `Data` when available:

- `state` - current solution, active document, caret, and bridge identity
- `openTabs` - currently open document tabs
- `errorList` - quick Error List snapshot
- `errorSymbolContext` - nearby symbols for files/lines mentioned in current error rows

For flat output, use the native CLI:

```bat
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe state --sln VsIdeBridge.sln --format keyvalue
```

## Named Pipe Server

The VSIX automatically starts a **persistent named pipe server** when Visual Studio loads the extension. Connecting directly over the pipe eliminates PowerShell process startup and COM mutex overhead on every call.

Use `scripts\start_bridge.ps1` for bootstrap only. The low-overhead runtime interface is the native C# CLI in `src\VsIdeBridgeCli\bin\<Configuration>\net8.0\vs-ide-bridge.exe`.

### Discovery

When the VSIX auto-loads, it writes:

```
%TEMP%\vs-ide-bridge\pipes\bridge-{pid}.json
```

```json
{
  "instanceId": "vs18-12345-20260303T040244Z",
  "pid": 12345,
  "startedAtUtc": "2026-03-03T04:02:44.0000000Z",
  "pipeName": "VsIdeBridge18_12345",
  "solutionPath": "C:\\path\\to\\Your.sln",
  "solutionName": "Your.sln"
}
```

The file is deleted when Visual Studio exits. A missing file means VS is not running or the extension failed to load.

### Native CLI (`vs-ide-bridge.exe`)

List live bridge instances:

```bat
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe instances --format summary
```

Resolve the current instance:

```bat
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe current
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe current --format keyvalue
```

Built-in help and command catalog:

```bat
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe help
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe prompts
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe help parse
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe help send
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe catalog --instance vs18-12345-20260303T040244Z
```

Single command:

```bat
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe ready --instance vs18-12345-20260303T040244Z
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe state --sln VsIdeBridge.sln --format summary
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe state --instance vs18-12345-20260303T040244Z --format summary
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe state --pid 12345 --format summary
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe find-files --instance vs18-12345-20260303T040244Z --query GUI_App.cpp
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe parse --json-file output\errors.json --select /Data/errors/rows/0/message
```

> **Git Bash note:** Git Bash converts arguments that start with `/` to Windows paths. If `--select` fails with a Windows-path error, either omit the leading slash (`Data/foo` instead of `/Data/foo`) or prefix the command with `MSYS_NO_PATHCONV=1`.

Batch request:

```bat
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe batch --file output\pipe-test-batch.json --format summary
```

Raw request object or array:

```bat
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe request --json "{ ""batch"": [ { ""command"": ""state"" } ] }"
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe request --json-file output\pipe-test-batch.json
```

Supported verbs:

- `help`
- `prompts` (alias: `recipes`)
- `current`
- `instances`
- `ensure`
- `ready`
- `state`
- `catalog` (alias: `commands`)
- `parse`
- `find-files`
- `find-text`
- `search-symbols` (alias: `symbols`)
- `goto-definition`
- `peek-definition` (alias: `peek`)
- `goto-implementation`
- `find-references`
- `call-hierarchy`
- `quick-info`
- `document-slice` (alias: `slice`)
- `document-slices` (alias: `slices`)
- `file-symbols`
- `file-outline`
- `open-document`
- `list-documents`
- `list-tabs`
- `activate-document`
- `close-document`
- `close-file`
- `close-others`
- `list-windows`
- `activate-window`
- `apply-diff` (alias: `patch`)
- `build`
- `warnings` (alias: `warning`)
- `errors`
- `build-errors`
- `send` (alias: `call`)
- `batch`
- `request`
- `mcp-server`

Common options:

- `--instance ID`
- `--pid PID`
- `--pipe NAME`
- `--sln HINT`
- `--timeout-ms 10000`
- `--out FILE`
- `--format json|summary|keyvalue`
- `--verbose`

Selection behavior:

- If exactly one live instance matches, the CLI uses it.
- If multiple live instances match, the CLI fails and tells you to run `instances`.
- Use `--instance` when you need an exact target across multiple open Visual Studio windows.

Recommended agent pattern:

- run `help` when you need to re-learn the interface
- run `prompts` for task-oriented examples you can copy directly
- use `ensure` first when you know the solution path
- otherwise use `current`
- use `catalog` to retrieve the live command list from Visual Studio
- use `parse` when you already have a JSON result and only need one field or list
- use `--instance` for all follow-up commands in the same task
- only fall back to `instances` when `current` says more than one IDE is live

### MCP server (`mcp-server`)

Run a stdio MCP server on Windows next to Visual Studio:

```bat
src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe mcp-server --instance <instanceId>
```

Exposed MCP tools use simple names:

- `state`
- `errors`
- `warnings`
- `list_tabs`
- `open_file`
- `search_symbols`
- `quick_info`
- `apply_diff`
- `git_status`
- `git_current_branch`
- `git_remote_list`
- `git_tag_list`
- `git_stash_list`
- `git_diff_unstaged`
- `git_diff_staged`
- `git_log`
- `git_show`
- `git_branch_list`
- `git_checkout`
- `git_create_branch`
- `git_add`
- `git_restore`
- `git_commit`
- `git_commit_amend`
- `git_reset`
- `git_fetch`
- `git_stash_push`
- `git_stash_pop`
- `git_pull`
- `git_push`
- `github_issue_search`
- `github_issue_close`

Exposed MCP resources:

- `bridge://current-solution`
- `bridge://active-document`
- `bridge://open-tabs`
- `bridge://error-list-snapshot`

Exposed MCP prompts:

- `help`
- `fix_current_errors`
- `open_solution_and_wait_ready`
- `git_review_before_commit`
- `git_sync_with_remote`
- `github_issue_triage`

The MCP layer is intentionally thin: it forwards to the bridge command surface and keeps edit approval/safety enforcement inside the existing Visual Studio bridge flow.

Implementation note: the server advertises MCP `tools` capability only. `resources/*` and `prompts/*` methods are still implemented, but not advertised, to avoid eager startup probes from MCP clients that can trigger Visual Studio automation calls before the IDE is fully ready.

Example MCP client registration (Claude Code and Codex CLI use the same stdio command):

```json
{
  "mcpServers": {
    "vs-ide-bridge": {
      "command": "C:\\path\\to\\vs-ide-bridge.exe",
      "args": [
        "mcp-server",
        "--instance",
        "<instanceId>"
      ]
    }
  }
}
```

### Pipe protocol

Each request is one JSON object terminated by `\n`:

```json
{ "id": "req-1", "command": "state", "args": "--out \"C:\\temp\\state.json\"" }
```

Batch requests can also send multiple logical commands in one envelope:

```json
{
  "id": "req-batch-1",
  "command": "batch",
  "stopOnError": false,
  "batch": [
    { "id": "ready", "command": "ready", "args": "--timeout-ms 120000" },
    { "id": "state", "command": "state", "args": "" }
  ]
}
```

Each response is one JSON envelope terminated by `\n`, using the same shape as the file-based command output.
For batched requests, the envelope `Data.results[]` array contains per-step `id`, `command`, `success`, `summary`, `warnings`, `data`, and `error`.

Use the discovery file to find the current pipe name, then send newline-delimited UTF-8 JSON over that named pipe. `args` uses the same command-line argument format as the bridge commands already use inside Visual Studio.

## Scripts

| Script | Purpose |
|--------|---------|
| `scripts\build.bat` | Build the solution |
| `scripts\start_bridge.ps1` | Thin PowerShell wrapper over `vs-ide-bridge ensure` |

## Repo Layout

```
src/VsIdeBridge/          VSIX package, commands, services, infrastructure
scripts/                  Build and startup entry points only
output/                   Local smoke-test artifacts (git-ignored)
```

## Notes

- The Tools menu exposes **IDE Bridge > Help**, **Allow Bridge Edits**, and **Go To Edited Parts**. `Help` opens the repo README when the current solution resolves to this repo and points you to `help` and `catalog` for the full command catalog. All other commands remain `CommandWellOnly` (available in the Command Window and via DTE, not in menus).
- Search scans files on disk; unsaved in-memory editor content is not included in `find-text` results.
- Symbol commands rely on VS language services, not bridge-side parsing.
- `execute-command` is the escape hatch for native VS commands that have no first-class bridge equivalent.
- Simple pipe names are the preferred public contract. The legacy `Tools.Ide*` names remain supported for compatibility.
