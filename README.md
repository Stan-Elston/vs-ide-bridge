# VS IDE Bridge

Visual Studio extension that exposes scriptable `Tools.Ide*` commands for external automation — wrapper scripts, agent tooling, and shell pipelines.

## Disclaimer

This project is experimental. Use it at your own risk.

## Purpose

Lets you drive a running Visual Studio instance from outside the IDE: search code, navigate to symbols, slice documents, apply diffs, control the debugger, and capture build output — all without touching the keyboard.

Commands are invoked through stable `Tools.*` names, the same surface used by the VS Command Window and DTE automation. Results are written as JSON to a caller-specified output file.

## Requirements

- Windows
- Visual Studio 2026 / 18

The extension targets VS 18, but the helper scripts in `scripts\` currently probe the `Community` install path first and then fall back to the VS 2022 Community path. If you use Professional or Enterprise, adjust the script paths before using the build/install wrappers.

## Script Requirements

### PowerShell scripts

The PowerShell entry points in `scripts\` use built-in Windows PowerShell features only:

- PowerShell with COM interop support (`Add-Type`, ROT/DTE automation)
- Standard Windows cmdlets such as `Get-Process`, `Stop-Process`, `Start-Process`, `Get-CimInstance`, and `ConvertFrom-Json`

No extra PowerShell modules are required.

### Python pipe client

`scripts\vs_bridge_pipe.py` requires:

- Python 3
- `pywin32`

Install with:

```bash
python -m pip install -r scripts/requirements-vs_bridge_pipe.txt
```

Or:

```bash
python -m pip install pywin32
```

### Native helper used by result readers

`scripts\read_bridge_result.ps1` and `scripts\read_bridge_result.bat` depend on `IdeBridgeJsonProbe.exe`, which is built from the `src\IdeBridgeJsonProbe` project in this repo. Build the solution first:

```bat
scripts\build_vsix.bat
```

## Build

```bat
scripts\build_vsix.bat
```

Builds the Debug configuration by default. Pass a configuration name as the first argument:

```bat
scripts\build_vsix.bat Release
```

## Install

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\install_vsix.ps1
```

Or via the batch shim (calls the same PS1):

```bat
scripts\install_vsix.bat
```

The installer:
1. Builds the VSIX (skip with `-NoBuild`).
2. Closes any running Visual Studio instances gracefully, then force-kills lingering `MSBuild.exe` workers that would otherwise block the installer.
3. Runs `VSIXInstaller /quiet` without `/shutdownprocesses` (which hangs when VS has unsaved files).
4. Relaunches Visual Studio with the solutions that were open before the install.

Pass `-NoRelaunch` to skip step 4.

## Quick Start

1. Build and install the VSIX.
2. Open a solution in Visual Studio, or let the wrapper open one.
3. Invoke `Tools.Ide*` commands from the Command Window or the PowerShell wrapper.

**Command Window:**

```text
Tools.IdeGetState --out "C:\temp\ide-state.json"
```

**PowerShell wrapper:**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\invoke_vs_ide_command.ps1 `
  -SolutionPath "C:\path\to\your.sln" `
  -CommandName  Tools.IdeGetState `
  -OutputPath   "C:\temp\ide-state.json"
```

## Command Surface

### Core

| Command | Description |
|---------|-------------|
| `Tools.IdeHelp` | List all registered commands |
| `Tools.IdeSmokeTest` | Capture a minimal state snapshot for smoke validation |
| `Tools.IdeGetState` | Snapshot of IDE state (solution path, active document, etc.) |
| `Tools.IdeWaitForReady` | Block until IntelliSense is available |
| `Tools.IdeOpenSolution` | Open a solution file |
| `Tools.IdeBatchCommands` | Execute multiple commands in one wrapper round-trip |

### Search and Navigation

| Command | Description |
|---------|-------------|
| `Tools.IdeFindText` | Text search across solution or project |
| `Tools.IdeFindFiles` | Find files by name |
| `Tools.IdeGetDocumentSlice` | Fetch lines around a location |
| `Tools.IdeGetFileOutline` | Symbol list for a file (functions, classes, etc.) |
| `Tools.IdeSearchSymbols` | Search symbols by name across the current scope |
| `Tools.IdeGetQuickInfo` | Resolve symbol/definition info at a location |
| `Tools.IdeGetDocumentSlices` | Fetch multiple document ranges from a JSON ranges file |
| `Tools.IdeGetSmartContextForQuery` | Multi-context gather for agent queries |
| `Tools.IdeFindAllReferences` | Semantic find-all-references |
| `Tools.IdeShowCallHierarchy` | Callers and callees |
| `Tools.IdeGoToDefinition` | Navigate to definition (F12) |
| `Tools.IdeOpenDocument` | Open a file at a line/column |
| `Tools.IdeListDocuments` | List open documents |
| `Tools.IdeListOpenTabs` | List open editor tabs |
| `Tools.IdeActivateDocument` | Activate a document by name |
| `Tools.IdeCloseDocument` | Close a document by name |
| `Tools.IdeCloseFile` | Close a document by full path |
| `Tools.IdeCloseAllExceptCurrent` | Close all tabs except the active one |
| `Tools.IdeActivateWindow` | Activate a tool window |
| `Tools.IdeListWindows` | List tool windows |
| `Tools.IdeExecuteVsCommand` | Execute any native VS command by name |
| `Tools.IdeApplyUnifiedDiff` | Apply a unified diff and reopen changed files |

### Breakpoints

| Command | Description |
|---------|-------------|
| `Tools.IdeSetBreakpoint` | Set a breakpoint at file/line |
| `Tools.IdeListBreakpoints` | List all breakpoints |
| `Tools.IdeRemoveBreakpoint` | Remove a breakpoint |
| `Tools.IdeClearAllBreakpoints` | Remove all breakpoints |
| `Tools.IdeEnableBreakpoint` | Enable a breakpoint at file/line |
| `Tools.IdeDisableBreakpoint` | Disable a breakpoint at file/line |
| `Tools.IdeEnableAllBreakpoints` | Enable every breakpoint |
| `Tools.IdeDisableAllBreakpoints` | Disable every breakpoint |

### Debugger

| Command | Description |
|---------|-------------|
| `Tools.IdeDebugGetState` | Current debugger state |
| `Tools.IdeDebugStart` | Start debugging |
| `Tools.IdeDebugStop` | Stop debugging |
| `Tools.IdeDebugBreak` | Break into the debugger |
| `Tools.IdeDebugContinue` | Continue execution |
| `Tools.IdeDebugStepOver` | Step over |
| `Tools.IdeDebugStepInto` | Step into |
| `Tools.IdeDebugStepOut` | Step out |

### Build and Diagnostics

| Command | Description |
|---------|-------------|
| `Tools.IdeBuildSolution` | Build the solution |
| `Tools.IdeGetErrorList` | Capture the Error List as JSON |
| `Tools.IdeBuildAndCaptureErrors` | Build then capture the Error List in one call |

## Argument Contract

- `--out "C:\path\result.json"` — output path (falls back to `%TEMP%\vs-ide-bridge\<command>.json`)
- `--request-id "abc123"` — optional correlation id echoed in the envelope
- `--timeout-ms 120000` — timeout on wait/build commands
- boolean flags: `--flag` (bare, implies `true`) or `--flag true` / `--flag false`
- enum values: lowercase kebab-case

## Command Examples

```text
Tools.IdeWaitForReady --timeout-ms 120000 --out "C:\temp\ready.json"
Tools.IdeGetState --out "C:\temp\state.json"
Tools.IdeOpenSolution --solution "C:\path\to\your.sln" --out "C:\temp\open.json"

Tools.IdeFindFiles --query "GUI_App.cpp" --out "C:\temp\files.json"
Tools.IdeFindText --query "OnInit" --scope solution --out "C:\temp\find.json"
Tools.IdeGetDocumentSlice --file "C:\repo\src\foo.cpp" --line 42 --context-before 8 --context-after 24 --out "C:\temp\slice.json"
Tools.IdeGetFileOutline --file "C:\repo\src\foo.cpp" --max-depth 2 --out "C:\temp\outline.json"
Tools.IdeGetSmartContextForQuery --query "where is OnInit called" --out "C:\temp\smart-context.json"
Tools.IdeFindAllReferences --file "C:\repo\src\foo.cpp" --line 42 --column 13 --out "C:\temp\refs.json"
Tools.IdeShowCallHierarchy --file "C:\repo\src\foo.cpp" --line 42 --column 13 --out "C:\temp\hierarchy.json"
Tools.IdeGoToDefinition --file "C:\repo\src\foo.cpp" --line 42 --column 13 --out "C:\temp\goto.json"

Tools.IdeOpenDocument --file "C:\repo\src\foo.cpp" --line 42 --column 1 --out "C:\temp\open-doc.json"
Tools.IdeListDocuments --out "C:\temp\documents.json"
Tools.IdeListOpenTabs --out "C:\temp\tabs.json"
Tools.IdeActivateDocument --query "foo.cpp" --out "C:\temp\activate.json"
Tools.IdeCloseDocument --query "foo.cpp" --out "C:\temp\close.json"
Tools.IdeCloseFile --file "C:\repo\src\foo.cpp" --out "C:\temp\close-file.json"
Tools.IdeCloseAllExceptCurrent --out "C:\temp\close-others.json"
Tools.IdeListWindows --query "Error List" --out "C:\temp\windows.json"
Tools.IdeExecuteVsCommand --command "View.SolutionExplorer" --out "C:\temp\exec.json"
Tools.IdeApplyUnifiedDiff --patch-file "C:\temp\change.diff" --out "C:\temp\apply.json"

Tools.IdeSetBreakpoint --file "C:\repo\src\foo.cpp" --line 42 --condition "count == 12" --out "C:\temp\bp.json"
Tools.IdeListBreakpoints --out "C:\temp\breakpoints.json"
Tools.IdeDebugStart --wait-for-break true --timeout-ms 120000 --out "C:\temp\debug-start.json"
Tools.IdeDebugContinue --wait-for-break true --timeout-ms 30000 --out "C:\temp\continue.json"

Tools.IdeGetErrorList --wait-for-intellisense false --quick --out "C:\temp\errors.json"
Tools.IdeBuildSolution --configuration Debug --platform x64 --out "C:\temp\build.json"
Tools.IdeBuildAndCaptureErrors --timeout-ms 600000 --out "C:\temp\build-errors.json"
```

### `IdeGetErrorList` flags

- `--wait-for-intellisense true` (default) — waits for IntelliSense to finish loading before reading
- `--quick` — reads the Error List once immediately; skips the stability polling loop (use after a build has finished)

### `IdeBatchCommands`

Execute multiple commands in a single wrapper invocation — one PS process, one mutex acquire, one COM setup:

```json
[
  { "command": "Tools.IdeGetState" },
  { "command": "Tools.IdeFindFiles", "args": "--query \"GUI_App.cpp\"" },
  { "command": "Tools.IdeGetDocumentSlice", "args": "--file \"C:\\repo\\src\\GUI_App.cpp\" --line 1384 --context-before 10 --context-after 30" }
]
```

```text
Tools.IdeBatchCommands --batch-file "C:\temp\batch.json" --out "C:\temp\batch-result.json"
```

Add `--stop-on-error` to halt on first failure. The result envelope contains a `results[]` array with per-step `success`, `summary`, and `data`.

### `IdeGoToDefinition`

Positions the cursor at `--file`/`--line`/`--column`, posts `Edit.GoToDefinition` through the VS shell dispatcher (same path as F12), and returns both the source location and the resolved definition location. Works on any language with a VS language service. `definitionFound` is `true` when the definition is at a different file or line.

### `IdeGetFileOutline`

Returns symbols in a file (functions, classes, structs, enums, namespaces) using VS's FileCodeModel. C# and VB support is complete; C++ support is partial and depends on VS having a code model for the file.

### `IdeApplyUnifiedDiff`

Accepts `--patch-file` (path) or `--patch-text-base64` (inline). After patching, reopens all changed files in VS so the editor immediately shows the new content.

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

For shell clients that prefer flat output:

```bat
scripts\read_bridge_result.bat C:\temp\ide-state.json
```

```text
command=Tools.IdeGetState
success=true
summary=IDE state captured.
```

## Named Pipe Server

The VSIX automatically starts a **persistent named pipe server** when Visual Studio loads the extension. Connecting directly over the pipe eliminates PowerShell process startup and COM mutex overhead on every call.

**Benchmark (measured):**

| Method | 1 command | 3 commands (batch) |
|--------|-----------|-------------------|
| PS wrapper | ~11 s | ~11 s |
| Pipe client (Python startup) | ~3 s | ~3.2 s |
| Pipe round-trip only | < 1 ms | < 3 ms |

### Discovery

When the VSIX starts it writes:

```
%TEMP%\vs-ide-bridge\pipes\bridge-{pid}.json
```

```json
{ "pid": 12345, "pipeName": "VsIdeBridge18_12345" }
```

The file is deleted when Visual Studio exits. A missing file means VS is not running or the extension failed to load.

### Python client (`scripts\vs_bridge_pipe.py`)

Requires **Python 3** and **pywin32**:

```bash
python -m pip install -r scripts/requirements-vs_bridge_pipe.txt
```

Single command:

```bash
conda run -n superslicer python scripts/vs_bridge_pipe.py \
  --cmd Tools.IdeGetState --format summary
```

With arguments:

```bash
conda run -n superslicer python scripts/vs_bridge_pipe.py \
  --cmd Tools.IdeFindFiles --args '--query "GUI_App.cpp"' --format json
```

Batch (all commands share one Python process):

```bash
conda run -n superslicer python scripts/vs_bridge_pipe.py \
  --batch output/test-batch.json --format summary
```

Batch file format (same as `IdeBatchCommands`):

```json
[
  { "command": "Tools.IdeGetState" },
  { "command": "Tools.IdeFindFiles", "args": "--query \"GUI_App.cpp\"" }
]
```

`--format` choices: `json` (default) · `summary` · `keyvalue`
`--out FILE` writes the full JSON response(s) to a file.
`--sln HINT` filters discovery when multiple VS instances are running.

### Pipe protocol

Each request is one JSON object terminated by `\n`:

```json
{ "id": "req-1", "command": "Tools.IdeGetState", "args": "--out \"C:\\temp\\state.json\"" }
```

Each response is one JSON envelope terminated by `\n`, using the same shape as the file-based command output.

Use the discovery file to find the current pipe name, then send newline-delimited UTF-8 JSON over that named pipe. `args` uses the same command-line argument format as the DTE and PowerShell wrappers.

## Scripts

| Script | Purpose |
|--------|---------|
| `scripts\build_vsix.bat` | Build the VSIX |
| `scripts\install_vsix.ps1` | Build + install (handles VS shutdown/relaunch) |
| `scripts\install_vsix.bat` | Shim that delegates to `install_vsix.ps1` |
| `scripts\invoke_vs_ide_command.ps1` | Launch/attach VS and invoke a bridge command |
| `scripts\invoke_active_vs_command.ps1` | Attach to an already-open VS instance and run one command |
| `scripts\list_vs_commands.ps1` | Enumerate native VS command names from a live DTE instance |
| `scripts\smoke_test.ps1` | End-to-end validation flow |
| `scripts\read_bridge_result.bat` | Read a JSON envelope as `key=value` pairs |
| `scripts\read_bridge_result.ps1` | PowerShell entry point for the same reader |
| `scripts\vs_bridge_pipe.py` | Direct named-pipe client for low-overhead command execution |
| `scripts\vs_dte_probe.ps1` | Inspect live VS 18 DTE instances |

## Wrapper Reference

`invoke_vs_ide_command.ps1` parameters:

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-SolutionPath` | — | Open this solution before running the command |
| `-CommandName` | (required) | `Tools.Ide*` command name |
| `-CommandArgs` | `""` | Arguments string passed to the command |
| `-OutputPath` | auto | JSON output file path |
| `-StartupTimeoutSeconds` | 60 | Timeout waiting for VS to start |
| `-CommandTimeoutSeconds` | 120 | Timeout waiting for the output file |
| `-ReuseVisualStudio` | `$true` | Attach to an existing VS instance |
| `-ResultFormat` | `summary` | `summary`, `keyvalue`, or `json` |
| `-CloseVisualStudio` | `$false` | Close VS after the command completes |

Key behaviors:
- Reuses an existing VS 18 instance when the solution is already open.
- Starts a new VS instance and opens the solution if needed.
- Serializes concurrent wrapper calls through a named mutex.
- Resolves the internal `Tools.Tools.Ide*` registration quirk transparently.

## Repo Layout

```
src/VsIdeBridge/          VSIX package, commands, services, infrastructure
scripts/                  Build, install, and automation helper scripts
output/                   Local smoke-test artifacts (git-ignored)
```

## Notes

- The only visible menu items are **IDE Bridge > Help** and **IDE Bridge > Smoke Test**. All other commands are `CommandWellOnly` (available in the Command Window and via DTE, not in menus).
- Search scans files on disk; unsaved in-memory editor content is not included in `IdeFindText` results.
- Symbol commands rely on VS language services, not bridge-side parsing.
- `Tools.IdeExecuteVsCommand` is the escape hatch for native VS commands that have no first-class bridge equivalent.
- `Tools.Ide*` is the public contract. VS internally registers them as `Tools.Tools.Ide*`; the wrapper resolves this automatically.
