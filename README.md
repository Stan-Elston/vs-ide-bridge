# VS IDE Bridge

Local Visual Studio 2026 extension for scriptable IDE control through stable `Tools.*` commands.

## Disclaimer

This project is experimental. Use it at your own risk.

## Purpose

This repo exposes Visual Studio automation commands that are easy to call from:

- the Visual Studio Command Window
- DTE automation
- wrapper scripts and agent tooling

The first release includes:

- IDE state snapshots
- solution readiness waiting
- file and text search with JSON output
- document/tab/window activation
- open-tab listing and cleanup
- symbol actions through native Visual Studio commands
- smart context gathering for large codebases
- unified diff application with changed-file reopen
- breakpoint management
- debugger control and state capture
- build and Error List capture

It does not edit source text.

## Status

Current state:

- VSIX builds and installs successfully on Visual Studio 2026 / 18
- wrapper automation works against both this repo and SuperSlicer
- current feature update adds:
  - `Tools.IdeListOpenTabs`
  - `Tools.IdeCloseFile`
  - `Tools.IdeCloseAllExceptCurrent`
  - `Tools.IdeGetSmartContextForQuery`
  - `Tools.IdeApplyUnifiedDiff`
  - breakpoint reveal-on-set
  - debugger wait options for start/continue/step commands
- validated commands:
  - `Tools.IdeWaitForReady`
  - `Tools.IdeGetState`
  - `Tools.IdeFindFiles`
  - `Tools.IdeFindText`
  - `Tools.IdeListOpenTabs`
  - `Tools.IdeCloseFile`
  - `Tools.IdeCloseAllExceptCurrent`
  - `Tools.IdeGetSmartContextForQuery`
  - `Tools.IdeGetErrorList`

SuperSlicer validation used:

- `C:\Users\elsto\source\repos\Stan-Elston\SuperSlicer\build\Slic3r.sln`

## Requirements

- Visual Studio 2026 / 18 Community, Pro, or Enterprise
- Windows

## Repo Layout

- `src/VsIdeBridge`
  - VSIX package, command registrations, services, and infrastructure
- `scripts/build_vsix.bat`
  - local build entry point
- `scripts/install_vsix.bat`
  - local VSIX installer entry point
- `scripts/invoke_vs_ide_command.ps1`
  - attach/open/invoke wrapper for external automation
- `scripts/invoke_active_vs_command.ps1`
  - attach to an already-open Visual Studio 18 instance and execute one bridge command
- `scripts/list_vs_commands.ps1`
  - enumerate native Visual Studio command names from a live DTE instance
- `scripts/read_bridge_result.bat`
  - native JSON-to-`key=value` reader for shell clients and LLM tooling
- `scripts/read_bridge_result.ps1`
  - PowerShell entry point for the same native result reader
- `scripts/smoke_test.ps1`
  - end-to-end validation flow
- `scripts/vs_dte_probe.ps1`
  - inspect live Visual Studio 18 DTE instances
- `output/`
  - local smoke-test and validation artifacts, ignored by git

## Build

```bat
scripts\build_vsix.bat
```

## Install

```bat
scripts\install_vsix.bat
```

Installing the VSIX updates the extension in Visual Studio and may close `devenv.exe` if it is running.

## Quick Start

1. Build the VSIX.
2. Install it into Visual Studio.
3. Open a solution in Visual Studio, or let the wrapper open one for you.
4. Invoke `Tools.Ide*` commands from the Command Window or the PowerShell wrapper.

Minimal external example:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\invoke_vs_ide_command.ps1 `
  -SolutionPath C:\Users\elsto\source\repos\vs-ide-bridge\VsIdeBridge.sln `
  -CommandName Tools.IdeGetState `
  -OutputPath C:\temp\ide-state.json
```

Minimal Command Window example:

```text
Tools.IdeGetState --out "C:\temp\ide-state.json"
```

## Command Surface

Core:

- `Tools.IdeGetState`
- `Tools.IdeWaitForReady`

Search and navigation:

- `Tools.IdeFindFiles`
- `Tools.IdeFindText`
- `Tools.IdeOpenDocument`
- `Tools.IdeListDocuments`
- `Tools.IdeListOpenTabs`
- `Tools.IdeActivateDocument`
- `Tools.IdeCloseDocument`
- `Tools.IdeCloseFile`
- `Tools.IdeCloseAllExceptCurrent`
- `Tools.IdeActivateWindow`
- `Tools.IdeListWindows`
- `Tools.IdeExecuteVsCommand`
- `Tools.IdeFindAllReferences`
- `Tools.IdeShowCallHierarchy`
- `Tools.IdeGetDocumentSlice`
- `Tools.IdeGetSmartContextForQuery`
- `Tools.IdeApplyUnifiedDiff`

Breakpoints:

- `Tools.IdeSetBreakpoint`
- `Tools.IdeListBreakpoints`
- `Tools.IdeRemoveBreakpoint`
- `Tools.IdeClearAllBreakpoints`

Debugger:

- `Tools.IdeDebugGetState`
- `Tools.IdeDebugStart`
- `Tools.IdeDebugStop`
- `Tools.IdeDebugBreak`
- `Tools.IdeDebugContinue`
- `Tools.IdeDebugStepOver`
- `Tools.IdeDebugStepInto`
- `Tools.IdeDebugStepOut`

Build and diagnostics:

- `Tools.IdeBuildSolution`
- `Tools.IdeGetErrorList`
- `Tools.IdeBuildAndCaptureErrors`

All operational commands accept an argument string and write a JSON result file.

Example:

```text
Tools.IdeGetState --out "C:\temp\ide-state.json"
```

Every command writes:

- a JSON envelope to the requested `--out` path
- a one-line summary to the `IDE Bridge` Output pane
- a status-bar summary

If `--out` is omitted, the extension writes to `%TEMP%\vs-ide-bridge\*.json`.

## Argument Contract

- `--out "C:\path\result.json"`: preferred output path
- `--request-id "abc123"`: optional correlation id
- `--timeout-ms 120000`: optional on wait/build commands
- booleans use `true` or `false`
- enum values use lowercase kebab-case

Examples:

```text
Tools.IdeWaitForReady --out "C:\temp\ready.json" --timeout-ms 120000
Tools.IdeFindFiles --query "VsIdeBridge.csproj" --out "C:\temp\files.json"
Tools.IdeFindText --query "Tools.IdeGetState" --scope solution --out "C:\temp\find.json"
Tools.IdeOpenDocument --file "C:\repo\src\foo.cpp" --line 42 --column 1 --out "C:\temp\open.json"
Tools.IdeListDocuments --out "C:\temp\documents.json"
Tools.IdeListOpenTabs --out "C:\temp\tabs.json"
Tools.IdeActivateDocument --query "foo.cpp" --out "C:\temp\activate-document.json"
Tools.IdeCloseDocument --query "foo.cpp" --out "C:\temp\close-document.json"
Tools.IdeCloseFile --file "C:\repo\src\foo.cpp" --out "C:\temp\close-file.json"
Tools.IdeCloseAllExceptCurrent --out "C:\temp\close-other-tabs.json"
Tools.IdeListWindows --query "References" --out "C:\temp\windows.json"
Tools.IdeExecuteVsCommand --command "View.SolutionExplorer" --out "C:\temp\execute-command.json"
Tools.IdeFindAllReferences --file "C:\repo\src\foo.cpp" --line 42 --column 13 --out "C:\temp\references.json"
Tools.IdeShowCallHierarchy --file "C:\repo\src\foo.cpp" --line 42 --column 13 --out "C:\temp\call-hierarchy.json"
Tools.IdeGetDocumentSlice --file "C:\repo\src\foo.cpp" --line 42 --context-before 8 --context-after 24 --out "C:\temp\slice.json"
Tools.IdeGetSmartContextForQuery --query "where is GUI_App::OnInit used" --out "C:\temp\smart-context.json"
Tools.IdeApplyUnifiedDiff --patch-file "C:\temp\change.diff" --out "C:\temp\apply-diff.json"
Tools.IdeSetBreakpoint --file "C:\repo\src\foo.cpp" --line 42 --condition "count == 12" --reveal true --out "C:\temp\bp.json"
Tools.IdeDebugStart --wait-for-break true --timeout-ms 120000 --out "C:\temp\debug-start.json"
Tools.IdeBuildAndCaptureErrors --out "C:\temp\build-errors.json" --timeout-ms 600000
```

`Tools.IdeFindAllReferences` and `Tools.IdeShowCallHierarchy` use the symbol at the active caret by default. Use `--file`, `--document`, `--line`, and `--column` when you want the bridge to position the editor first.

`Tools.IdeApplyUnifiedDiff` accepts either `--patch-file` or `--patch-text-base64`.

`Tools.IdeSetBreakpoint` reveals the target line in the editor by default so the user can immediately see where the breakpoint landed.

Search and navigation results use normalized absolute file paths so an external tool can safely open or edit the same file in the background.

Bridge automation is serialized through a shared mutex so overlapping wrapper calls do not open multiple Visual Studio instances against the same solution at the same time.

## Large-File Workflow

For large C++ files, the intended bridge flow is:

1. Use `Tools.IdeFindText` to get the exact match location from Visual Studio.
2. Use `Tools.IdeGetDocumentSlice` to fetch only the lines around that hit.

Example:

```text
Tools.IdeFindText --query "choose_app_dir" --scope solution --out "C:\temp\find.json"
Tools.IdeGetDocumentSlice --file "C:\repo\src\GUI_App.cpp" --line 1045 --context-before 6 --context-after 20 --out "C:\temp\slice.json"
```

This keeps the bridge useful on large codebases without repeatedly rescanning full source files outside Visual Studio.

For symbol-heavy work, the next step after a slice is usually a native Visual Studio command:

```text
Tools.IdeFindAllReferences --file "C:\repo\src\GUI_App.cpp" --line 1045 --column 12 --out "C:\temp\references.json"
Tools.IdeShowCallHierarchy --file "C:\repo\src\GUI_App.cpp" --line 1045 --column 12 --out "C:\temp\call-hierarchy.json"
```

For agent-driven edits, the bridge can also apply a unified diff directly and reopen the changed files:

```text
Tools.IdeApplyUnifiedDiff --patch-file "C:\temp\change.diff" --base-directory "C:\repo" --out "C:\temp\apply-diff.json"
```

## JSON Output

Each command writes a JSON envelope. Current implementation uses these top-level keys:

```json
{
  "SchemaVersion": 1,
  "Command": "Tools.IdeGetState",
  "RequestId": null,
  "Success": true,
  "StartedAtUtc": "2026-03-01T18:50:06.0479106Z",
  "FinishedAtUtc": "2026-03-01T18:50:06.0589103Z",
  "Summary": "IDE state captured.",
  "Warnings": [],
  "Error": null,
  "Data": {}
}
```

Failure results keep the same envelope and populate `Error`.

If a shell client or LLM does not want to parse JSON directly, use the native result reader:

```bat
scripts\read_bridge_result.bat C:\temp\ide-state.json
```

Example output:

```text
command=Tools.IdeGetState
success=true
summary=IDE state captured.
```

## Scripts

- `scripts\invoke_vs_ide_command.ps1`
  Launch or attach to Visual Studio and invoke an IDE Bridge command.
- `scripts\invoke_active_vs_command.ps1`
  Attach to an already-open Visual Studio 18 instance and execute one command without solution-launch logic.
- `scripts\list_vs_commands.ps1`
  Enumerate Visual Studio command names so bridge wrappers can target native IDE actions accurately.
- `scripts\smoke_test.ps1`
  Run a small end-to-end validation against the extension.
- `scripts\vs_dte_probe.ps1`
  Inspect live Visual Studio 18 DTE instances.

## Wrapper Usage

Use the PowerShell wrapper when you want to drive the extension from outside Visual Studio:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\invoke_vs_ide_command.ps1 `
  -SolutionPath C:\Users\elsto\source\repos\vs-ide-bridge\VsIdeBridge.sln `
  -CommandName Tools.IdeGetState `
  -OutputPath C:\temp\ide-state.json
```

Key wrapper behavior:

- reuses an existing Visual Studio 18 instance when the requested solution is already open
- can open the requested solution in a blank VS instance
- leaves Visual Studio open by default
- only closes Visual Studio when `-CloseVisualStudio` is passed
- waits for the output JSON file to be written before returning
- resolves the internal `Tools.Tools.Ide*` registration quirk so callers can use `Tools.Ide*`
- supports `-ResultFormat summary|keyvalue|json` so callers can choose human summary text, shell-friendly `key=value`, or raw JSON

Run the smoke test with:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\smoke_test.ps1
```

Run it against SuperSlicer with:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\invoke_vs_ide_command.ps1 `
  -SolutionPath C:\Users\elsto\source\repos\Stan-Elston\SuperSlicer\build\Slic3r.sln `
  -CommandName Tools.IdeGetErrorList `
  -CommandArgs "--wait-for-intellisense true --timeout-ms 240000" `
  -OutputPath C:\Users\elsto\source\repos\vs-ide-bridge\output\superslicer-errors.json
```

## Validation

Validated locally:

- build and install of the VSIX
- wrapper-driven smoke test on this repo
- wrapper-driven run on SuperSlicer:
  - readiness wait
  - state capture
  - file search for `GUI_App.cpp`
  - tab open/activate/close flow across `GUI_App.cpp` and `GUI_App.hpp`
  - smart-context query for `where is GUI_App::OnInit used`
  - native `View.CallHierarchy` against `choose_app_dir`
  - native `Edit.FindAllReferences` against `choose_app_dir`
  - Error List export

Example generated artifacts:

- `output\smoke-test.txt`
- `output\superslicer-ready.json`
- `output\superslicer-state.json`
- `output\superslicer-find-files.json`
- `output\superslicer-list-documents.json`
- `output\superslicer-list-windows.json`
- `output\superslicer-call-hierarchy.json`
- `output\superslicer-find-all-references.json`
- `output\superslicer-errors.json`

## Notes

- Search results are written to JSON and surfaced in the `IDE Bridge` Output pane.
- Build and Error List commands reuse the same Error List extraction logic that worked in `vs-errorlist-export`.
- The extension is command-first. The only visible menu items are `Help` and `Smoke Test`.
- Search currently scans files on disk; unsaved in-memory editor changes are not yet part of search results.
- Symbol commands intentionally rely on Visual Studio's own language services rather than bridge-side parsing.
- `Tools.IdeExecuteVsCommand` remains available for host-specific VS commands that have not been promoted to first-class bridge commands.
- `Tools.Ide*` is the public contract even though Visual Studio internally registers these commands as `Tools.Tools.Ide*`.
