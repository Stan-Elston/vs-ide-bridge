# VS IDE Bridge

Local Visual Studio 2026 extension for scriptable IDE control through stable `Tools.*` commands.

## Purpose

This repo exposes Visual Studio automation commands that are easy to call from:

- the Visual Studio Command Window
- DTE automation
- wrapper scripts and agent tooling

The first release includes:

- IDE state snapshots
- solution readiness waiting
- file and text search with JSON output
- document/window activation
- breakpoint management
- debugger control and state capture
- build and Error List capture

It does not edit source text.

## Status

Current state:

- VSIX builds and installs successfully on Visual Studio 2026 / 18
- wrapper automation works against both this repo and SuperSlicer
- validated commands:
  - `Tools.IdeWaitForReady`
  - `Tools.IdeGetState`
  - `Tools.IdeFindFiles`
  - `Tools.IdeFindText`
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
- `Tools.IdeActivateWindow`

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
Tools.IdeSetBreakpoint --file "C:\repo\src\foo.cpp" --line 42 --out "C:\temp\bp.json"
Tools.IdeBuildAndCaptureErrors --out "C:\temp\build-errors.json" --timeout-ms 600000
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

## Scripts

- `scripts\invoke_vs_ide_command.ps1`
  Launch or attach to Visual Studio and invoke an IDE Bridge command.
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
  - Error List export

Example generated artifacts:

- `output\smoke-test.txt`
- `output\superslicer-ready.json`
- `output\superslicer-state.json`
- `output\superslicer-find-files.json`
- `output\superslicer-errors.json`

## Notes

- Search results are written to JSON and surfaced in the `IDE Bridge` Output pane.
- Build and Error List commands reuse the same Error List extraction logic that worked in `vs-errorlist-export`.
- The extension is command-first. The only visible menu items are `Help` and `Smoke Test`.
- Search currently scans files on disk; unsaved in-memory editor changes are not yet part of search results.
- `Tools.Ide*` is the public contract even though Visual Studio internally registers these commands as `Tools.Tools.Ide*`.
