#!/usr/bin/env python3
"""
vs_bridge_pipe.py -- Direct named pipe client for VS IDE Bridge.

Eliminates per-call PowerShell overhead (~1500 ms -> ~50 ms) by connecting
directly to the persistent pipe server running inside the VSIX.

Discovery: %TEMP%\\vs-ide-bridge\\pipes\\bridge-{pid}.json
Protocol:  newline-delimited JSON (one request / one response per line)

Usage:
  python vs_bridge_pipe.py --cmd Tools.IdeGetState [--args "..."] [--out file.json] [--format json|summary|keyvalue]
  python vs_bridge_pipe.py --batch batch.json [--out file.json] [--format json|summary|keyvalue]

Batch file format (same as IdeBatchCommands):
  [{"command": "Tools.IdeGetState", "args": ""},
   {"command": "Tools.IdeFindFiles", "args": "--query Foo.cs"}]
"""

import argparse
import json
import os
import sys
import time
import uuid
from pathlib import Path


def _discovery_dir() -> Path:
    temp = Path(os.environ.get("TEMP", os.environ.get("TMP", "/tmp")))
    return temp / "vs-ide-bridge" / "pipes"


def find_pipe(sln_hint: str | None = None) -> str:
    """Return the pipe name from the discovery file, or raise RuntimeError."""
    disc_dir = _discovery_dir()
    if not disc_dir.exists():
        raise RuntimeError(
            f"Discovery directory not found: {disc_dir}\n"
            "Is the VS IDE Bridge VSIX installed and Visual Studio running?"
        )

    candidates = list(disc_dir.glob("bridge-*.json"))
    if not candidates:
        raise RuntimeError(
            f"No bridge discovery files in {disc_dir}\n"
            "Is Visual Studio running with the VS IDE Bridge extension loaded?"
        )

    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)

    for path in candidates:
        try:
            info = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue

        pipe_name = info.get("pipeName", "")
        if not pipe_name:
            continue

        if sln_hint:
            sln_path = info.get("solutionPath", "")
            if sln_hint.lower() not in sln_path.lower():
                continue

        pid = info.get("pid")
        if pid and not _is_process_alive(pid):
            continue

        return pipe_name

    raise RuntimeError(
        "No live VS IDE Bridge instance found.\n"
        "Start Visual Studio with the VS IDE Bridge extension, then retry."
    )


def _is_process_alive(pid: int) -> bool:
    try:
        import ctypes

        synchronize = 0x00100000
        handle = ctypes.windll.kernel32.OpenProcess(synchronize, False, pid)
        if handle == 0:
            return False
        ctypes.windll.kernel32.CloseHandle(handle)
        return True
    except Exception:
        return True


class PipeConnection:
    """Wrap a Win32 named pipe handle with line-oriented read/write."""

    def __init__(self, handle):
        self._handle = handle
        self._buf = b""

    @classmethod
    def connect(cls, pipe_name: str, timeout_ms: int = 10_000) -> "PipeConnection":
        try:
            import win32file
            import win32pipe
        except ImportError:
            sys.exit(
                "pywin32 is required: conda run -n superslicer pip install pywin32"
            )

        full_name = rf"\\.\pipe\{pipe_name}"
        deadline = time.monotonic() + timeout_ms / 1000.0

        while True:
            try:
                win32pipe.WaitNamedPipe(full_name, 2000)
                handle = win32file.CreateFile(
                    full_name,
                    win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                    0,
                    None,
                    win32file.OPEN_EXISTING,
                    0,
                    None,
                )
                win32pipe.SetNamedPipeHandleState(
                    handle, win32pipe.PIPE_READMODE_BYTE, None, None
                )
                return cls(handle)
            except Exception as exc:
                err_str = str(exc)
                if time.monotonic() >= deadline:
                    raise RuntimeError(
                        f"Could not connect to pipe '{full_name}': {exc}"
                    ) from exc
                if "does not exist" in err_str.lower() or "2" in err_str:
                    time.sleep(0.1)
                    continue
                raise RuntimeError(
                    f"Could not connect to pipe '{full_name}': {exc}"
                ) from exc

    def send_line(self, data: str) -> None:
        import win32file

        encoded = (data + "\n").encode("utf-8")
        win32file.WriteFile(self._handle, encoded)

    def recv_line(self) -> str:
        import win32file

        while b"\n" not in self._buf:
            _, chunk = win32file.ReadFile(self._handle, 8192)
            if not chunk:
                raise EOFError("Pipe closed unexpectedly")
            self._buf += chunk
        idx = self._buf.index(b"\n")
        line = self._buf[:idx].decode("utf-8")
        self._buf = self._buf[idx + 1 :]
        return line

    def close(self) -> None:
        try:
            import win32file

            win32file.CloseHandle(self._handle)
        except Exception:
            pass

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()


def _make_request(command: str, args: str | None, req_id: str | None = None) -> dict:
    return {
        "id": req_id or str(uuid.uuid4())[:8],
        "command": command,
        "args": args or "",
    }


def _send_recv(conn: PipeConnection, req: dict) -> dict:
    conn.send_line(json.dumps(req))
    raw = conn.recv_line()
    return json.loads(raw)


def _format_response(resp: dict, fmt: str) -> str:
    if fmt == "summary":
        summary = resp.get("Summary", resp.get("summary", ""))
        ok = resp.get("Success", resp.get("success", False))
        prefix = "OK" if ok else "FAIL"
        return f"[{prefix}] {summary}"

    if fmt == "keyvalue":
        lines = []
        for key, value in resp.items():
            if key == "Data":
                data = value or {}
                if isinstance(data, dict):
                    for data_key, data_value in data.items():
                        lines.append(f"data.{data_key}={data_value}")
                else:
                    lines.append(f"data={value}")
            else:
                lines.append(f"{key}={value}")
        return "\n".join(lines)

    return json.dumps(resp, indent=2)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Direct named pipe client for VS IDE Bridge (no PowerShell overhead)"
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "--cmd",
        metavar="COMMAND",
        help="Single command to execute (e.g. Tools.IdeGetState)",
    )
    group.add_argument(
        "--batch",
        metavar="FILE",
        help="JSON file containing a list of {command, args} objects",
    )

    parser.add_argument("--args", metavar="ARGS", default="", help="Arguments string for --cmd")
    parser.add_argument(
        "--sln",
        metavar="HINT",
        default=None,
        help="Solution path substring for discovery filtering",
    )
    parser.add_argument(
        "--out",
        metavar="FILE",
        default=None,
        help="Write full JSON response(s) to this file",
    )
    parser.add_argument(
        "--format",
        choices=["json", "summary", "keyvalue"],
        default="json",
        help="Output format (default: json)",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=10_000,
        metavar="MS",
        help="Connection timeout in milliseconds (default: 10000)",
    )
    args = parser.parse_args()

    try:
        pipe_name = find_pipe(args.sln)
    except RuntimeError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    if args.cmd:
        requests = [_make_request(args.cmd, args.args)]
    else:
        batch_path = Path(args.batch)
        if not batch_path.exists():
            print(f"ERROR: Batch file not found: {batch_path}", file=sys.stderr)
            sys.exit(1)
        batch = json.loads(batch_path.read_text(encoding="utf-8-sig"))
        requests = [_make_request(item["command"], item.get("args", "")) for item in batch]

    try:
        with PipeConnection.connect(pipe_name, timeout_ms=args.timeout) as conn:
            responses = []
            for req in requests:
                started = time.monotonic()
                resp = _send_recv(conn, req)
                elapsed_ms = (time.monotonic() - started) * 1000
                resp["_elapsed_ms"] = round(elapsed_ms, 1)
                responses.append(resp)
    except (RuntimeError, EOFError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    if args.out:
        out_path = Path(args.out)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        payload = responses[0] if len(responses) == 1 else responses
        out_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    any_failed = False
    for resp in responses:
        print(_format_response(resp, args.format))
        ok = resp.get("Success", resp.get("success", False))
        if not ok:
            any_failed = True

    sys.exit(1 if any_failed else 0)


if __name__ == "__main__":
    main()
