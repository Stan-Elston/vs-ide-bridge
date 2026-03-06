#!/usr/bin/env python3
"""
Launch vs-ide-bridge ensure and capture Visual Studio popup handles in real time.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import subprocess
import sys
import time
from typing import Any

import psutil
import win32con
import win32gui
import win32process

TARGET_TEXT = "Exception has been thrown by the target of an invocation."
DIALOG_CLASS = "#32770"
VISUAL_STUDIO_PROCESS = "devenv.exe"


def utc_now_iso() -> str:
    return dt.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S.%fZ")


def log_event(event: str, payload: dict[str, Any]) -> None:
    output = {"timestampUtc": utc_now_iso(), "event": event}
    output.update(payload)
    print(json.dumps(output, ensure_ascii=True), flush=True)


def safe_window_text(hwnd: int) -> str:
    try:
        return win32gui.GetWindowText(hwnd).strip()
    except Exception:
        return ""


def safe_class_name(hwnd: int) -> str:
    try:
        return win32gui.GetClassName(hwnd)
    except Exception:
        return ""


def safe_window_pid(hwnd: int) -> int:
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid
    except Exception:
        return 0


def safe_process_name(pid: int) -> str:
    if pid <= 0:
        return ""
    try:
        return psutil.Process(pid).name()
    except Exception:
        return ""


def get_existing_devenv_pids() -> set[int]:
    pids: set[int] = set()
    for proc in psutil.process_iter(["name"]):
        try:
            name = (proc.info.get("name") or "").lower()
            if name == VISUAL_STUDIO_PROCESS:
                pids.add(proc.pid)
        except Exception:
            continue
    return pids


def collect_child_text(hwnd: int) -> list[dict[str, Any]]:
    children: list[dict[str, Any]] = []

    def callback(child_hwnd: int, _lparam: int) -> bool:
        text = safe_window_text(child_hwnd)
        cls = safe_class_name(child_hwnd)
        if text:
            children.append(
                {
                    "hwnd": f"0x{child_hwnd:08X}",
                    "className": cls,
                    "text": text,
                }
            )
        return True

    try:
        win32gui.EnumChildWindows(hwnd, callback, 0)
    except Exception:
        pass
    return children


def enumerate_candidate_windows() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    def callback(hwnd: int, _lparam: int) -> bool:
        if not win32gui.IsWindow(hwnd):
            return True

        title = safe_window_text(hwnd)
        class_name = safe_class_name(hwnd)
        pid = safe_window_pid(hwnd)
        process_name = safe_process_name(pid).lower()
        is_vs_process = process_name == VISUAL_STUDIO_PROCESS
        is_vs_title = "visual studio" in title.lower()
        is_dialog = class_name == DIALOG_CLASS

        if not (is_vs_process or is_vs_title or is_dialog):
            return True

        children = collect_child_text(hwnd) if is_dialog else []
        text_blobs = [title] + [item["text"] for item in children]
        joined = "\n".join(part for part in text_blobs if part)
        target_match = TARGET_TEXT in joined

        rows.append(
            {
                "hwnd": f"0x{hwnd:08X}",
                "title": title,
                "className": class_name,
                "pid": pid,
                "processName": process_name,
                "childText": children,
                "targetMatch": target_match,
            }
        )
        return True

    win32gui.EnumWindows(callback, 0)
    return rows


def click_ok_button(dialog_hwnd: int) -> bool:
    clicked = False

    def callback(child_hwnd: int, _lparam: int) -> bool:
        nonlocal clicked
        if clicked:
            return False

        cls = safe_class_name(child_hwnd)
        txt = safe_window_text(child_hwnd)
        if cls == "Button" and txt.strip().upper() == "OK":
            win32gui.PostMessage(child_hwnd, win32con.BM_CLICK, 0, 0)
            clicked = True
            return False
        return True

    try:
        win32gui.EnumChildWindows(dialog_hwnd, callback, 0)
    except Exception:
        return False

    return clicked


def run_probe(args: argparse.Namespace) -> int:
    baseline_devenv = get_existing_devenv_pids()
    log_event(
        "baseline",
        {
            "baselineDevenvPids": sorted(baseline_devenv),
            "bridgeExe": args.bridge_exe,
            "solution": args.solution,
        },
    )

    ensure_command = [
        args.bridge_exe,
        "ensure",
        "--solution",
        args.solution,
        "--json",
    ]
    log_event("launch", {"command": ensure_command})

    ensure_process = subprocess.Popen(
        ensure_command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    seen_hwnds: set[str] = set()
    matched_dialogs: list[str] = []
    deadline = time.time() + args.timeout_seconds
    interval = max(args.poll_ms, 50) / 1000.0

    try:
        while time.time() < deadline:
            active_devenv = get_existing_devenv_pids()
            new_devenv = sorted(active_devenv - baseline_devenv)
            windows = enumerate_candidate_windows()
            for window in windows:
                hwnd = window["hwnd"]
                pid = int(window["pid"])
                process_name = (window["processName"] or "").lower()
                is_new_devenv_window = process_name == VISUAL_STUDIO_PROCESS and pid in new_devenv
                should_report = is_new_devenv_window or bool(window["targetMatch"])

                if not should_report or hwnd in seen_hwnds:
                    continue

                seen_hwnds.add(hwnd)
                log_event("window", window)

                if window["targetMatch"]:
                    matched_dialogs.append(hwnd)
                    if args.auto_ok:
                        hwnd_int = int(hwnd, 16)
                        clicked = click_ok_button(hwnd_int)
                        log_event("dialog-action", {"hwnd": hwnd, "action": "click-ok", "clicked": clicked})

            if ensure_process.poll() is not None and matched_dialogs:
                break

            time.sleep(interval)
    finally:
        return_code = ensure_process.poll()
        if return_code is None:
            try:
                ensure_process.terminate()
            except Exception:
                pass

        try:
            stdout, stderr = ensure_process.communicate(timeout=5)
        except Exception:
            stdout, stderr = "", ""

        log_event(
            "ensure-exit",
            {
                "returnCode": ensure_process.returncode,
                "stdout": stdout.strip(),
                "stderr": stderr.strip(),
            },
        )

    if matched_dialogs:
        log_event("match-summary", {"matchedCount": len(matched_dialogs), "matchedHwnds": matched_dialogs})
        return 0

    log_event("match-summary", {"matchedCount": 0, "matchedHwnds": []})
    return 2


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run bridge ensure and capture VS popup hwnd/text in real time.",
    )
    parser.add_argument(
        "--bridge-exe",
        default=r"C:\Users\elsto\source\repos\vs-ide-bridge\src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe",
        help="Path to vs-ide-bridge.exe",
    )
    parser.add_argument(
        "--solution",
        default=r"C:\Users\elsto\source\repos\vs-ide-bridge\VsIdeBridge.sln",
        help="Path to solution used for ensure.",
    )
    parser.add_argument(
        "--timeout-seconds",
        type=int,
        default=120,
        help="How long to monitor for popup windows.",
    )
    parser.add_argument(
        "--poll-ms",
        type=int,
        default=150,
        help="Window polling interval in milliseconds.",
    )
    parser.add_argument(
        "--auto-ok",
        action="store_true",
        help="Click OK on the matched dialog when found.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return run_probe(args)


if __name__ == "__main__":
    sys.exit(main())
