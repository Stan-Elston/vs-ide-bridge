#!/usr/bin/env python3
"""
Find a codex window, capture a screenshot, and send prompt text to it.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys
import time
from typing import Any

import psutil
from PIL import ImageGrab
import win32clipboard
import win32con
import win32gui
import win32process
from pywinauto import keyboard, mouse


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


def safe_process_name(pid: int) -> str:
    if pid <= 0:
        return ""
    try:
        return psutil.Process(pid).name()
    except Exception:
        return ""


def enumerate_windows(title_contains: str) -> list[dict[str, Any]]:
    matches: list[dict[str, Any]] = []
    needle = title_contains.lower()

    def callback(hwnd: int, _lparam: int) -> bool:
        if not win32gui.IsWindowVisible(hwnd):
            return True

        title = safe_window_text(hwnd)
        if not title:
            return True

        if needle not in title.lower():
            return True

        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = max(0, right - left)
        height = max(0, bottom - top)
        area = width * height
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        matches.append(
            {
                "hwnd": f"0x{hwnd:08X}",
                "title": title,
                "className": safe_class_name(hwnd),
                "pid": pid,
                "processName": safe_process_name(pid),
                "rect": {"left": left, "top": top, "right": right, "bottom": bottom},
                "area": area,
            }
        )
        return True

    win32gui.EnumWindows(callback, 0)
    matches.sort(key=lambda item: item["area"], reverse=True)
    return matches


def ensure_parent_dir(path: str) -> None:
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)


def capture_window_screenshot(hwnd: int, screenshot_path: str) -> None:
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    if right <= left or bottom <= top:
        raise RuntimeError("Window rectangle is empty.")

    image = ImageGrab.grab(bbox=(left, top, right, bottom), all_screens=True)
    ensure_parent_dir(screenshot_path)
    image.save(screenshot_path)


def set_clipboard_text(text: str) -> None:
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32con.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()


def activate_window(hwnd: int) -> None:
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    except Exception:
        pass

    try:
        win32gui.SetForegroundWindow(hwnd)
        return
    except Exception:
        pass

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    click_x = left + max(10, (right - left) // 2)
    click_y = top + max(10, (bottom - top) // 2)
    mouse.click(button="left", coords=(click_x, click_y))


def send_prompt_text(prompt: str, press_enter: bool) -> None:
    set_clipboard_text(prompt)
    keyboard.send_keys("^v", pause=0.02)
    if press_enter:
        keyboard.send_keys("{ENTER}", pause=0.02)


def default_screenshot_path() -> str:
    timestamp = dt.datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    return rf"C:\Users\elsto\source\repos\vs-ide-bridge\output\codex_window_{timestamp}.png"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Capture codex window screenshot and send prompt text.")
    parser.add_argument(
        "--title-contains",
        default="codex",
        help="Case-insensitive title substring used to select the target window.",
    )
    parser.add_argument(
        "--window-index",
        type=int,
        default=0,
        help="Index in matched window list after sorting by area (0 is largest).",
    )
    parser.add_argument(
        "--screenshot-path",
        default=default_screenshot_path(),
        help="Output PNG path for the captured window screenshot.",
    )
    parser.add_argument(
        "--prompt",
        required=True,
        help="Text to paste into the window prompt.",
    )
    parser.add_argument(
        "--no-enter",
        action="store_true",
        help="Paste prompt text without sending Enter.",
    )
    parser.add_argument(
        "--focus-delay-ms",
        type=int,
        default=500,
        help="Delay after focusing the window before typing.",
    )
    parser.add_argument(
        "--list-only",
        action="store_true",
        help="Only list matching windows without screenshot or typing.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    windows = enumerate_windows(args.title_contains)
    log_event("windows-found", {"count": len(windows), "titleContains": args.title_contains, "windows": windows})

    if not windows:
        log_event("error", {"message": "No matching windows found."})
        return 2

    if args.window_index < 0 or args.window_index >= len(windows):
        log_event(
            "error",
            {
                "message": "window-index out of range.",
                "windowIndex": args.window_index,
                "maxIndex": len(windows) - 1,
            },
        )
        return 2

    target = windows[args.window_index]
    hwnd_int = int(target["hwnd"], 16)
    log_event("target-window", target)

    if args.list_only:
        return 0

    capture_window_screenshot(hwnd_int, args.screenshot_path)
    log_event("screenshot", {"path": args.screenshot_path})

    activate_window(hwnd_int)
    time.sleep(max(args.focus_delay_ms, 0) / 1000.0)
    send_prompt_text(args.prompt, press_enter=not args.no_enter)
    log_event("prompt-sent", {"promptLength": len(args.prompt), "pressedEnter": not args.no_enter})
    return 0


if __name__ == "__main__":
    sys.exit(main())
