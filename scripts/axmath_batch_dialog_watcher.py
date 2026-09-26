from __future__ import annotations

import argparse
import ctypes
import datetime
import json
import time
import winreg
from pathlib import Path

import win32api
import win32con
import win32gui
import win32process

POLL = 0.20
kernel32 = ctypes.windll.kernel32
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
STILL_ACTIVE = 259


def process_alive(pid: int) -> bool:
    h = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not h:
        return False
    code = ctypes.c_ulong()
    ok = kernel32.GetExitCodeProcess(h, ctypes.byref(code))
    kernel32.CloseHandle(h)
    return bool(ok and code.value == STILL_ACTIVE)


def exe_basename(pid: int) -> str:
    try:
        h = win32api.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        try:
            buf = ctypes.create_unicode_buffer(32768)
            size = ctypes.c_ulong(len(buf))
            if kernel32.QueryFullProcessImageNameW(int(h), 0, buf, ctypes.byref(size)):
                return Path(buf.value).name.lower()
        finally:
            win32api.CloseHandle(h)
    except Exception:
        pass
    return ""


def read_control(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception:
        return {}


def read_waiting_convert():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\AxMath\WordCmds") as key:
            value, _ = winreg.QueryValueEx(key, "WaitingConvert")
        try:
            return int(value)
        except (TypeError, ValueError):
            return None
    except OSError:
        return None


def is_confirmed_batch_completion(waiting_convert) -> bool:
    return waiting_convert == 1


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--main-pid", type=int, required=True)
    ap.add_argument("--log", required=True)
    ap.add_argument("--control", required=True)
    args = ap.parse_args()

    log = Path(args.log)
    control = Path(args.control)
    handled_tokens: set[str] = set()
    ignored_dialogs: set[tuple[str, int]] = set()
    closed = 0
    log.write_text(
        f"{datetime.datetime.now().isoformat()} watcher start main_pid={args.main_pid}\n",
        encoding="utf-8",
    )

    while process_alive(args.main_pid):
        state = read_control(control)
        token = str(state.get("token") or "")
        if state.get("phase") != "awaiting_batch_dialog" or not token or token in handled_tokens:
            time.sleep(POLL)
            continue

        preexisting = {int(x) for x in (state.get("preexisting_dialog_pids") or []) if str(x).isdigit()}
        windows: list[int] = []
        win32gui.EnumWindows(lambda h, _: windows.append(h) or True, None)
        did_close = False

        for hwnd in windows:
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    continue
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if not pid or pid in preexisting or exe_basename(pid) != "axmmdlg.exe":
                    continue

                title = win32gui.GetWindowText(hwnd)
                waiting_convert = read_waiting_convert()
                # Historical successful batch-completion dialogs all reached
                # WaitingConvert=1; the recorded conflict/no-progress case had
                # WaitingConvert=0. Only confirmed completion is auto-closed.
                if not is_confirmed_batch_completion(waiting_convert):
                    key = (token, int(pid))
                    if key not in ignored_dialogs:
                        ignored_dialogs.add(key)
                        with log.open("a", encoding="utf-8") as f:
                            f.write(
                                f"{datetime.datetime.now().isoformat()} ignored non-completion dialog "
                                f"token={token} pid={pid} hwnd={hwnd} title={title!r} "
                                f"WaitingConvert={waiting_convert!r}\n"
                            )
                    continue
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                closed += 1
                handled_tokens.add(token)
                did_close = True
                with log.open("a", encoding="utf-8") as f:
                    f.write(
                        f"{datetime.datetime.now().isoformat()} closed batch dialog "
                        f"token={token} pid={pid} hwnd={hwnd} title={title!r}\n"
                    )
                break
            except Exception as e:
                with log.open("a", encoding="utf-8") as f:
                    f.write(f"{datetime.datetime.now().isoformat()} watcher error: {e!r}\n")

        if not did_close:
            time.sleep(POLL)

    with log.open("a", encoding="utf-8") as f:
        f.write(f"{datetime.datetime.now().isoformat()} watcher exit closed={closed}\n")
    print(f"closed={closed}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
