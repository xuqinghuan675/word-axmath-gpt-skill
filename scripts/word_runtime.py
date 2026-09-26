from __future__ import annotations

import gc
import time
from dataclasses import dataclass, asdict

import psutil
import pythoncom
import win32com.client
import win32process


@dataclass
class WordSessionMeta:
    pid: int | None = None
    hwnd: int | None = None
    owned_pid: bool = False
    preexisting_word_pids: list[int] | None = None
    forced_cleanup: bool = False
    quit_error: str | None = None
    cleanup_error: str | None = None

    def to_dict(self):
        return asdict(self)


class OwnedWord:
    """Isolated Word COM session with deterministic cleanup."""

    def __init__(self, *, visible: bool = False, require_clean: bool = True):
        self.visible = visible
        self.require_clean = require_clean
        self.app = None
        self.meta = WordSessionMeta()
        self._normal_saved_before = None
        self._coinit = False

    @staticmethod
    def _word_pids() -> list[int]:
        out = []
        for p in psutil.process_iter(["pid", "name"]):
            try:
                if (p.info.get("name") or "").lower() == "winword.exe":
                    out.append(int(p.info["pid"]))
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        return sorted(out)

    def __enter__(self):
        before = self._word_pids()
        self.meta.preexisting_word_pids = before
        if self.require_clean and before:
            raise RuntimeError(
                "Refusing to start Word automation while WINWORD.EXE already exists: "
                + ",".join(map(str, before))
            )

        pythoncom.CoInitialize()
        self._coinit = True
        self.app = win32com.client.DispatchEx("Word.Application")
        self.app.Visible = self.visible
        self.app.DisplayAlerts = 0

        # Prefer process-diff to identify the isolated WINWORD process. Some
        # Word builds do not expose a useful HWND immediately after DispatchEx.
        time.sleep(0.25)
        after = self._word_pids()
        created = [pid for pid in after if pid not in before]
        if len(created) == 1:
            self.meta.pid = int(created[0])
            self.meta.owned_pid = True
        try:
            self.app.ScreenUpdating = False
        except Exception:
            pass
        try:
            self.app.Options.SaveNormalPrompt = False
        except Exception:
            pass
        try:
            self._normal_saved_before = bool(self.app.NormalTemplate.Saved)
        except Exception:
            self._normal_saved_before = None

        try:
            hwnd = int(self.app.Hwnd)
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            self.meta.hwnd = hwnd
            if not self.meta.pid and pid:
                self.meta.pid = int(pid)
                self.meta.owned_pid = int(pid) not in before
        except Exception:
            pass
        return self.app, self.meta

    def __exit__(self, exc_type, exc, tb):
        app = self.app
        pid = self.meta.pid

        if app is not None:
            try:
                if self._normal_saved_before is True and not bool(app.NormalTemplate.Saved):
                    app.NormalTemplate.Saved = True
            except Exception:
                pass
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            try:
                app.Quit()
            except Exception as e:
                self.meta.quit_error = repr(e)

        self.app = None
        app = None
        gc.collect()
        gc.collect()

        if self._coinit:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._coinit = False

        if pid and self.meta.owned_pid:
            deadline = time.time() + 6.0
            while time.time() < deadline:
                if not psutil.pid_exists(pid):
                    break
                time.sleep(0.25)
            if psutil.pid_exists(pid):
                try:
                    p = psutil.Process(pid)
                    p.terminate()
                    try:
                        p.wait(4)
                    except psutil.TimeoutExpired:
                        p.kill()
                        p.wait(4)
                    self.meta.forced_cleanup = True
                except psutil.NoSuchProcess:
                    pass
                except Exception as e:
                    self.meta.cleanup_error = repr(e)
        return False


def clean_com_text(value: str | None) -> str:
    """Normalize UTF-16 surrogate pairs returned by some Word COM calls."""
    s = value or ""
    try:
        return s.encode("utf-16", "surrogatepass").decode("utf-16", "replace")
    except Exception:
        return "".join("\ufffd" if 0xD800 <= ord(ch) <= 0xDFFF else ch for ch in s)


def text_anchor(doc, start: int, end: int, span: int = 48) -> dict[str, str]:
    content_end = max(0, int(doc.Content.End) - 1)
    b0 = max(0, int(start) - span)
    a1 = min(content_end, int(end) + span)
    before = doc.Range(b0, max(b0, int(start))).Text if int(start) >= b0 else ""
    after = doc.Range(min(content_end, int(end)), a1).Text if a1 >= int(end) else ""
    return {
        "before": clean_com_text(before).replace("\r", "\n"),
        "after": clean_com_text(after).replace("\r", "\n"),
    }
