"""Single-user AxMath automation mutex.

AxMath's HKCU registry completion flag is shared by Word instances. Serialize
this skill's conversion runners without altering arbitrary user Word sessions.
"""
from __future__ import annotations

import contextlib
import os
import tempfile
from pathlib import Path


@contextlib.contextmanager
def conversion_lock():
    folder = Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()) / "AxMath"
    folder.mkdir(parents=True, exist_ok=True)
    lock_path = folder / "word_axmath_gpt_skill_conversion.lock"
    with lock_path.open("a+b") as handle:
        handle.seek(0, os.SEEK_END)
        if handle.tell() == 0:
            handle.write(b"1")
            handle.flush()
        handle.seek(0)
        if os.name == "nt":
            import msvcrt
            try:
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError as exc:
                raise RuntimeError(
                    "Another AxMath Skill conversion holds the exclusive batch mutex."
                ) from exc
            try:
                yield lock_path
            finally:
                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            import fcntl
            try:
                fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
            except OSError as exc:
                raise RuntimeError("Another AxMath conversion holds the mutex.") from exc
            try:
                yield lock_path
            finally:
                fcntl.flock(handle, fcntl.LOCK_UN)
