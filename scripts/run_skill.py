from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import shutil
import sys
import time
from pathlib import Path

import psutil

from audit_docx import compare


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _word_pids() -> list[int]:
    out = []
    for p in psutil.process_iter(["pid", "name"]):
        try:
            if (p.info.get("name") or "").lower() == "winword.exe":
                out.append(int(p.info["pid"]))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return sorted(out)


def _performance_summary(conversion: dict | None) -> dict | None:
    if not isinstance(conversion, dict):
        return None

    batches = conversion.get("batches") or []
    if not isinstance(batches, list):
        batches = []

    def _num(value, default=0.0) -> float:
        try:
            return float(value)
        except (TypeError, ValueError):
            return float(default)

    macro_seconds = _num(
        conversion.get("macro_seconds_total"),
        sum(_num(x.get("seconds")) for x in batches if isinstance(x, dict)),
    )
    save_seconds = _num(
        conversion.get("save_seconds_total"),
        sum(_num(x.get("save_seconds")) for x in batches if isinstance(x, dict)),
    )
    total_seconds = _num(conversion.get("total_seconds"))
    converted_total = sum(
        int(_num(x.get("converted"))) for x in batches if isinstance(x, dict)
    )
    observed_max_batch = int(
        _num(
            conversion.get("observed_max_batch_converted"),
            max(
                (int(_num(x.get("converted"))) for x in batches if isinstance(x, dict)),
                default=0,
            ),
        )
    )
    macro_share = _num(
        conversion.get("macro_share_percent"),
        (100.0 * macro_seconds / total_seconds) if total_seconds > 0 else 0.0,
    )

    slowest = sorted(
        (x for x in batches if isinstance(x, dict)),
        key=lambda x: _num(x.get("seconds")),
        reverse=True,
    )[:5]
    top_slow_batches = [
        {
            "batch": x.get("batch"),
            "converted": x.get("converted"),
            "macro_seconds": _num(x.get("seconds")),
            "save_seconds": _num(x.get("save_seconds")),
            "crash_dump_count": len(x.get("crash_dumps_new") or []),
        }
        for x in slowest
    ]

    return {
        "batch_count": len(batches),
        "converted_total": converted_total,
        "macro_seconds_total": macro_seconds,
        "save_seconds_total": save_seconds,
        "total_seconds": total_seconds,
        "macro_share_percent": macro_share,
        "macro_seconds_per_formula": (
            macro_seconds / converted_total if converted_total else None
        ),
        "observed_max_batch_converted": observed_max_batch,
        "new_crash_dump_count": len(conversion.get("new_crash_dumps") or []),
        "batch_limit_owner": conversion.get("batch_limit_owner") or "AxMath plugin",
        "top_slow_batches": top_slow_batches,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--overwrite-output", action="store_true")
    args = ap.parse_args()

    src = Path(args.input).resolve()
    out = Path(args.output).resolve()
    if not src.is_file():
        raise FileNotFoundError(src)
    if src == out:
        raise RuntimeError("Refusing to overwrite the input DOCX.")
    if out.exists() and not args.overwrite_output:
        raise FileExistsError(f"Output already exists: {out}")
    out.parent.mkdir(parents=True, exist_ok=True)
    source_sha_before = _sha256_file(src)
    here = Path(__file__).resolve().parent
    conv = here / "convert_officemath_to_axmath.ps1"
    watcher_script = here / "axmath_batch_dialog_watcher.py"
    conv_report = Path(str(out) + ".conversion.json")
    final_report = Path(str(out) + ".skill-report.json")
    watcher_log = Path(str(out) + ".dialog-watcher.log")
    control = Path(str(out) + ".control.json")
    t0 = time.perf_counter()

    preexisting_word = _word_pids()

    powershell = shutil.which("pwsh") or shutil.which("powershell")
    if not powershell:
        raise RuntimeError("PowerShell not found (pwsh.exe or powershell.exe).")
    cmd = [
        powershell,
        "-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", str(conv),
        "-InputDocx", str(src),
        "-OutputDocx", str(out),
        "-ReportPath", str(conv_report),
        "-ControlPath", str(control),
    ]
    if args.overwrite_output:
        cmd.append("-OverwriteOutput")
    p = subprocess.Popen(
        cmd,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    watcher = subprocess.Popen(
        [
            sys.executable, str(watcher_script),
            "--main-pid", str(p.pid),
            "--log", str(watcher_log),
            "--control", str(control),
        ],
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )

    stdout, stderr = p.communicate()
    try:
        watcher_stdout, watcher_stderr = watcher.communicate(timeout=8)
    except subprocess.TimeoutExpired:
        watcher.terminate()
        try:
            watcher_stdout, watcher_stderr = watcher.communicate(timeout=5)
        except subprocess.TimeoutExpired:
            watcher.kill()
            watcher_stdout, watcher_stderr = watcher.communicate()

    source_sha_after = _sha256_file(src)
    source_unchanged = source_sha_after == source_sha_before
    conversion = None
    if conv_report.exists():
        conversion = json.loads(conv_report.read_text(encoding="utf-8-sig"))
    audit = compare(src, out) if out.exists() else None

    # Word can linger briefly while COM/OLE tears down. Judge only the PID that
    # the converter proved belongs to this run; an unrelated Word session that
    # starts during conversion must never become a false leak/failure.
    owned_word_pid = None
    if conversion and conversion.get("word_pid_owned") and conversion.get("word_pid"):
        try:
            owned_word_pid = int(conversion["word_pid"])
        except (TypeError, ValueError):
            owned_word_pid = None
    deadline = time.time() + 6.0
    lingering_word = [owned_word_pid] if owned_word_pid and psutil.pid_exists(owned_word_pid) else []
    while lingering_word and time.time() < deadline:
        time.sleep(0.25)
        lingering_word = [owned_word_pid] if owned_word_pid and psutil.pid_exists(owned_word_pid) else []

    report = {
        "runner_seconds": time.perf_counter() - t0,
        "returncode": p.returncode,
        "conversion": conversion,
        "performance": _performance_summary(conversion),
        "audit": audit,
        "source_sha256_before": source_sha_before,
        "source_sha256_after": source_sha_after,
        "source_unchanged": source_unchanged,
        "preexisting_word_pids": preexisting_word,
        "lingering_word_pids": lingering_word,
        "stdout_tail": stdout[-8000:],
        "stderr_tail": stderr[-8000:],
        "watcher_log": str(watcher_log),
        "watcher_stdout_tail": watcher_stdout[-4000:],
        "watcher_stderr_tail": watcher_stderr[-4000:],
    }
    final_report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=True, indent=2))

    ok = bool(
        conversion
        and conversion.get("success")
        and conversion.get("complete")
        and not lingering_word
        and audit
        and audit.get("paragraph_count_equal")
        and audit.get("nonmath_text_exact")
        and source_unchanged
    )
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
