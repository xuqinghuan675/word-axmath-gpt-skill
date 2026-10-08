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
from execution_lock import conversion_lock


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


def _execute(args):
    src = Path(args.input).resolve()
    out = Path(args.output).resolve()
    if not src.is_file():
        raise FileNotFoundError(src)
    if src == out:
        raise RuntimeError("Refusing to overwrite the input DOCX.")
    if out.exists() and not args.overwrite_output:
        if not args.resume:
            raise FileExistsError(f"Output already exists: {out}")
    if args.resume and args.overwrite_output:
        raise ValueError("Resume and overwrite cannot be combined.")
    if args.resume:
        checkpoint = Path(str(out) + ".conversion-checkpoint.json")
        if not out.is_file() or not checkpoint.is_file():
            raise FileNotFoundError("Resume requires existing working DOCX and saved checkpoint.")
        state = json.loads(checkpoint.read_text(encoding="utf-8-sig"))
        if (
            state.get("schema") != "axmath-conversion-checkpoint/v1"
            or state.get("source_sha256") != _sha256_file(src)
            or state.get("working_sha256") != _sha256_file(out)
            or Path(state.get("source_path") or "").resolve() != src
            or Path(state.get("working_path") or "").resolve() != out
        ):
            raise ValueError("Resume checkpoint does not match frozen source and saved working DOCX.")
    out.parent.mkdir(parents=True, exist_ok=True)
    source_sha_before = _sha256_file(src)
    here = Path(__file__).resolve().parent
    conv = here / "convert_officemath_to_axmath.ps1"
    watcher_script = here / "axmath_batch_dialog_watcher.py"
    conv_report = Path(str(out) + ".conversion.json")
    previous_conv_report_time_ns = (
        conv_report.stat().st_mtime_ns if conv_report.exists() else None
    )
    final_report = Path(str(out) + ".skill-report.json")
    watcher_log = Path(str(out) + ".dialog-watcher.log")
    control = Path(str(out) + ".control.json")
    watchdog_report = Path(str(out) + f".watchdog-{time.time_ns()}.json")
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
    if args.resume:
        cmd.append("-ResumeOutput")
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
            "--stall-seconds", str(args.watchdog_seconds),
            "--timeout-report", str(watchdog_report),
        ],
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )

    # Do not wait forever in communicate() after the dialog watchdog has
    # already failed. Only the Popen we created here may be terminated;
    # pre-existing Word sessions are never touched.
    watcher_lost_while_running = False
    while p.poll() is None:
        if watcher.poll() is not None:
            time.sleep(0.5)  # avoid a false race during normal shutdown
            if p.poll() is None:
                watcher_lost_while_running = True
                p.terminate()
                try:
                    p.wait(timeout=8)
                except subprocess.TimeoutExpired:
                    p.kill()
                    p.wait(timeout=8)
            break
        time.sleep(0.35)
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
    report_fresh = bool(
        conv_report.exists()
        and (
            previous_conv_report_time_ns is None
            or conv_report.stat().st_mtime_ns != previous_conv_report_time_ns
        )
    )
    if report_fresh:
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
        "conversion_report_fresh": report_fresh,
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
        "watcher_returncode": watcher.returncode,
        "watcher_lost_while_converter_running": watcher_lost_while_running,
        "watchdog_report": str(watchdog_report) if watchdog_report.exists() else None,
        "resume_requested": bool(args.resume),
    }
    final_report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    # The parent batch entrypoint reads the persisted report directly.
    # Avoid piping a second multi-megabyte JSON audit through stdout.
    if not getattr(args, "quiet", False):
        print(json.dumps(report, ensure_ascii=True, indent=2))

    ok = bool(
        p.returncode == 0
        and report_fresh
        and conversion
        and conversion.get("success")
        and conversion.get("complete")
        and watcher.returncode == 0
        and not watcher_lost_while_running
        and not watchdog_report.exists()
        and not lingering_word
        and audit
        and audit.get("paragraph_count_equal")
        and audit.get("nonmath_text_exact")
        and source_unchanged
    )
    if ok:
        # The successful converter/control state is fully included in the
        # skill report; retain transport files only for failed-run debugging.
        conv_report.unlink(missing_ok=True)
        control.unlink(missing_ok=True)
    return 0 if ok else 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--overwrite-output", action="store_true")
    ap.add_argument("--resume", action="store_true", help="Continue a SHA-bound saved conversion checkpoint")
    ap.add_argument("--watchdog-seconds", type=int, default=1800)
    ap.add_argument("--quiet", action="store_true", help="Write the normal skill-report.json without duplicating it on stdout")
    args = ap.parse_args()
    if args.watchdog_seconds < 60:
        ap.error("--watchdog-seconds must be >= 60")
    with conversion_lock():
        return _execute(args)


if __name__ == "__main__":
    raise SystemExit(main())
