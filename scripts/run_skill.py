from __future__ import annotations

import argparse
import json
import subprocess
import shutil
import sys
import time
from pathlib import Path

import psutil

from audit_docx import compare


def _word_pids() -> list[int]:
    out = []
    for p in psutil.process_iter(["pid", "name"]):
        try:
            if (p.info.get("name") or "").lower() == "winword.exe":
                out.append(int(p.info["pid"]))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return sorted(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", required=True)
    args = ap.parse_args()

    src = Path(args.input).resolve()
    out = Path(args.output).resolve()
    here = Path(__file__).resolve().parent
    conv = here / "convert_officemath_to_axmath.ps1"
    watcher_script = here / "axmath_batch_dialog_watcher.py"
    conv_report = Path(str(out) + ".conversion.json")
    final_report = Path(str(out) + ".skill-report.json")
    watcher_log = Path(str(out) + ".dialog-watcher.log")
    control = Path(str(out) + ".control.json")
    t0 = time.perf_counter()

    preexisting_word = _word_pids()
    if preexisting_word:
        report = {
            "runner_seconds": time.perf_counter() - t0,
            "returncode": 2,
            "conversion": None,
            "audit": None,
            "error": "Word is already running; refusing unattended conversion.",
            "preexisting_word_pids": preexisting_word,
        }
        final_report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 2

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
        watcher_stdout, watcher_stderr = watcher.communicate(timeout=5)

    conversion = None
    if conv_report.exists():
        conversion = json.loads(conv_report.read_text(encoding="utf-8-sig"))
    audit = compare(src, out) if out.exists() else None

    # Word can linger briefly while COM/OLE tears down. Give only this
    # automation-created instance a bounded grace period before judging it a leak.
    deadline = time.time() + 6.0
    lingering_word = _word_pids()
    while lingering_word and time.time() < deadline:
        time.sleep(0.25)
        lingering_word = _word_pids()

    report = {
        "runner_seconds": time.perf_counter() - t0,
        "returncode": p.returncode,
        "conversion": conversion,
        "audit": audit,
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
    )
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
