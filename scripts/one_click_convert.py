from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

import psutil

from audit_docx import analyze


SKILL_PATH = Path(__file__).resolve().parents[1] / "SKILL.md"


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def word_pids() -> list[int]:
    out: list[int] = []
    for p in psutil.process_iter(["pid", "name"]):
        try:
            if (p.info.get("name") or "").lower() == "winword.exe":
                out.append(int(p.info["pid"]))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return sorted(out)


def discover_docs(target: Path) -> list[Path]:
    if target.is_file():
        return [target]
    if target.is_dir():
        return sorted(p.resolve() for p in target.glob("*.docx") if not p.name.startswith("~$"))
    raise FileNotFoundError(str(target))


def inspect_source(path: Path) -> dict:
    row = {
        "source": str(path),
        "name": path.name,
        "stem": path.stem,
        "exists": path.is_file(),
        "lock_present": path.with_name("~$" + path.name).exists(),
    }
    if not path.is_file():
        row["state"] = "missing"
        return row
    try:
        stats = analyze(path)
        row.update({
            "sha256": sha256_file(path),
            "paragraphs": stats["paragraphs"],
            "omath": stats["omath"],
            "axmath": stats["axmath_ole"],
        })
    except Exception as exc:
        row["state"] = "invalid_docx"
        row["error"] = str(exc)
        return row

    if row["lock_present"]:
        row["state"] = "blocked_locked"
    elif row["omath"] > 0 and row["axmath"] == 0:
        row["state"] = "ready_officemath"
    elif row["omath"] > 0 and row["axmath"] > 0:
        row["state"] = "mixed_math_needs_review"
    elif row["omath"] == 0 and row["axmath"] > 0:
        row["state"] = "already_axmath"
    else:
        row["state"] = "no_native_math"
    return row


def workspace_root_for(target: Path) -> Path:
    if target.is_file():
        return target.parent / f"{target.stem}_AxMath-workspace"
    return target.parent / f"{target.name}_AxMath-workspace"


def unique_run_dir(root: Path) -> Path:
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    candidate = root / f"run-{stamp}"
    n = 2
    while candidate.exists():
        candidate = root / f"run-{stamp}-{n}"
        n += 1
    return candidate


def write_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def run_conversion(source: Path, doc_dir: Path) -> dict:
    source_dir = doc_dir / "source"
    working_dir = doc_dir / "working"
    logs_dir = doc_dir / "logs"
    source_dir.mkdir(parents=True, exist_ok=True)
    working_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)

    frozen_source = source_dir / source.name
    shutil.copy2(source, frozen_source)
    working = working_dir / f"{source.stem}_AxMath-working.docx"
    log_path = logs_dir / "conversion.log.json"

    cmd = [
        sys.executable,
        str(Path(__file__).resolve().parent / "run_skill.py"),
        "--input", str(frozen_source),
        "--output", str(working),
    ]
    started = time.time()
    proc = subprocess.run(cmd, cwd=str(Path(__file__).resolve().parent), text=True, capture_output=True)
    write_json(log_path, {
        "returncode": proc.returncode,
        "seconds": time.time() - started,
        "stdout_tail": proc.stdout[-12000:],
        "stderr_tail": proc.stderr[-12000:],
    })

    source_stats = analyze(frozen_source)
    candidate_stats = analyze(working) if working.exists() else None
    skill_report_path = Path(str(working) + ".skill-report.json")
    skill_report = None
    if skill_report_path.exists():
        try:
            skill_report = json.loads(skill_report_path.read_text(encoding="utf-8-sig"))
        except Exception:
            skill_report = None
    audit = (skill_report or {}).get("audit") if isinstance(skill_report, dict) else None

    success = bool(
        proc.returncode == 0
        and candidate_stats
        and candidate_stats["omath"] == 0
        and candidate_stats["axmath_ole"] == source_stats["omath"]
        and audit
        and audit.get("paragraph_count_equal")
        and audit.get("nonmath_text_exact")
    )
    return {
        "status": "converted_ready_for_gpt_review" if success else "conversion_failed",
        "original_source": str(source),
        "original_source_sha256": sha256_file(source),
        "frozen_source": str(frozen_source),
        "working_docx": str(working) if working.exists() else None,
        "source_omath": source_stats["omath"],
        "working_omath": candidate_stats["omath"] if candidate_stats else None,
        "working_axmath": candidate_stats["axmath_ole"] if candidate_stats else None,
        "paragraph_count_equal": bool(audit and audit.get("paragraph_count_equal")),
        "nonmath_text_exact": bool(audit and audit.get("nonmath_text_exact")),
        "conversion_log": str(log_path),
    }


def main() -> int:
    ap = argparse.ArgumentParser(description="Preflight -> adjacent workspace -> OfficeMath to AxMath -> GPT-review handoff")
    group = ap.add_mutually_exclusive_group(required=True)
    group.add_argument("--input", help="DOCX file or directory")
    group.add_argument("--input-dir", help="Compatibility alias for a directory")
    ap.add_argument("--inspect-only", action="store_true")
    args = ap.parse_args()

    target = Path(args.input or args.input_dir).resolve()
    docs = discover_docs(target)
    states = [inspect_source(p) for p in docs]
    preflight = {
        "schema": "axmath-preflight/v1",
        "target": str(target),
        "checked_at": dt.datetime.now().isoformat(),
        "word_pids": word_pids(),
        "documents": states,
    }

    if args.inspect_only:
        print(json.dumps(preflight, ensure_ascii=False, indent=2))
        return 0

    if preflight["word_pids"]:
        preflight["status"] = "blocked_word_running"
        print(json.dumps(preflight, ensure_ascii=False, indent=2))
        return 2

    ready = [Path(x["source"]) for x in states if x.get("state") == "ready_officemath"]
    if not ready:
        preflight["status"] = "nothing_to_convert"
        print(json.dumps(preflight, ensure_ascii=False, indent=2))
        return 0

    run_dir = unique_run_dir(workspace_root_for(target))
    run_dir.mkdir(parents=True)
    write_json(run_dir / "SOURCE_STATE.json", preflight)

    results = []
    for source in ready:
        result = run_conversion(source, run_dir / source.stem)
        results.append(result)
        write_json(run_dir / source.stem / "CONVERSION_RESULT.json", result)

    failed = [x for x in results if x["status"] != "converted_ready_for_gpt_review"]
    succeeded = [x for x in results if x["status"] == "converted_ready_for_gpt_review"]
    manifest = {
        "schema": "axmath-gpt-review-handoff/v1",
        "status": "ready" if succeeded and not failed else ("partial" if succeeded else "blocked"),
        "created_at": dt.datetime.now().isoformat(),
        "skill": str(SKILL_PATH),
        "workspace": str(run_dir),
        "documents": results,
        "review_rule": "Read SKILL.md first. Conversion is finished; GPT review/repair has not run yet.",
    }
    write_json(run_dir / "READY_FOR_GPT_REVIEW.json", manifest)

    final = {
        "status": "ready_for_gpt_review" if succeeded and not failed else ("partial" if succeeded else "failed"),
        "workspace": str(run_dir),
        "ready_manifest": str(run_dir / "READY_FOR_GPT_REVIEW.json"),
        "documents": results,
    }
    write_json(run_dir / "RUN_STATE.json", final)
    print(json.dumps(final, ensure_ascii=False, indent=2))
    return 0 if succeeded and not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
