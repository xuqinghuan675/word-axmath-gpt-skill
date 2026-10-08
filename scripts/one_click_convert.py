from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

import psutil

from audit_docx import analyze, compare as compare_docx
from normalize_numbering_punctuation import normalize_docx, numbering_dunhao_positions


SKILL_PATH = Path(__file__).resolve().parents[1] / "SKILL.md"
SCRIPTS_DIR = Path(__file__).resolve().parent


def _utf8_console() -> None:
    # Chinese paths must never make inspect-only/preflight fail because the
    # parent console happens to expose cp1252 or another legacy encoding.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


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


def _axmath_template_candidates() -> list[Path]:
    roots = [
        os.environ.get("ProgramFiles(x86)"),
        os.environ.get("ProgramFiles"),
    ]
    return [
        Path(root) / "AxMath" / "MSOffice" / "AxMath.dotm"
        for root in roots
        if root
    ]


def environment_preflight() -> dict:
    modules = {
        "psutil": "psutil",
        "pywin32": "win32com",
        "lxml": "lxml",
        "numpy": "numpy",
        "Pillow": "PIL",
        "PyMuPDF": "fitz",
        "omml2latex": "omml2latex",
    }
    module_state = {
        label: bool(importlib.util.find_spec(import_name))
        for label, import_name in modules.items()
    }
    missing = [name for name, ok in module_state.items() if not ok]
    candidates = _axmath_template_candidates()
    template = next((p for p in candidates if p.is_file()), None)
    powershell = shutil.which("pwsh") or shutil.which("powershell")
    return {
        "python": sys.executable,
        "python_modules": module_state,
        "missing_python_modules": missing,
        "powershell": powershell,
        "axmath_template": str(template) if template else None,
        "axmath_template_candidates": [str(x) for x in candidates],
        "ready_for_end_to_end": bool(not missing and powershell and template),
    }


def discover_docs(target: Path) -> list[Path]:
    if target.is_file():
        return [target]
    if target.is_dir():
        return sorted(
            p.resolve()
            for p in target.glob("*.docx")
            if not p.name.startswith("~$")
        )
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
        structure = stats.get("source_structure") or {}
        row.update({
            "sha256": sha256_file(path),
            "paragraphs": stats["paragraphs"],
            "omath": stats["omath"],
            "axmath": stats["axmath_ole"],
            "multi_sibling_group_count": structure.get("multi_sibling_group_count", 0),
            "multi_sibling_extra_nodes": structure.get("multi_sibling_extra_nodes", 0),
            "batch_collapse_signature_axmath_count": structure.get(
                "batch_collapse_signature_axmath_count"
            ),
        })
    except Exception as exc:
        row["state"] = "invalid_docx"
        row["error"] = str(exc)
        return row

    if row["omath"] > 0 and row["axmath"] == 0:
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
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def run_conversion(source: Path, doc_dir: Path, *, resume: bool = False, numbering_scope: str = "line-start") -> dict:
    source_sha_before = sha256_file(source)
    source_dir = doc_dir / "source"
    working_dir = doc_dir / "working"
    logs_dir = doc_dir / "logs"
    source_dir.mkdir(parents=True, exist_ok=True)
    working_dir.mkdir(parents=True, exist_ok=True)
    # Logs directory is created only if an actual conversion failure occurs.

    frozen_source = source_dir / source.name
    if resume:
        if not frozen_source.is_file():
            raise FileNotFoundError(f"Resume frozen source is missing: {frozen_source}")
    else:
        shutil.copy2(source, frozen_source)
    frozen_source_sha256 = sha256_file(frozen_source)
    if frozen_source_sha256 != source_sha_before:
        raise RuntimeError("Frozen source copy hash mismatch; refusing conversion.")
    conversion_working = working_dir / f"{source.stem}_AxMath-working.docx"
    log_path = logs_dir / (
        f"conversion.resume-{dt.datetime.now():%Y%m%d-%H%M%S}.log.json"
        if resume else "conversion.log.json"
    )

    cmd = [
        sys.executable,
        str(SCRIPTS_DIR / "run_skill.py"),
        "--input", str(frozen_source),
        "--output", str(conversion_working),
        "--quiet",
    ]
    if resume:
        cmd.append("--resume")
        normalized = working_dir / f"{source.stem}_AxMath-working_numbering-normalized.docx"
        if normalized.exists():
            raise FileExistsError(
                f"Resume cannot overwrite an existing normalized working copy: {normalized}"
            )
    started = time.time()
    proc = subprocess.run(
        cmd, cwd=str(SCRIPTS_DIR), text=True, encoding="utf-8",
        errors="replace", capture_output=True,
    )
    # On success run_skill.py already persisted the full performance/audit
    # report. Only create a separate wrapper log on failure.
    if proc.returncode:
        write_json(log_path, {
            "returncode": proc.returncode,
            "seconds": time.time() - started,
            "stdout_tail": proc.stdout[-12000:],
            "stderr_tail": proc.stderr[-12000:],
        })

    skill_report_path = Path(str(conversion_working) + ".skill-report.json")
    skill_report = None
    if skill_report_path.exists():
        try:
            skill_report = json.loads(skill_report_path.read_text(encoding="utf-8-sig"))
        except (ValueError, OSError):
            pass

    # The child already performed the expensive full source/candidate XML
    # comparison. Reuse that audit on the unchanged default path instead of
    # decompressing the same large OLE-containing DOCX multiple times.
    raw_audit = None
    if (
        proc.returncode == 0
        and isinstance(skill_report, dict)
        and skill_report.get("source_unchanged")
        and skill_report.get("conversion_report_fresh")
        and skill_report.get("returncode") == 0
    ):
        raw_audit = skill_report.get("audit")
    source_stats = raw_audit["source"] if isinstance(raw_audit, dict) else analyze(frozen_source)

    if numbering_scope == "line-start" and raw_audit:
        expected_replacements = int(raw_audit.get("numbering_punctuation_expected_change_count") or 0)
    else:
        if "plain_contract" not in source_stats:
            source_stats = analyze(frozen_source)
        expected_replacements = sum(
            len(numbering_dunhao_positions(text, scope=numbering_scope))
            for text in source_stats["plain_contract"]
        )

    source_sha_after = sha256_file(source)
    source_unchanged = source_sha_after == source_sha_before
    working = conversion_working
    normalization = None
    if proc.returncode == 0 and conversion_working.exists():
        if expected_replacements:
            working = working_dir / f"{source.stem}_AxMath-working_numbering-normalized.docx"
            normalization = normalize_docx(conversion_working, working, scope=numbering_scope)
        else:
            normalization = {
                "schema": "axmath-numbering-punctuation-normalization/v1",
                "scope": numbering_scope,
                "replacements": 0,
                "skipped_no_matching_labels": True,
                "input": str(conversion_working),
                "output": str(conversion_working),
            }

    if raw_audit and working == conversion_working:
        candidate_stats = raw_audit["candidate"]
        audit = raw_audit
    else:
        candidate_stats = analyze(working) if working.exists() else None
        if candidate_stats:
            if "plain_contract" not in source_stats:
                source_stats = analyze(frozen_source)
            audit = compare_docx(
                frozen_source, working, numbering_scope=numbering_scope,
                source_analysis=source_stats, candidate_analysis=candidate_stats,
            )
        else:
            audit = None

    performance = (
        (skill_report or {}).get("performance")
        if isinstance(skill_report, dict) else None
    )
    count_state = (audit or {}).get("formula_count_state") or {}
    count_name = count_state.get("state")
    count_acceptable = count_name in {
        "exact",
        "known_multisibling_collapse",
    }
    repair_required = count_name == "known_multisibling_collapse"

    success = bool(
        proc.returncode == 0
        and candidate_stats
        and candidate_stats["omath"] == 0
        and count_acceptable
        and audit
        and audit.get("paragraph_count_equal")
        and audit.get("nonmath_text_contract_exact")
        and source_unchanged
    )
    return {
        "status": (
            "converted_ready_for_gpt_review"
            if success
            else "conversion_failed"
        ),
        "repair_required": repair_required,
        "next_repair_class": (
            "M1_MULTISIBLING_OMATHPARA" if repair_required else None
        ),
        "formula_count_state": count_state,
        "original_source": str(source),
        "original_source_sha256": source_sha_before,
        "original_source_sha256_after": source_sha_after,
        "source_unchanged": source_unchanged,
        "resumed": resume,
        "numbering_scope": numbering_scope,
        "frozen_source": str(frozen_source),
        "frozen_source_sha256": frozen_source_sha256,
        "conversion_working_docx": (
            str(conversion_working) if conversion_working.exists() else None
        ),
        "working_docx": str(working) if working.exists() else None,
        "numbering_punctuation_normalization": normalization,
        "source_omath": source_stats["omath"],
        "working_omath": (
            candidate_stats["omath"] if candidate_stats else None
        ),
        "working_axmath": (
            candidate_stats["axmath_ole"] if candidate_stats else None
        ),
        "paragraph_count_equal": bool(
            audit and audit.get("paragraph_count_equal")
        ),
        "nonmath_text_exact": bool(
            audit and audit.get("nonmath_text_exact")
        ),
        "nonmath_text_contract_exact": bool(
            audit and audit.get("nonmath_text_contract_exact")
        ),
        "performance": performance,
        "conversion_log": str(log_path) if proc.returncode else None,
        "mandatory_next_step": (
            "Run diagnose_after_conversion.py on frozen_source + working_docx "
            "before any repair or page/layout tuning."
        ),
    }


def main() -> int:
    _utf8_console()
    ap = argparse.ArgumentParser(
        description=(
            "Preflight -> adjacent workspace -> OfficeMath to AxMath -> "
            "state-machine diagnosis handoff"
        )
    )
    group = ap.add_mutually_exclusive_group(required=True)
    group.add_argument("--input", help="DOCX file or directory")
    group.add_argument(
        "--input-dir", help="Compatibility alias for a directory"
    )
    ap.add_argument("--inspect-only", action="store_true")
    ap.add_argument("--resume-run", help="Existing run-YYYYMMDD-HHMMSS workspace to resume, never a new conversion")
    ap.add_argument("--numbering-scope", choices=("line-start", "anywhere"), default="line-start")
    args = ap.parse_args()

    target = Path(args.input or args.input_dir).resolve()
    if args.inspect_only and args.resume_run:
        ap.error("--inspect-only and --resume-run cannot be combined")
    docs = discover_docs(target)
    states = [inspect_source(p) for p in docs]
    environment = environment_preflight()
    preflight = {
        "schema": "axmath-preflight/v2",
        "target": str(target),
        "checked_at": dt.datetime.now().isoformat(),
        "word_pids": word_pids(),
        "environment": environment,
        "documents": states,
    }

    if args.inspect_only:
        print(json.dumps(preflight, ensure_ascii=False, indent=2))
        return 0

    if not environment["ready_for_end_to_end"]:
        preflight["status"] = "blocked_missing_environment_dependency"
        print(json.dumps(preflight, ensure_ascii=False, indent=2))
        return 1

    ready = [
        Path(x["source"])
        for x in states
        if x.get("state") == "ready_officemath"
    ]
    if not ready:
        preflight["status"] = "nothing_to_convert"
        print(json.dumps(preflight, ensure_ascii=False, indent=2))
        return 0

    if args.resume_run:
        if len(ready) != 1 or not target.is_file():
            ap.error("--resume-run requires a single source DOCX (not a directory)")
        run_dir = Path(args.resume_run).resolve()
        saved_preflight_file = run_dir / "SOURCE_STATE.json"
        if not saved_preflight_file.is_file():
            raise FileNotFoundError(f"Resume run has no original SOURCE_STATE.json: {run_dir}")
        saved = json.loads(saved_preflight_file.read_text(encoding="utf-8-sig"))
        old_docs = saved.get("documents") or []
        if (
            len(old_docs) != 1
            or Path(old_docs[0].get("source") or "").resolve() != target
            or old_docs[0].get("sha256") != sha256_file(target)
        ):
            raise RuntimeError("Resume refused: original source path or hash differs from this run.")
    else:
        run_dir = unique_run_dir(workspace_root_for(target))
        run_dir.mkdir(parents=True)
        write_json(run_dir / "SOURCE_STATE.json", preflight)

    results = []
    for source in ready:
        result = run_conversion(
            source, run_dir / source.stem,
            resume=bool(args.resume_run), numbering_scope=args.numbering_scope,
        )
        results.append(result)

    failed = [
        x
        for x in results
        if x["status"] != "converted_ready_for_gpt_review"
    ]
    succeeded = [
        x
        for x in results
        if x["status"] == "converted_ready_for_gpt_review"
    ]
    manifest = {
        "schema": "axmath-gpt-review-handoff/v2",
        "status": (
            "ready"
            if succeeded and not failed
            else ("partial" if succeeded else "blocked")
        ),
        "created_at": dt.datetime.now().isoformat(),
        "skill": str(SKILL_PATH),
        "workspace": str(run_dir),
        "numbering_scope": args.numbering_scope,
        "resumed": bool(args.resume_run),
        "documents": results,
        "mandatory_entrypoint": "scripts/diagnose_after_conversion.py",
        "review_rule": (
            "MANDATORY: read SKILL.md first, then run diagnose_after_conversion.py "
            "on the frozen source and working DOCX. Do not improvise repairs, do "
            "not chase page count, do not run width/height/w:position sweeps, and "
            "do not run geometry repair while an M1 structural count state remains."
        ),
    }
    write_json(run_dir / "READY_FOR_GPT_REVIEW.json", manifest)

    final = {
        "status": (
            "ready_for_gpt_review"
            if succeeded and not failed
            else ("partial" if succeeded else "failed")
        ),
        "workspace": str(run_dir),
        "ready_manifest": str(run_dir / "READY_FOR_GPT_REVIEW.json"),
        "documents": results,
    }
    print(json.dumps(final, ensure_ascii=False, indent=2))
    return 0 if succeeded and not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
