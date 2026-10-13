#!/usr/bin/env python3
"""Local, connector-free CLI for the existing Word -> AxMath skill.

This is a thin launcher, not a replacement converter, validator, or agent.
Word COM / AxMath execute in the *local* Windows user session.
"""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
SCRIPTS = ROOT / "scripts"

REPORT_NAMES = (
    "SOURCE_STATE.json",
    "READY_FOR_GPT_REVIEW.json",
    "NEXT_ACTION.json",
    "STRICT_FINAL_COMPARE.json",
)
STATUS_FIELDS = (
    "schema",
    "status",
    "state",
    "source",
    "frozen_source",
    "working_docx",
    "working",
    "final",
    "source_sha256",
    "final_sha256",
    "acceptance_pass",
    "visual_review_passed",
    "visual_review_required",
    "next_actions",
)


def local_path(value: str) -> str:
    """Resolve user paths *before* switching subprocess cwd to the skill root."""
    return str(Path(value).expanduser().absolute())


def add_numbering_option(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--numbering-scope", choices=("line-start", "anywhere"),
        default="line-start",
        help="Keep line-start unless the user explicitly approved anywhere.",
    )


def make_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run the existing AxMath skill directly from local Codex or a terminal; no ChatGPT PC connector."
    )
    commands = parser.add_subparsers(dest="command", required=True)

    inspect = commands.add_parser("inspect", help="Read-only DOCX and Word/AxMath preflight.")
    inspect.add_argument("--input", required=True, help="Input DOCX.")
    add_numbering_option(inspect)

    convert = commands.add_parser("convert", help="Freeze the source, then run the original official conversion.")
    convert.add_argument("--input", required=True, help="Original DOCX.")
    add_numbering_option(convert)

    cache_analyze = commands.add_parser("cache-analyze", help="Read-only count of exact repeated OfficeMath formulas.")
    cache_analyze.add_argument("--input", required=True)

    cache_prepare = commands.add_parser("cache-prepare", help="Freeze and prepare repeated formulas for one official conversion each.")
    cache_prepare.add_argument("--input", required=True)
    cache_prepare.add_argument("--outdir", required=True)
    cache_prepare.add_argument("--min-duplicates", type=int, default=64)

    cache_restore = commands.add_parser("cache-restore", help="Restore genuine AxMath OLE after official conversion of the prepared DOCX.")
    cache_restore.add_argument("--manifest", required=True)
    cache_restore.add_argument("--ready", required=True)
    cache_restore.add_argument("--output", required=True)
    cache_restore.add_argument("--raw", action="store_true", help="Benchmark only: use the unnormalized conversion output.")
    add_numbering_option(cache_restore)

    resume = commands.add_parser("resume", help="Resume the original checkpoint; NEVER reconvert from scratch.")
    resume.add_argument("--input", required=True, help="Original DOCX.")
    resume.add_argument("--resume-run", required=True, help="Existing run directory.")
    add_numbering_option(resume)

    diagnose = commands.add_parser("diagnose", help="Run the authoritative post-conversion dispatcher.")
    diagnose.add_argument("--source", required=True, help="Frozen source from READY_FOR_GPT_REVIEW.json.")
    diagnose.add_argument("--working", required=True, help="Working DOCX from READY_FOR_GPT_REVIEW.json.")
    diagnose.add_argument("--outdir", required=True)
    diagnose.add_argument("--deep-geometry", action="store_true")
    diagnose.add_argument("--prime-review", help="Already verified hash-bound prime review manifest.")
    add_numbering_option(diagnose)

    compare = commands.add_parser("compare", help="Generate final source-vs-AxMath full-page visual review.")
    compare.add_argument("--source", required=True, help="Frozen source DOCX.")
    compare.add_argument("--final", required=True, help="Final candidate DOCX, not the source.")
    compare.add_argument("--outdir", required=True)
    compare.add_argument("--expected-source-sha256", required=True)
    compare.add_argument("--repair-ledger", help="Validated ledger for explicitly approved M2-B line splitting.")
    compare.add_argument("--center-tolerance-pt", type=float, help="Normally do not override the original default.")
    add_numbering_option(compare)

    finalize = commands.add_parser("finalize", help="Validate manually reviewed pages; does not auto-approve.")
    finalize.add_argument("--report", required=True, help="STRICT_FINAL_COMPARE.json.")
    finalize.add_argument("--review", required=True, help="Human/agent-reviewed VISUAL_REVIEW.json.")

    status = commands.add_parser("status", help="Read existing JSON artifacts; NEVER infer visual acceptance.")
    status.add_argument("--path", required=True, help="A report JSON or the directory containing reports.")

    return parser


def command_for(args: argparse.Namespace) -> list[str]:
    """Translate only declared commands to the existing first-party entrypoints."""
    if args.command in ("inspect", "convert", "resume"):
        params = ["--input", local_path(args.input)]
        if args.command == "inspect":
            params.append("--inspect-only")
        if args.command == "resume":
            params.extend(["--resume-run", local_path(args.resume_run)])
        params.extend(["--numbering-scope", args.numbering_scope])
        script = "one_click_convert.py"
    elif args.command == "cache-analyze":
        params = ["analyze", "--input", local_path(args.input)]
        script = "exact_duplicate_cache.py"
    elif args.command == "cache-prepare":
        params = ["prepare", "--input", local_path(args.input), "--outdir", local_path(args.outdir), "--min-duplicates", str(args.min_duplicates)]
        script = "exact_duplicate_cache.py"
    elif args.command == "cache-restore":
        params = ["restore", "--manifest", local_path(args.manifest), "--ready", local_path(args.ready), "--output", local_path(args.output), "--numbering-scope", args.numbering_scope]
        if args.raw:
            params.append("--raw")
        script = "exact_duplicate_cache.py"
    elif args.command == "diagnose":
        params = [
            "--source", local_path(args.source),
            "--working", local_path(args.working),
            "--outdir", local_path(args.outdir),
            "--numbering-scope", args.numbering_scope,
        ]
        if args.deep_geometry:
            params.append("--deep-geometry")
        if args.prime_review:
            params.extend(["--prime-review", local_path(args.prime_review)])
        script = "diagnose_after_conversion.py"
    elif args.command == "compare":
        params = [
            "--source", local_path(args.source),
            "--final", local_path(args.final),
            "--outdir", local_path(args.outdir),
            "--expected-source-sha256", args.expected_source_sha256,
            "--numbering-scope", args.numbering_scope,
        ]
        if args.repair_ledger:
            params.extend(["--repair-ledger", local_path(args.repair_ledger)])
        if args.center_tolerance_pt is not None:
            params.extend(["--center-tolerance-pt", str(args.center_tolerance_pt)])
        script = "strict_final_compare.py"
    elif args.command == "finalize":
        params = ["--report", local_path(args.report), "--review", local_path(args.review)]
        script = "finalize_visual_review.py"
    else:
        raise ValueError(f"Unknown launch command: {args.command}")
    return [sys.executable, str(SCRIPTS / script), *params]


def read_status(path_arg: str) -> int:
    target = Path(local_path(path_arg))
    if target.is_file():
        reports = [target]
    elif target.is_dir():
        reports = [target / name for name in REPORT_NAMES if (target / name).is_file()]
    else:
        print(f"No report file or directory: {target}", file=sys.stderr)
        return 2
    if not reports:
        print(f"No supported AxMath report JSON in: {target}", file=sys.stderr)
        return 2

    output = []
    try:
        for report in reports:
            payload = json.loads(report.read_text(encoding="utf-8-sig"))
            if not isinstance(payload, dict):
                raise ValueError(f"Not a JSON object: {report}")
            summary = {
                "report": str(report),
                **{key: payload[key] for key in STATUS_FIELDS if key in payload},
            }
            if report.name == "STRICT_FINAL_COMPARE.json" and payload.get("acceptance_pass") is not True:
                summary["acceptance_note"] = (
                    "NOT ACCEPTED: a generated visual comparison is not evidence that every page was reviewed."
                )
            output.append(summary)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"Cannot read AxMath report: {exc}", file=sys.stderr)
        return 2

    print(json.dumps({"read_only": True, "reports": output}, ensure_ascii=False, indent=2))
    return 0


def main(argv: list[str] | None = None) -> int:
    args = make_parser().parse_args(argv)
    if args.command == "status":
        return read_status(args.path)

    command = command_for(args)
    if not Path(command[1]).is_file():
        print(f"Missing official skill entrypoint: {command[1]}", file=sys.stderr)
        return 2

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    try:
        # Inherit stdout/stderr, exit status and long-job visibility. No plugin,
        # lease, hidden daemon, second conversion engine or auto-approval.
        return subprocess.run(command, cwd=str(ROOT), env=env, check=False).returncode
    except OSError as exc:
        print(f"Failed to launch the local AxMath script: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
