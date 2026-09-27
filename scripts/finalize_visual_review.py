from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

VISUAL_REVIEW_SCHEMA = "axmath-visual-review/v1"


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def finalize(report_path: Path, review_path: Path) -> dict:
    report_path = report_path.resolve()
    review_path = review_path.resolve()
    report = json.loads(report_path.read_text(encoding="utf-8-sig"))
    review = json.loads(review_path.read_text(encoding="utf-8-sig"))
    errors: list[str] = []

    if report.get("schema") != "axmath-strict-final-compare/v2":
        errors.append("unsupported compare report schema")
    if review.get("schema") != VISUAL_REVIEW_SCHEMA:
        errors.append("unsupported visual review schema")

    source = Path(report.get("source") or "")
    final = Path(report.get("final") or "")
    if not source.is_file():
        errors.append("source file missing")
    elif _sha256_file(source) != report.get("source_sha256"):
        errors.append("source file changed after strict compare")
    if not final.is_file():
        errors.append("final file missing")
    elif _sha256_file(final) != report.get("final_sha256"):
        errors.append("final file changed after strict compare")

    if review.get("source_sha256") != report.get("source_sha256"):
        errors.append("review source hash does not match compare report")
    if review.get("final_sha256") != report.get("final_sha256"):
        errors.append("review final hash does not match compare report")

    inventory = report.get("side_by_side_inventory") or []
    expected = {int(x["page"]): x for x in inventory}
    if not expected:
        errors.append("no side-by-side page evidence in compare report")
    if not report.get("visual_evidence_complete"):
        errors.append("compare report visual evidence is incomplete")
    reviewed_rows = review.get("pages") or []
    reviewed: dict[int, dict] = {}
    for row in reviewed_rows:
        try:
            page = int(row.get("page"))
        except Exception:
            errors.append("review contains invalid page number")
            continue
        if page in reviewed:
            errors.append(f"duplicate review page {page}")
            continue
        reviewed[page] = row

    if set(reviewed) != set(expected):
        errors.append(
            f"review page set mismatch: expected={sorted(expected)} reviewed={sorted(reviewed)}"
        )

    for page, expected_row in expected.items():
        row = reviewed.get(page)
        image = Path(expected_row.get("path") or "")
        if not image.is_file():
            errors.append(f"review image missing for page {page}")
            continue
        actual_image_hash = _sha256_file(image)
        if actual_image_hash != expected_row.get("sha256"):
            errors.append(f"review image changed for page {page}")
        if row is None:
            continue
        if row.get("image_sha256") != expected_row.get("sha256"):
            errors.append(f"review hash mismatch for page {page}")
        if row.get("passed") is not True:
            errors.append(f"page {page} was not explicitly marked passed")

    if review.get("overall_passed") is not True:
        errors.append("overall_passed is not true")
    if not report.get("hard_content_pass"):
        errors.append("hard content gate did not pass")

    passed = not errors
    report["visual_review_manifest"] = str(review_path)
    report["visual_review_passed"] = passed
    report["visual_review_errors"] = errors
    report["acceptance_pass"] = passed
    report["acceptance_status"] = "passed" if passed else "failed"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--report", required=True)
    ap.add_argument("--review", required=True)
    args = ap.parse_args()

    report = finalize(Path(args.report), Path(args.review))
    print(json.dumps({
        "acceptance_pass": report["acceptance_pass"],
        "acceptance_status": report["acceptance_status"],
        "visual_review_passed": report["visual_review_passed"],
        "visual_review_errors": report.get("visual_review_errors", []),
        "report": str(Path(args.report).resolve()),
    }, ensure_ascii=True, indent=2))
    return 0 if report["acceptance_pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
