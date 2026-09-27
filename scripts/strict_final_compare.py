from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

from PIL import Image, ImageDraw

from audit_docx import compare as audit_docx_compare
from snapshot_docx import snapshot


CENTER_ALIGNMENT = 1  # Word WdParagraphAlignment.wdAlignParagraphCenter
VISUAL_REVIEW_SCHEMA = "axmath-visual-review/v1"


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _formula_center_x(rec: dict, source: bool) -> float | None:
    x = rec.get("x_pt")
    width = rec.get("visual_width_pt") if source else rec.get("width_pt")
    if x is None or width is None:
        return None
    try:
        x = float(x)
        width = float(width)
    except (TypeError, ValueError):
        return None
    if x < 0 or width <= 0:
        return None
    return x + width / 2.0


def _build_side_by_side(source_pages: list[dict], final_pages: list[dict], outdir: Path) -> list[dict]:
    outdir.mkdir(parents=True, exist_ok=True)
    source_map = {int(p["page"]): p for p in source_pages}
    final_map = {int(p["page"]): p for p in final_pages}
    page_numbers = sorted(set(source_map) | set(final_map))
    outputs: list[dict] = []

    for page_no in page_numbers:
        src_meta = source_map.get(page_no)
        fin_meta = final_map.get(page_no)
        src = Image.open(src_meta["path"]).convert("RGB") if src_meta else None
        fin = Image.open(fin_meta["path"]).convert("RGB") if fin_meta else None

        widths = [im.width for im in (src, fin) if im is not None]
        heights = [im.height for im in (src, fin) if im is not None]
        if not widths or not heights:
            continue
        panel_w = max(widths)
        panel_h = max(heights)
        header_h = 36
        sheet = Image.new("RGB", (panel_w * 2 + 24, panel_h + header_h), "white")
        draw = ImageDraw.Draw(sheet)
        draw.text((8, 10), f"SOURCE page {page_no}", fill="black")
        draw.text((panel_w + 20, 10), f"FINAL page {page_no}", fill="black")
        if src is not None:
            sheet.paste(src, (0, header_h))
        if fin is not None:
            sheet.paste(fin, (panel_w + 24, header_h))

        out = outdir / f"page-{page_no:03d}-source-vs-final.png"
        sheet.save(out)
        outputs.append({
            "page": page_no,
            "path": str(out.resolve()),
            "sha256": _sha256_file(out),
            "source_present": src is not None,
            "final_present": fin is not None,
        })
    return outputs


def _write_visual_review_template(report: dict, outdir: Path) -> Path:
    template = {
        "schema": VISUAL_REVIEW_SCHEMA,
        "source_sha256": report["source_sha256"],
        "final_sha256": report["final_sha256"],
        "overall_passed": None,
        "reviewer": "",
        "notes": "",
        "pages": [
            {
                "page": row["page"],
                "image": row["path"],
                "image_sha256": row["sha256"],
                "passed": None,
                "notes": "",
            }
            for row in report["side_by_side_inventory"]
        ],
    }
    path = outdir / "VISUAL_REVIEW_TEMPLATE.json"
    path.write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def compare(
    source_docx: Path,
    final_docx: Path,
    outdir: Path,
    expected_source_sha256: str,
    center_tolerance_pt: float = 3.0,
) -> dict:
    source_docx = source_docx.resolve()
    final_docx = final_docx.resolve()
    outdir = outdir.resolve()

    if source_docx == final_docx:
        raise ValueError("Source and final DOCX must be different files.")
    if not source_docx.is_file():
        raise FileNotFoundError(source_docx)
    if not final_docx.is_file():
        raise FileNotFoundError(final_docx)

    outdir.mkdir(parents=True, exist_ok=True)
    source_sha_before = _sha256_file(source_docx)
    final_sha_before = _sha256_file(final_docx)

    source_snapshot = snapshot(source_docx, outdir / "source", "source")
    final_snapshot = snapshot(final_docx, outdir / "final", "final")
    content_audit = audit_docx_compare(source_docx, final_docx)

    source_sha_after = _sha256_file(source_docx)
    final_sha_after = _sha256_file(final_docx)
    source_stable = source_sha_before == source_sha_after
    final_stable = final_sha_before == final_sha_after
    source_baseline_ok = source_sha_before.lower() == expected_source_sha256.strip().lower()

    source_formulas = source_snapshot.get("inventory", {}).get("omath", [])
    if not source_formulas:
        source_formulas = source_snapshot.get("inventory", {}).get("axmath", [])
    final_formulas = final_snapshot.get("inventory", {}).get("axmath", [])
    final_omath = final_snapshot.get("inventory", {}).get("omath", [])

    report = {
        "schema": "axmath-strict-final-compare/v2",
        "source": str(source_docx),
        "final": str(final_docx),
        "source_sha256": source_sha_before,
        "final_sha256": final_sha_before,
        "expected_source_sha256": expected_source_sha256.strip().lower(),
        "source_baseline_ok": source_baseline_ok,
        "source_stable_during_compare": source_stable and bool(source_snapshot.get("docx_stable")),
        "final_stable_during_compare": final_stable and bool(final_snapshot.get("docx_stable")),
        "source_pages": len(source_snapshot.get("pages", [])),
        "final_pages": len(final_snapshot.get("pages", [])),
        "source_formula_count": len(source_formulas),
        "final_axmath_count": len(final_formulas),
        "final_omath_count": len(final_omath),
        "page_count_equal": len(source_snapshot.get("pages", [])) == len(final_snapshot.get("pages", [])),
        "formula_count_equal": len(source_formulas) == len(final_formulas),
        "paragraph_count_equal": bool(content_audit.get("paragraph_count_equal")),
        "nonmath_text_exact": bool(content_audit.get("nonmath_text_exact")),
        "nonmath_text_diff_count": int(content_audit.get("nonmath_text_diff_count") or 0),
        "alignment_breaks": [],
        "center_alignment_breaks": [],
        "center_position_breaks": [],
        "center_tolerance_pt": center_tolerance_pt,
        "visual_review_required": True,
        "visual_review_passed": None,
        "acceptance_pass": False,
    }

    for src, fin in zip(source_formulas, final_formulas):
        ordinal = int(src.get("ordinal") or 0)
        source_fmt = src.get("paragraph_format") or {}
        final_fmt = fin.get("paragraph_format") or {}
        source_alignment = source_fmt.get("alignment")
        final_alignment = final_fmt.get("alignment")

        if (
            source_alignment is not None
            and final_alignment is not None
            and int(source_alignment) != int(final_alignment)
        ):
            rec = {
                "ordinal": ordinal,
                "source_page": src.get("page"),
                "final_page": fin.get("page"),
                "source_alignment": source_alignment,
                "final_alignment": final_alignment,
            }
            report["alignment_breaks"].append(rec)
            if int(source_alignment) == CENTER_ALIGNMENT:
                report["center_alignment_breaks"].append(rec)

        if source_alignment is not None and int(source_alignment) == CENTER_ALIGNMENT:
            src_center = _formula_center_x(src, True)
            fin_center = _formula_center_x(fin, False)
            if src_center is not None and fin_center is not None:
                delta = abs(fin_center - src_center)
                if delta > center_tolerance_pt:
                    report["center_position_breaks"].append({
                        "ordinal": ordinal,
                        "source_page": src.get("page"),
                        "final_page": fin.get("page"),
                        "source_center_x_pt": src_center,
                        "final_center_x_pt": fin_center,
                        "delta_pt": delta,
                    })

    report["side_by_side_inventory"] = _build_side_by_side(
        source_snapshot.get("pages", []),
        final_snapshot.get("pages", []),
        outdir / "side_by_side",
    )
    report["side_by_side_pages"] = [x["path"] for x in report["side_by_side_inventory"]]

    report["structural_pass"] = bool(
        report["page_count_equal"]
        and report["formula_count_equal"]
        and not report["center_alignment_breaks"]
        and not report["center_position_breaks"]
    )
    report["hard_content_pass"] = bool(
        report["source_baseline_ok"]
        and report["source_stable_during_compare"]
        and report["final_stable_during_compare"]
        and report["formula_count_equal"]
        and report["final_omath_count"] == 0
        and report["paragraph_count_equal"]
        and report["nonmath_text_exact"]
    )
    report["acceptance_status"] = (
        "awaiting_visual_review" if report["hard_content_pass"] else "hard_content_failed"
    )
    report["acceptance_instruction"] = (
        "Inspect every side-by-side page image. Fill VISUAL_REVIEW_TEMPLATE.json only from direct "
        "visual inspection, save it as a review manifest, then run finalize_visual_review.py. "
        "Geometry/centering diagnostics are triage warnings; hard content gates and per-page visual "
        "review are the acceptance authority."
    )

    template_path = _write_visual_review_template(report, outdir)
    report["visual_review_template"] = str(template_path.resolve())

    out = outdir / "STRICT_FINAL_COMPARE.json"
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--final", required=True)
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--expected-source-sha256", required=True)
    ap.add_argument("--center-tolerance-pt", type=float, default=3.0)
    args = ap.parse_args()

    report = compare(
        Path(args.source),
        Path(args.final),
        Path(args.outdir),
        args.expected_source_sha256,
        args.center_tolerance_pt,
    )
    print(json.dumps({
        "hard_content_pass": report["hard_content_pass"],
        "page_count_equal": report["page_count_equal"],
        "formula_count_equal": report["formula_count_equal"],
        "final_omath_count": report["final_omath_count"],
        "paragraph_count_equal": report["paragraph_count_equal"],
        "nonmath_text_exact": report["nonmath_text_exact"],
        "center_alignment_breaks": len(report["center_alignment_breaks"]),
        "center_position_breaks": len(report["center_position_breaks"]),
        "structural_pass": report["structural_pass"],
        "visual_review_required": report["visual_review_required"],
        "side_by_side_pages": len(report["side_by_side_pages"]),
        "visual_review_template": report["visual_review_template"],
        "report": str((Path(args.outdir).resolve() / "STRICT_FINAL_COMPARE.json")),
    }, ensure_ascii=True, indent=2))
    # Never return success before the required per-page visual review is finalized.
    return 2 if report["hard_content_pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
