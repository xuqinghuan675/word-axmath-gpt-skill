from __future__ import annotations

import argparse
import json
from pathlib import Path

from PIL import Image, ImageDraw

from snapshot_docx import snapshot


CENTER_ALIGNMENT = 1  # Word WdParagraphAlignment.wdAlignParagraphCenter


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


def _build_side_by_side(source_pages: list[dict], final_pages: list[dict], outdir: Path) -> list[str]:
    outdir.mkdir(parents=True, exist_ok=True)
    source_map = {int(p["page"]): p for p in source_pages}
    final_map = {int(p["page"]): p for p in final_pages}
    page_numbers = sorted(set(source_map) | set(final_map))
    outputs = []

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
        outputs.append(str(out))
    return outputs


def compare(source_docx: Path, final_docx: Path, outdir: Path, center_tolerance_pt: float = 3.0) -> dict:
    outdir.mkdir(parents=True, exist_ok=True)
    source_snapshot = snapshot(source_docx, outdir / "source", "source")
    final_snapshot = snapshot(final_docx, outdir / "final", "final")

    source_formulas = source_snapshot.get("inventory", {}).get("omath", [])
    if not source_formulas:
        source_formulas = source_snapshot.get("inventory", {}).get("axmath", [])
    final_formulas = final_snapshot.get("inventory", {}).get("axmath", [])

    report = {
        "source": str(source_docx.resolve()),
        "final": str(final_docx.resolve()),
        "source_pages": len(source_snapshot.get("pages", [])),
        "final_pages": len(final_snapshot.get("pages", [])),
        "source_formula_count": len(source_formulas),
        "final_axmath_count": len(final_formulas),
        "page_count_equal": len(source_snapshot.get("pages", [])) == len(final_snapshot.get("pages", [])),
        "formula_count_equal": len(source_formulas) == len(final_formulas),
        "alignment_breaks": [],
        "center_alignment_breaks": [],
        "center_position_breaks": [],
        "center_tolerance_pt": center_tolerance_pt,
        "visual_review_required": True,
        "visual_review_passed": None,
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

    report["side_by_side_pages"] = _build_side_by_side(
        source_snapshot.get("pages", []),
        final_snapshot.get("pages", []),
        outdir / "side_by_side",
    )

    report["structural_pass"] = bool(
        report["page_count_equal"]
        and report["formula_count_equal"]
        and not report["center_alignment_breaks"]
        and not report["center_position_breaks"]
    )
    report["acceptance_instruction"] = (
        "GPT MUST inspect every source-vs-final page image. Structural pass alone is not acceptance. "
        "Reject completion for any unexplained visual difference in formula size, baseline, wrapping, "
        "line breaks, page breaks, paragraph spacing, indentation, or centering. Any formula centered "
        "in the source must remain centered in the final document."
    )

    out = outdir / "STRICT_FINAL_COMPARE.json"
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--final", required=True)
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--center-tolerance-pt", type=float, default=3.0)
    args = ap.parse_args()

    report = compare(
        Path(args.source).resolve(),
        Path(args.final).resolve(),
        Path(args.outdir).resolve(),
        args.center_tolerance_pt,
    )
    print(json.dumps({
        "page_count_equal": report["page_count_equal"],
        "formula_count_equal": report["formula_count_equal"],
        "center_alignment_breaks": len(report["center_alignment_breaks"]),
        "center_position_breaks": len(report["center_position_breaks"]),
        "structural_pass": report["structural_pass"],
        "visual_review_required": report["visual_review_required"],
        "side_by_side_pages": len(report["side_by_side_pages"]),
        "report": str((Path(args.outdir).resolve() / "STRICT_FINAL_COMPARE.json")),
    }, ensure_ascii=True, indent=2))
    return 0 if report["structural_pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
