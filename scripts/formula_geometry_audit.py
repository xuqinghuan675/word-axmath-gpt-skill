from __future__ import annotations

import argparse
import hashlib
import json
import math
import statistics
from pathlib import Path

from normalize_axmath_tex import contains_prime_or_derivative_marker
from snapshot_docx import _axmath_layout_inventory, _collect_com_inventory
from word_runtime import OwnedWord


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _median(values):
    return statistics.median(values) if values else None


def _mad(values, center):
    if not values or center is None:
        return None
    return statistics.median(abs(v - center) for v in values)


def _source_same_line_groups(rows, tolerance_pt: float = 1.5):
    by_para = {}
    for row in rows:
        p = row.get("paragraph_start")
        if p is None or row.get("page") is None or row.get("y_pt") is None:
            continue
        by_para.setdefault(int(p), []).append(row)

    groups = []
    for paragraph_start, members in by_para.items():
        members = sorted(members, key=lambda r: int(r.get("ordinal") or 0))
        current = []
        for row in members:
            # A source OfficeMath range may start on one visual line and end on
            # the next. Such a formula is internally multiline, so formulas
            # sharing only its start-y are not a hard same-line constraint for
            # an indivisible AxMath OLE object.
            start_y = row.get("y_pt")
            end_y = row.get("end_y_pt")
            source_multiline = bool(
                start_y is not None
                and end_y is not None
                and abs(float(end_y) - float(start_y)) > tolerance_pt
            )
            if source_multiline:
                if len(current) >= 2:
                    groups.append(current)
                current = []
                continue

            if not current:
                current = [row]
                continue
            anchor = current[0]
            same_page = int(row.get("page") or 0) == int(anchor.get("page") or 0)
            same_y = abs(float(row["y_pt"]) - float(anchor["y_pt"])) <= tolerance_pt
            if same_page and same_y:
                current.append(row)
            else:
                if len(current) >= 2:
                    groups.append(current)
                current = [row]
        if len(current) >= 2:
            groups.append(current)
    return groups


def _collect_working_inventory(docx: Path):
    session_meta = None
    with OwnedWord(visible=False, require_clean=False) as (word, meta):
        doc = word.Documents.OpenNoRepairDialog(str(docx), False, True, False)
        try:
            inventory = _collect_com_inventory(doc)
        finally:
            doc.Close(False)
        session_meta = meta
    return inventory, session_meta.to_dict() if session_meta else None


def audit(source_snapshot: dict, working_docx: Path):
    src = source_snapshot.get("inventory", {}).get("omath", [])
    working_inventory, session = _collect_working_inventory(working_docx)
    dst = working_inventory.get("axmath", [])
    layout = _axmath_layout_inventory(working_docx)
    layout_by_ord = {int(x["ordinal"]): x for x in layout}

    report = {
        "schema": "axmath-geometry-audit/v2",
        "source_snapshot": source_snapshot.get("source"),
        "source_sha256": source_snapshot.get("docx_sha256"),
        "working_docx": str(working_docx.resolve()),
        "working_sha256": _sha256_file(working_docx),
        "word_session": session,
        "source_count": len(src),
        "working_count": len(dst),
        "status": "ok",
        "pairs": [],
        "alignment_breaks": [],
        "center_alignment_breaks": [],
        "same_line_groups": [],
        "same_line_breaks": [],
        "semantic_rebuild_candidates": [],
        "semantic_rebuild_ordinals": [],
        "prime_semantic_candidates": [],
        "prime_semantic_ordinals": [],
        "roundtrip_semantic_risk_candidates": [],
        "roundtrip_semantic_risk_ordinals": [],
        "visual_wrap_loss_candidates": [],
        "visual_wrap_loss_ordinals": [],
        "calibration_plan": [],
        "calibration_ordinals": [],
    }
    if len(src) != len(dst):
        report["status"] = "count_mismatch"
        return report

    raw_pairs = []
    ratios = []
    for a, b in zip(src, dst):
        sw = a.get("visual_width_pt")
        bw = b.get("width_pt")
        reliable = bool(a.get("visual_width_reliable")) and sw is not None and float(sw) > 4.0
        ratio = None
        if reliable and bw is not None and float(bw) > 0:
            ratio = float(bw) / float(sw)
            if math.isfinite(ratio) and 0.1 < ratio < 10:
                ratios.append(ratio)
        raw_pairs.append((a, b, reliable, ratio))

    enough_reference_pairs = len(ratios) >= 3
    center = _median(ratios) if enough_reference_pairs else None
    mad = _mad(ratios, center)
    # Three pairs is the structural minimum for a median to distinguish one
    # outlier from two agreeing references. With fewer references, fail closed
    # and require review instead of letting the target calibrate itself.
    spread = max(0.04, (mad or 0.0) * 5.0) if enough_reference_pairs else None
    report["learned_geometry"] = {
        "reliable_pair_count": len(ratios),
        "enough_reference_pairs": enough_reference_pairs,
        "minimum_pairs_for_auto_calibration": 3,
        "median_axmath_to_source_width_ratio": center,
        "ratio_mad": mad,
        "outlier_half_band": spread,
    }

    by_ord = {}
    for a, b, reliable, ratio in raw_pairs:
        ordinal = int(a.get("ordinal") or 0)
        residual = (ratio - center) if ratio is not None and center is not None else None
        outlier = bool(residual is not None and abs(residual) > spread)
        row = {
            "ordinal": ordinal,
            "source": {
                "page": a.get("page"),
                "x_pt": a.get("x_pt"),
                "y_pt": a.get("y_pt"),
                "end_page": a.get("end_page"),
                "end_x_pt": a.get("end_x_pt"),
                "end_y_pt": a.get("end_y_pt"),
                "same_visual_line": a.get("same_visual_line"),
                "visual_width_pt": a.get("visual_width_pt"),
                "text_right_pt": a.get("text_right_pt"),
                "text_bounds_reliable": a.get("text_bounds_reliable"),
                "visual_width_reliable": a.get("visual_width_reliable"),
                "paragraph_start": a.get("paragraph_start"),
                "source_inline_context": a.get("source_inline_context"),
                "source_inline_context_strong": a.get("source_inline_context_strong"),
                "source_inline_reasons": a.get("source_inline_reasons", []),
                "paragraph_format": a.get("paragraph_format"),
                "text": a.get("text"),
            },
            "working": {
                "page": b.get("page"),
                "x_pt": b.get("x_pt"),
                "y_pt": b.get("y_pt"),
                "width_pt": b.get("width_pt"),
                "height_pt": b.get("height_pt"),
                "paragraph_start": b.get("paragraph_start"),
                "paragraph_format": b.get("paragraph_format"),
                "text_right_pt": b.get("text_right_pt"),
                "text_bounds_reliable": b.get("text_bounds_reliable"),
                "right_edge_pt": b.get("right_edge_pt"),
                "right_overflow_pt": b.get("right_overflow_pt"),
                "overflows_text_right": b.get("overflows_text_right"),
            },
            "width_ratio": ratio,
            "width_ratio_residual": residual,
            "width_ratio_outlier": outlier,
        }
        report["pairs"].append(row)
        by_ord[ordinal] = row

        # Prime/derivative semantics are independent of geometry.  A formula
        # may look perfectly aligned while the AxMath parser received the wrong
        # prime syntax.  Route every frozen-source prime formula through the
        # source-semantic prime contract.
        source_text = str(a.get("text") or "")
        if contains_prime_or_derivative_marker(source_text):
            report["prime_semantic_candidates"].append({
                "ordinal": ordinal,
                "reason": "source_contains_prime_or_derivative_marker",
                "source_text": source_text,
                "preferred_route": "source_semantic_rebuild_native_prime",
                "requires_source_semantic_rebuild": True,
                "auto_apply": False,
            })

        # A tiny shell is legitimate for a single glyph such as x, 0 or alpha.
        # It is suspicious when the frozen source at the same ordinal contains
        # a non-trivial expression. This is triage only; GPT must visually or
        # semantically confirm the item before a Class E source rebuild.
        source_text_compact = "".join(str(a.get("text") or "").split())
        if (
            b.get("width_pt") is not None
            and float(b["width_pt"]) <= 10.0
            and len(source_text_compact) > 1
        ):
            report["semantic_rebuild_candidates"].append({
                "ordinal": ordinal,
                "reason": "nontrivial_source_in_tiny_axmath_shell",
                "source_text": a.get("text"),
                "source_text_compact_length": len(source_text_compact),
                "working_width_pt": float(b["width_pt"]),
                "working_height_pt": b.get("height_pt"),
                "requires_visual_confirmation": True,
                "auto_apply": False,
            })

        source_page = a.get("page")
        source_end_page = a.get("end_page")
        source_y = a.get("y_pt")
        source_end_y = a.get("end_y_pt")
        source_multiline = bool(
            a.get("same_visual_line") is False
            and source_page is not None
            and source_end_page is not None
            and (
                int(source_page) != int(source_end_page)
                or (
                    source_y is not None
                    and source_end_y is not None
                    and abs(float(source_end_y) - float(source_y)) > 1.5
                )
            )
        )
        if source_multiline and b.get("overflows_text_right") is True:
            report["visual_wrap_loss_candidates"].append({
                "ordinal": ordinal,
                "reason": "source_multiline_working_ole_crosses_measured_text_boundary",
                "repair_class": "M2_SOURCE_VISUAL_WRAP_LOSS",
                "source_page": source_page,
                "source_end_page": source_end_page,
                "source_y_pt": source_y,
                "source_end_y_pt": source_end_y,
                "working_page": b.get("page"),
                "working_width_pt": b.get("width_pt"),
                "working_right_edge_pt": b.get("right_edge_pt"),
                "working_text_right_pt": b.get("text_right_pt"),
                "working_right_overflow_pt": b.get("right_overflow_pt"),
                "auto_apply": False,
                "next_tool": "export_source_visual_lines.ps1",
                "repair_tool": "repair_axmath_from_approved_tex.ps1",
            })

        source_fmt = a.get("paragraph_format") or {}
        working_fmt = b.get("paragraph_format") or {}
        source_alignment = source_fmt.get("alignment")
        working_alignment = working_fmt.get("alignment")
        if (
            source_alignment is not None
            and working_alignment is not None
            and int(source_alignment) != int(working_alignment)
        ):
            rec = {
                "ordinal": ordinal,
                "source_alignment": source_alignment,
                "working_alignment": working_alignment,
                "source_page": a.get("page"),
                "working_page": b.get("page"),
                "source_paragraph_start": a.get("paragraph_start"),
                "working_paragraph_start": b.get("paragraph_start"),
            }
            report["alignment_breaks"].append(rec)
            # Word WdParagraphAlignment: wdAlignParagraphCenter == 1.
            if int(source_alignment) == 1:
                report["center_alignment_breaks"].append(rec)

    broken_members = set()
    for group in _source_same_line_groups(src):
        ordinals = [int(r["ordinal"]) for r in group]
        source_y = [float(r["y_pt"]) for r in group]
        targets = [by_ord[o]["working"] for o in ordinals if o in by_ord]
        target_pages = {int(t["page"]) for t in targets if t.get("page")}
        target_y = [float(t["y_pt"]) for t in targets if t.get("y_pt") is not None]
        broken = (
            len(targets) != len(ordinals)
            or len(target_pages) != 1
            or len(target_y) != len(ordinals)
            or (target_y and (max(target_y) - min(target_y) > max(2.0, (max(source_y) - min(source_y)) + 1.5)))
        )
        rec = {
            "paragraph_start": group[0].get("paragraph_start"),
            "ordinals": ordinals,
            "source_page": group[0].get("page"),
            "source_y_span_pt": max(source_y) - min(source_y),
            "working_pages": sorted(target_pages),
            "working_y_span_pt": (max(target_y) - min(target_y)) if target_y else None,
            "broken": broken,
        }
        report["same_line_groups"].append(rec)
        if broken:
            report["same_line_breaks"].append(rec)
            broken_members.update(ordinals)

    # AxMath -> TeX roundtrip has been observed to drop prime/derivative
    # semantics while still yielding a syntactically valid donor. Keep the
    # broken-same-line risk queue, while prime_semantic_candidates above covers
    # all prime formulas even when geometry looks normal.
    for ordinal in sorted(broken_members):
        row = by_ord.get(ordinal)
        if not row:
            continue
        source_text = str(row["source"].get("text") or "")
        if contains_prime_or_derivative_marker(source_text):
            report["roundtrip_semantic_risk_candidates"].append({
                "ordinal": ordinal,
                "reason": "source_contains_prime_or_derivative_marker",
                "source_text": source_text,
                "preferred_route": "source_semantic_rebuild",
            })

    if center is not None:
        for row in report["pairs"]:
            ordinal = row["ordinal"]
            source_width = row["source"].get("visual_width_pt")
            current_width = row["working"].get("width_pt")
            current_height = row["working"].get("height_pt")
            if (
                not row["source"].get("visual_width_reliable")
                or source_width is None or current_width is None or current_height is None
                or float(source_width) <= 4.0 or float(current_width) <= 0
            ):
                continue
            # Auto-calibrate only robust width-ratio outliers. A broken same-line
            # group is an additional reason/evidence, not permission by itself.
            if not row["width_ratio_outlier"]:
                continue
            target_width = float(source_width) * float(center)
            scale = target_width / float(current_width)
            if not (0.25 <= scale <= 4.0):
                continue
            layout_row = layout_by_ord.get(ordinal) or {}
            pos = layout_row.get("run_position_half_points")
            dxa = layout_row.get("dxa_orig")
            dya = layout_row.get("dya_orig")
            plan = {
                "ordinal": ordinal,
                "reason": [
                    "reference_width_ratio_outlier",
                    *(["source_same_line_broken"] if ordinal in broken_members else []),
                ],
                "source_width_pt": float(source_width),
                "source_inline_context": bool(row["source"].get("source_inline_context")),
                "source_inline_context_strong": bool(row["source"].get("source_inline_context_strong")),
                "source_inline_reasons": row["source"].get("source_inline_reasons", []),
                "learned_width_ratio": float(center),
                "current_width_pt": float(current_width),
                "current_height_pt": float(current_height),
                "target_width_pt": target_width,
                "target_height_pt": float(current_height) * scale,
                "scale": scale,
                "current_position_half_points": pos,
                "target_position_half_points": int(round(float(pos) * scale)) if isinstance(pos, (int, float)) else None,
                "current_dxa_orig": dxa,
                "current_dya_orig": dya,
                # dxaOrig/dyaOrig are Word twips. Native AxMath rebuild can
                # refresh v:shape width/height while leaving these original
                # extents stale, so scaling the stale values preserves the
                # inconsistency. Re-derive them from the calibrated shell.
                "target_dxa_orig": int(round(target_width * 20.0)),
                "target_dya_orig": int(round(float(current_height) * scale * 20.0)),
                "auto_apply": False,
                "requires_visual_confirmation": True,
            }
            report["calibration_plan"].append(plan)

    report["calibration_ordinals"] = [x["ordinal"] for x in report["calibration_plan"]]
    report["semantic_rebuild_ordinals"] = [
        x["ordinal"] for x in report["semantic_rebuild_candidates"]
    ]
    report["prime_semantic_ordinals"] = [
        x["ordinal"] for x in report["prime_semantic_candidates"]
    ]
    report["roundtrip_semantic_risk_ordinals"] = [
        x["ordinal"] for x in report["roundtrip_semantic_risk_candidates"]
    ]
    report["visual_wrap_loss_ordinals"] = [
        x["ordinal"] for x in report["visual_wrap_loss_candidates"]
    ]
    report["planned_same_line_breaks"] = [
        g for g in report["same_line_breaks"]
        if any(o in report["calibration_ordinals"] for o in g["ordinals"])
    ]
    # A proposed calibration is not a completed repair. Keep the group unresolved
    # until a later fresh audit or direct visual review proves it fixed.
    report["unresolved_same_line_breaks"] = list(report["same_line_breaks"])
    report["strict_layout_ok"] = bool(
        not report["center_alignment_breaks"]
        and not report["unresolved_same_line_breaks"]
        and not report["visual_wrap_loss_candidates"]
    )
    return report


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source-snapshot", required=True)
    ap.add_argument("--working", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    source_snapshot = json.loads(Path(args.source_snapshot).read_text(encoding="utf-8-sig"))
    report = audit(source_snapshot, Path(args.working).resolve())
    Path(args.out).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "status": report["status"],
        "source_count": report["source_count"],
        "working_count": report["working_count"],
        "reliable_pairs": report.get("learned_geometry", {}).get("reliable_pair_count"),
        "same_line_breaks": len(report.get("same_line_breaks", [])),
        "alignment_breaks": len(report.get("alignment_breaks", [])),
        "center_alignment_breaks": len(report.get("center_alignment_breaks", [])),
        "strict_layout_ok": report.get("strict_layout_ok"),
        "calibration_count": len(report.get("calibration_plan", [])),
        "calibration_ordinals": report.get("calibration_ordinals", []),
        "semantic_rebuild_candidate_count": len(report.get("semantic_rebuild_candidates", [])),
        "semantic_rebuild_ordinals": report.get("semantic_rebuild_ordinals", []),
        "prime_semantic_candidate_count": len(report.get("prime_semantic_candidates", [])),
        "prime_semantic_ordinals": report.get("prime_semantic_ordinals", []),
        "roundtrip_semantic_risk_count": len(report.get("roundtrip_semantic_risk_candidates", [])),
        "roundtrip_semantic_risk_ordinals": report.get("roundtrip_semantic_risk_ordinals", []),
        "visual_wrap_loss_count": len(report.get("visual_wrap_loss_candidates", [])),
        "visual_wrap_loss_ordinals": report.get("visual_wrap_loss_ordinals", []),
        "unresolved_same_line_breaks": len(report.get("unresolved_same_line_breaks", [])),
        "out": str(Path(args.out).resolve()),
    }, ensure_ascii=True, indent=2))
    return 0 if report["status"] == "ok" else 1


if __name__ == "__main__":
    raise SystemExit(main())
