from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path

from audit_docx import compare
from build_source_tex_map import build_multisibling_map, static_axmath_inventory
from formula_geometry_audit import audit as geometry_audit
from snapshot_docx import _range_visual_geometry, snapshot
from word_runtime import OwnedWord


def _utf8_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _write(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _source_multiline_probe(source: Path, ordinals: list[int]) -> dict:
    """Probe only candidate source formulas with Word layout.

    This is intentionally O(number of suspicious formulas), not O(all formulas).
    No AxMath add-in is loaded; source is opened read-only and SHA-bound.
    """
    targets = sorted({int(x) for x in ordinals if int(x) > 0})
    source = source.resolve()
    before = _sha256_file(source)
    result = {
        "schema": "axmath-source-multiline-probe/v1",
        "source": str(source),
        "source_sha256_before": before,
        "targets": targets,
        "rows": [],
        "word_session": None,
        "source_unchanged": False,
    }
    if not targets:
        result["source_sha256_after"] = before
        result["source_unchanged"] = True
        return result

    session_meta = None
    with OwnedWord(visible=False, require_clean=False) as (word, meta):
        doc = word.Documents.OpenNoRepairDialog(str(source), False, True, False)
        try:
            count = int(doc.OMaths.Count)
            for ordinal in targets:
                if ordinal > count:
                    result["rows"].append({
                        "ordinal": ordinal,
                        "success": False,
                        "error": f"ordinal out of range 1..{count}",
                    })
                    continue
                rng = doc.OMaths.Item(ordinal).Range
                start_page = None
                start_y = None
                try:
                    start = rng.Duplicate
                    start.Collapse(1)
                    start_page = int(start.Information(3))
                    start_y = float(start.Information(6))
                except Exception:
                    pass
                geo = _range_visual_geometry(rng)
                end_page = geo.get("end_page")
                end_y = geo.get("end_y_pt")
                multiline = bool(
                    not geo.get("same_visual_line")
                    and start_page is not None
                    and end_page is not None
                    and (
                        int(start_page) != int(end_page)
                        or (
                            start_y is not None
                            and end_y is not None
                            and abs(float(end_y) - float(start_y)) > 1.5
                        )
                    )
                )
                result["rows"].append({
                    "ordinal": ordinal,
                    "success": True,
                    "source_multiline": multiline,
                    "start_page": start_page,
                    "start_y_pt": start_y,
                    "end_page": end_page,
                    "end_y_pt": end_y,
                    "same_visual_line": geo.get("same_visual_line"),
                })
        finally:
            doc.Close(False)
        session_meta = meta
    result["word_session"] = session_meta.to_dict() if session_meta else None
    after = _sha256_file(source)
    result["source_sha256_after"] = after
    result["source_unchanged"] = after == before
    if not result["source_unchanged"]:
        raise RuntimeError("Frozen source changed during source multiline probe.")
    return result


def _fast_static_triage(source: Path, working: Path, content: dict, outdir: Path) -> dict:
    """High-confidence static triage before expensive whole-document COM geometry."""
    inv = static_axmath_inventory(working)
    source_residual = (content.get("source") or {}).get("residual") or []
    source_text_by_ordinal = {
        i: str(row.get("text") or "") for i, row in enumerate(source_residual, 1)
    }

    tiny = []
    for row in inv.get("locations") or []:
        width = row.get("width_pt")
        ordinal = int(row["ordinal"])
        compact = re.sub(r"\s+", "", source_text_by_ordinal.get(ordinal, ""))
        if width is not None and float(width) <= 10.0 and len(compact) > 1:
            tiny.append({
                "ordinal": ordinal,
                "paragraph_index": row.get("paragraph_index"),
                "width_pt": width,
                "height_pt": row.get("height_pt"),
                "source_text": source_text_by_ordinal.get(ordinal),
                "repair_class": "CLASS_E_SEMANTIC_REBUILD",
                "reason": "nontrivial_source_in_tiny_static_axmath_shell",
            })

    overwide = list(inv.get("conservative_overwide_candidates") or [])
    probe = _source_multiline_probe(
        source, [int(x["ordinal"]) for x in overwide]
    )
    _write(outdir / "FAST_SOURCE_MULTILINE_PROBE.json", probe)
    probe_by_ord = {
        int(x["ordinal"]): x
        for x in probe.get("rows") or []
        if x.get("success")
    }

    m2 = []
    overwide_singleline = []
    for row in overwide:
        ordinal = int(row["ordinal"])
        src = probe_by_ord.get(ordinal)
        merged = {**row, "source_probe": src}
        if src and src.get("source_multiline"):
            merged["repair_class"] = "M2_SOURCE_VISUAL_WRAP_LOSS"
            merged["reason"] = (
                "fixed AxMath OLE exceeds the entire measured section text width "
                "and the frozen source formula spans multiple Word visual lines"
            )
            m2.append(merged)
        else:
            merged["repair_class"] = "OVERWIDE_SOURCE_SINGLELINE_REVIEW"
            merged["reason"] = (
                "fixed AxMath OLE exceeds the entire measured section text width, "
                "but source multiline intent was not proven"
            )
            overwide_singleline.append(merged)

    report = {
        "schema": "axmath-fast-static-triage/v1",
        "working_inventory": {
            "paragraph_count": inv.get("paragraph_count"),
            "omath_count": inv.get("omath_count"),
            "axmath_count": inv.get("axmath_count"),
            "section_text_width_pt": inv.get("section_text_width_pt"),
            "section_measurement_reason": inv.get("section_measurement_reason"),
        },
        "semantic_tiny_shell_candidates": tiny,
        "m2_visual_wrap_loss_candidates": m2,
        "overwide_source_singleline_review": overwide_singleline,
        "source_multiline_probe": str(outdir / "FAST_SOURCE_MULTILINE_PROBE.json"),
    }
    _write(outdir / "FAST_TRIAGE.json", report)
    return report


def _append_m2_actions(report: dict, source: Path, working: Path, outdir: Path, ordinals: list[int]) -> None:
    if not ordinals:
        return
    ords = ",".join(str(x) for x in sorted({int(x) for x in ordinals}))
    visual_json = outdir / "M2_SOURCE_VISUAL_LINES.json"
    m2_map = outdir / "M2_ALIGNED_REPAIR_MAP.json"
    m2_out = outdir / (working.stem + "_M2-fixed.docx")
    report["next_actions"].extend([
        (
            "powershell -File scripts/export_source_visual_lines.ps1 "
            f'-SourceDocx "{source}" -Ordinals "{ords}" '
            f'-OutputJson "{visual_json}"'
        ),
        (
            "python scripts/build_source_tex_map.py visual-wrap "
            f'--visual-report "{visual_json}" --working "{working}" --out "{m2_map}"'
        ),
        (
            "powershell -File scripts/repair_axmath_from_approved_tex.ps1 "
            f'-InputDocx "{working}" -OutputDocx "{m2_out}" -MapPath "{m2_map}"'
        ),
        (
            "M2 must remain one source formula -> one aligned Equation.AxMath object; "
            "re-run this diagnosis on the new copy after repair."
        ),
    ])


def diagnose(source: Path, working: Path, outdir: Path, *, deep_geometry: bool = False) -> dict:
    source = source.resolve()
    working = working.resolve()
    outdir = outdir.resolve()
    outdir.mkdir(parents=True, exist_ok=True)

    content = compare(source, working)
    _write(outdir / "CONTENT_AUDIT.json", content)
    count = content.get("formula_count_state") or {}
    state = count.get("state")

    report: dict = {
        "schema": "axmath-post-conversion-diagnosis/v2",
        "source": str(source),
        "working": str(working),
        "diagnosis_profile": "deep_geometry" if deep_geometry else "fast_static_targeted",
        "formula_count_state": count,
        "paragraph_count_equal": bool(content.get("paragraph_count_equal")),
        "nonmath_text_exact": bool(content.get("nonmath_text_exact")),
        "nonmath_text_contract_exact": bool(content.get("nonmath_text_contract_exact")),
        "numbering_punctuation_expected_change_count": int(
            content.get("numbering_punctuation_expected_change_count") or 0
        ),
        "status": "starting",
        "repair_class": None,
        "next_actions": [],
        "do_not": [
            "Do not tune page count directly.",
            "Do not run global width/height/w:position sweeps.",
            "Do not split one source auto-wrapped formula into multiple AxMath objects.",
            "Do not use AxMath->TeX roundtrip as the semantic source for derivative/prime formulas.",
        ],
    }

    if not content.get("paragraph_count_equal") or not content.get("nonmath_text_contract_exact"):
        report["status"] = "blocked_content_drift"
        report["repair_class"] = "STOP_CONTENT_DRIFT"
        return report

    if state == "known_multisibling_collapse":
        repair_map = build_multisibling_map(source, working)
        map_path = outdir / "M1_MULTISIBLING_REPAIR_MAP.json"
        _write(map_path, repair_map)
        repaired = outdir / (working.stem + "_M1-fixed.docx")
        report.update({
            "status": "repair_required",
            "repair_class": "M1_MULTISIBLING_OMATHPARA",
            "reason": (
                "The raw OfficeMath count gap exactly equals the frozen source "
                "m:oMathPara multi-sibling extra-node count. Repair M1 before "
                "geometry or page-flow tuning."
            ),
            "repair_map": str(map_path),
            "next_actions": [
                (
                    "powershell -File scripts/repair_multisibling_groups.ps1 "
                    f'-InputDocx "{working}" -OutputDocx "{repaired}" '
                    f'-MapPath "{map_path}"'
                ),
                (
                    "Re-run diagnose_after_conversion.py on the M1-fixed copy. "
                    "Do not continue to geometry repair until formula_count_state=exact."
                ),
            ],
        })
        return report

    if state != "exact":
        report["status"] = "blocked_unexplained_formula_count"
        report["repair_class"] = count.get("repair_class") or "STOP_UNEXPLAINED_COUNT_GAP"
        report["reason"] = (
            "Formula count is not explained by the frozen source structure. "
            "Stop mutation and investigate missing/excess/residual formulas."
        )
        return report

    fast = _fast_static_triage(source, working, content, outdir)
    report["fast_triage"] = str(outdir / "FAST_TRIAGE.json")
    fast_queues = {
        "M2_SOURCE_VISUAL_WRAP_LOSS": [
            int(x["ordinal"]) for x in fast["m2_visual_wrap_loss_candidates"]
        ],
        "CLASS_E_SEMANTIC_REBUILD": [
            int(x["ordinal"]) for x in fast["semantic_tiny_shell_candidates"]
        ],
        "OVERWIDE_SOURCE_SINGLELINE_REVIEW": [
            int(x["ordinal"]) for x in fast["overwide_source_singleline_review"]
        ],
    }
    report["queues"] = fast_queues

    if fast_queues["CLASS_E_SEMANTIC_REBUILD"]:
        report["status"] = "repair_review_required"
        report["repair_class"] = "CLASS_E_SEMANTIC_REBUILD"
        report["next_actions"].append(
            "For the listed tiny-shell ordinals, rebuild semantics only from the "
            "same frozen-source OfficeMath ordinal. Validate Word LaTeX and fall "
            "back to exact OMML->LaTeX when malformed; apply an approved map with "
            "repair_axmath_from_approved_tex.ps1."
        )

    if fast_queues["M2_SOURCE_VISUAL_WRAP_LOSS"]:
        report["status"] = "repair_review_required"
        report["repair_class"] = report["repair_class"] or "M2_SOURCE_VISUAL_WRAP_LOSS"
        _append_m2_actions(
            report, source, working, outdir, fast_queues["M2_SOURCE_VISUAL_WRAP_LOSS"]
        )

    if fast_queues["OVERWIDE_SOURCE_SINGLELINE_REVIEW"]:
        report["status"] = "repair_review_required"
        report["repair_class"] = report["repair_class"] or "OVERWIDE_SOURCE_SINGLELINE_REVIEW"
        report["next_actions"].append(
            "The listed AxMath OLEs exceed the full section text width but source "
            "multiline intent was not proven. Inspect those pages/formulas; do not "
            "auto-shrink or split them. Use --deep-geometry only if the visual "
            "defect cannot be classified from source-vs-working evidence."
        )

    # Fast high-confidence defects are handled before any expensive all-formula
    # Word COM scan. This is the default production path.
    if report["status"] != "starting":
        return report

    if not deep_geometry:
        report["status"] = "ready_for_strict_final_compare"
        report["reason"] = (
            "Content/count gates are exact and fast static+targeted triage found "
            "no high-confidence repair defect. Visual layout remains the final "
            "acceptance authority; do not run a whole-document geometry scan by "
            "default merely to make diagnostic metrics reach zero."
        )
        report["next_actions"] = [
            "Run strict_final_compare.py once, inspect every source-vs-final page, "
            "then finalize_visual_review.py.",
            (
                "If final visual review exposes an ambiguous formula defect that "
                "cannot be classified locally, re-run diagnose_after_conversion.py "
                "with --deep-geometry for the expensive all-formula diagnostic pass."
            ),
        ]
        return report

    # Explicit deep mode: preserve the existing comprehensive geometry route.
    source_snapshot = snapshot(
        source, outdir / "source_geometry", "source", profile="geometry"
    )
    geo = geometry_audit(source_snapshot, working)
    _write(outdir / "GEOMETRY_AUDIT.json", geo)
    report["geometry_audit"] = str(outdir / "GEOMETRY_AUDIT.json")

    deep_queues = {
        "M2_SOURCE_VISUAL_WRAP_LOSS": geo.get("visual_wrap_loss_ordinals", []),
        "CLASS_E_SEMANTIC_REBUILD": geo.get("semantic_rebuild_ordinals", []),
        "CLASS_E_PRIME_RISK": geo.get("roundtrip_semantic_risk_ordinals", []),
        "CLASS_A_SAME_LINE": sorted({
            int(o)
            for group in geo.get("same_line_breaks", [])
            for o in group.get("ordinals", [])
        }),
        "CLASS_C_EXTERNAL_BOX": geo.get("calibration_ordinals", []),
    }
    report["queues"] = deep_queues

    if deep_queues["CLASS_E_SEMANTIC_REBUILD"] or deep_queues["CLASS_E_PRIME_RISK"]:
        report["status"] = "repair_review_required"
        report["repair_class"] = "CLASS_E_SEMANTIC_REBUILD"
        report["next_actions"].append(
            "Export the same frozen-source ordinals with export_source_word_latex.ps1; "
            "validate the Word output, fall back to exact OMML->LaTeX when malformed, "
            "then apply only a GPT-approved map with repair_axmath_from_approved_tex.ps1."
        )

    if deep_queues["M2_SOURCE_VISUAL_WRAP_LOSS"]:
        report["status"] = "repair_review_required"
        report["repair_class"] = report["repair_class"] or "M2_SOURCE_VISUAL_WRAP_LOSS"
        _append_m2_actions(
            report, source, working, outdir, deep_queues["M2_SOURCE_VISUAL_WRAP_LOSS"]
        )

    if deep_queues["CLASS_A_SAME_LINE"]:
        report["status"] = "repair_review_required"
        report["next_actions"].append(
            "For visually confirmed low-semantic-risk same-line culprits only, use "
            "repair_axmath_inline_roundtrip.ps1 (max 12 reviewed targets). Prime/"
            "derivative formulas stay on Class E."
        )
    if deep_queues["CLASS_C_EXTERNAL_BOX"]:
        report["status"] = "repair_review_required"
        report["next_actions"].append(
            "Only after semantic/internal classes are cleared, visually confirm Class C "
            "and use calibrate_axmath_boxes.py with explicit ordinals."
        )

    if report["status"] == "starting":
        report["status"] = "ready_for_strict_final_compare"
        report["next_actions"] = [
            "Run strict_final_compare.py once, inspect every source-vs-final page, "
            "then finalize_visual_review.py. Do not mutate from diagnostics alone."
        ]
    return report


def main() -> int:
    _utf8_console()
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--working", required=True)
    ap.add_argument("--outdir", required=True)
    ap.add_argument(
        "--deep-geometry",
        action="store_true",
        help="Run the expensive all-formula Word COM geometry audit. Default is fast static+targeted triage.",
    )
    args = ap.parse_args()
    report = diagnose(
        Path(args.source),
        Path(args.working),
        Path(args.outdir),
        deep_geometry=args.deep_geometry,
    )
    out = Path(args.outdir).resolve() / "NEXT_ACTION.json"
    _write(out, report)
    print(json.dumps({
        "status": report["status"],
        "repair_class": report.get("repair_class"),
        "diagnosis_profile": report.get("diagnosis_profile"),
        "formula_count_state": report.get("formula_count_state", {}).get("state"),
        "queues": report.get("queues"),
        "next_action": str(out),
    }, ensure_ascii=True, indent=2))
    return 1 if str(report["status"]).startswith("blocked") else 0


if __name__ == "__main__":
    raise SystemExit(main())
