from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from build_source_tex_map import NS, static_axmath_inventory
from snapshot_docx import _range_visual_geometry
from word_runtime import OwnedWord


def _utf8_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _write(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def _normalize_anchor(value: str) -> str:
    s = unicodedata.normalize("NFKC", value or "")
    s = re.sub(r"[\s，,、。.;；:：'\"“”‘’（）()\[\]【】]+", "", s)
    return s


def _source_paragraph_inventory(source: Path) -> list[dict]:
    with zipfile.ZipFile(source) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    rows = []
    ordinal = 0
    for index, paragraph in enumerate(root.findall(".//w:p", NS), 1):
        ordinals = []
        for _ in paragraph.findall(".//m:oMath", NS):
            ordinal += 1
            ordinals.append(ordinal)
        text = "".join(x.text or "" for x in paragraph.findall(".//w:t", NS))
        rows.append(
            {
                "paragraph_index": index,
                "text": text,
                "normalized_text": _normalize_anchor(text),
                "omath_ordinals": ordinals,
            }
        )
    return rows


def _working_paragraph_texts(working: Path) -> list[dict]:
    with zipfile.ZipFile(working) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    rows = []
    for index, paragraph in enumerate(root.findall(".//w:p", NS), 1):
        text = "".join(x.text or "" for x in paragraph.findall(".//w:t", NS))
        rows.append(
            {
                "paragraph_index": index,
                "text": text,
                "normalized_text": _normalize_anchor(text),
            }
        )
    return rows


def _find_unique_anchor(rows: list[dict], anchor: str, label: str) -> int:
    needle = _normalize_anchor(anchor)
    if not needle:
        raise ValueError("start anchor normalizes to empty text")
    matches = [
        int(row["paragraph_index"])
        for row in rows
        if needle in str(row["normalized_text"])
    ]
    if len(matches) != 1:
        raise ValueError(
            f"{label} start anchor must match exactly one paragraph; got {matches}"
        )
    return matches[0]


def _probe_source_multiline(source: Path, ordinals: list[int]) -> dict:
    targets = sorted({int(x) for x in ordinals if int(x) > 0})
    before = _sha256_file(source)
    result = {
        "schema": "axmath-source-multiline-probe/v1",
        "source": str(source),
        "source_sha256_before": before,
        "targets": targets,
        "rows": [],
    }
    if not targets:
        result["source_sha256_after"] = before
        result["source_unchanged"] = True
        return result

    meta_out = None
    with OwnedWord(visible=False, require_clean=False) as (word, meta):
        doc = word.Documents.OpenNoRepairDialog(str(source), False, True, False)
        try:
            total = int(doc.OMaths.Count)
            for ordinal in targets:
                if ordinal > total:
                    result["rows"].append(
                        {
                            "ordinal": ordinal,
                            "success": False,
                            "error": f"ordinal out of range 1..{total}",
                        }
                    )
                    continue
                rng = doc.OMaths.Item(ordinal).Range
                start = rng.Duplicate
                start.Collapse(1)
                start_page = int(start.Information(3))
                start_y = float(start.Information(6))
                geo = _range_visual_geometry(rng)
                end_page = geo.get("end_page")
                end_y = geo.get("end_y_pt")
                multiline = bool(
                    not geo.get("same_visual_line")
                    and end_page is not None
                    and (
                        start_page != int(end_page)
                        or (
                            end_y is not None
                            and abs(float(end_y) - start_y) > 1.5
                        )
                    )
                )
                result["rows"].append(
                    {
                        "ordinal": ordinal,
                        "success": True,
                        "source_multiline": multiline,
                        "start_page": start_page,
                        "start_y_pt": start_y,
                        "end_page": end_page,
                        "end_y_pt": end_y,
                        "same_visual_line": geo.get("same_visual_line"),
                    }
                )
        finally:
            doc.Close(False)
        meta_out = meta.to_dict()

    after = _sha256_file(source)
    result["source_sha256_after"] = after
    result["source_unchanged"] = after == before
    result["word_session"] = meta_out
    if not result["source_unchanged"]:
        raise RuntimeError("frozen source changed during local-layout diagnosis")
    return result


def diagnose_local_layout(
    source: Path, working: Path, start_anchor: str, outdir: Path
) -> dict:
    source = source.resolve()
    working = working.resolve()
    outdir = outdir.resolve()
    outdir.mkdir(parents=True, exist_ok=True)

    source_rows = _source_paragraph_inventory(source)
    working_rows = _working_paragraph_texts(working)
    inv = static_axmath_inventory(working)

    if int(inv["omath_count"]) != 0:
        raise ValueError(
            f"local layout repair requires a finished AxMath baseline with OfficeMath=0; "
            f"found {inv['omath_count']}"
        )

    source_anchor = _find_unique_anchor(source_rows, start_anchor, "source")
    working_anchor = _find_unique_anchor(working_rows, start_anchor, "working")
    offset = working_anchor - source_anchor

    overwide = [
        row
        for row in inv.get("conservative_overwide_candidates") or []
        if int(row["paragraph_index"]) > working_anchor
    ]

    ax_by_paragraph: dict[int, list[dict]] = {}
    for row in inv.get("locations") or []:
        ax_by_paragraph.setdefault(int(row["paragraph_index"]), []).append(row)

    prelim = []
    mapping_errors = []
    for row in overwide:
        working_p = int(row["paragraph_index"])
        source_p = working_p - offset
        if source_p < 1 or source_p > len(source_rows):
            mapping_errors.append(
                f"working paragraph {working_p} maps outside source via offset {offset}"
            )
            continue
        source_row = source_rows[source_p - 1]
        target_ax = ax_by_paragraph.get(working_p, [])
        ordinals = list(source_row["omath_ordinals"])
        if len(target_ax) != 1:
            mapping_errors.append(
                f"working paragraph {working_p} has {len(target_ax)} AxMath objects; "
                "local visual-line split requires exactly one current object"
            )
            continue
        if len(ordinals) != 1:
            mapping_errors.append(
                f"source paragraph {source_p} has {len(ordinals)} OfficeMath objects; "
                "this class requires exactly one source OfficeMath"
            )
            continue
        prelim.append(
            {
                "working_paragraph_index": working_p,
                "source_paragraph_index": source_p,
                "source_ordinal": int(ordinals[0]),
                "working_width_pt": row.get("width_pt"),
                "section_text_width_pt": row.get("section_text_width_pt"),
                "overflow_vs_full_text_width_pt": row.get(
                    "overflow_vs_full_text_width_pt"
                ),
            }
        )

    max_source_target = max(
        [int(x["source_paragraph_index"]) for x in prelim] or [source_anchor]
    )
    alignment_checked = 0
    alignment_matches = 0
    alignment_mismatches = []
    for source_p in range(source_anchor, max_source_target + 1):
        src = source_rows[source_p - 1]
        text = str(src["normalized_text"])
        if len(text) < 4:
            continue
        working_p = source_p + offset
        if working_p < 1 or working_p > len(working_rows):
            alignment_mismatches.append(
                {"source_paragraph": source_p, "working_paragraph": working_p}
            )
            continue
        alignment_checked += 1
        if text == str(working_rows[working_p - 1]["normalized_text"]):
            alignment_matches += 1
        else:
            alignment_mismatches.append(
                {
                    "source_paragraph": source_p,
                    "working_paragraph": working_p,
                    "source_text": src["text"][:80],
                    "working_text": working_rows[working_p - 1]["text"][:80],
                }
            )
            if len(alignment_mismatches) >= 8:
                break

    if alignment_checked < 5 or alignment_mismatches:
        raise ValueError(
            "paragraph offset is not stable enough for automatic local repair: "
            f"checked={alignment_checked}, matches={alignment_matches}, "
            f"mismatches={alignment_mismatches}"
        )

    probe = _probe_source_multiline(
        source, [int(x["source_ordinal"]) for x in prelim]
    )
    _write(outdir / "SOURCE_MULTILINE_PROBE.json", probe)
    by_ordinal = {
        int(row["ordinal"]): row
        for row in probe.get("rows") or []
        if row.get("success")
    }

    rows = []
    for item in prelim:
        probe_row = by_ordinal.get(int(item["source_ordinal"]))
        if probe_row and probe_row.get("source_multiline"):
            rows.append({**item, "source_probe": probe_row})

    plan = {
        "schema": "axmath-local-layout-plan/v1",
        "repair_class": "M2_SOURCE_VISUAL_LINE_SPLIT",
        "source": str(source),
        "source_sha256": _sha256_file(source),
        "working": str(working),
        "working_sha256": _sha256_file(working),
        "working_paragraph_count": int(inv["paragraph_count"]),
        "working_axmath_count": int(inv["axmath_count"]),
        "working_omath_count": int(inv["omath_count"]),
        "start_anchor": start_anchor,
        "source_anchor_paragraph": source_anchor,
        "working_anchor_paragraph": working_anchor,
        "paragraph_offset": offset,
        "alignment_checked": alignment_checked,
        "alignment_matches": alignment_matches,
        "first_repair_paragraph": (
            min(int(x["working_paragraph_index"]) for x in rows) if rows else None
        ),
        "rows": rows,
    }
    plan_path = outdir / "LOCAL_LAYOUT_PLAN.json"
    _write(plan_path, plan)

    ordinals = ",".join(str(x["source_ordinal"]) for x in rows)
    visual_json = outdir / "SOURCE_VISUAL_LINES.json"
    split_map = outdir / "VISUAL_LINE_SPLIT_MAP.json"
    repaired = outdir / (working.stem + "_visual-lines-fixed.docx")
    repair_report = outdir / "VISUAL_LINE_SPLIT_REPAIR.json"
    validation = outdir / "VISUAL_LINE_SPLIT_VALIDATION.json"

    report = {
        "schema": "axmath-local-layout-diagnosis/v1",
        "status": "repair_required" if rows else "ready_no_local_overflow",
        "repair_class": "M2_SOURCE_VISUAL_LINE_SPLIT" if rows else None,
        "source": str(source),
        "working": str(working),
        "start_anchor": start_anchor,
        "source_anchor_paragraph": source_anchor,
        "working_anchor_paragraph": working_anchor,
        "paragraph_offset": offset,
        "candidate_count": len(rows),
        "mapping_errors": mapping_errors,
        "plan": str(plan_path),
        "next_actions": [],
    }
    if rows:
        report["next_actions"] = [
            (
                "powershell -File scripts/export_source_visual_lines.ps1 "
                f'-SourceDocx "{source}" -Ordinals "{ordinals}" '
                f'-OutputJson "{visual_json}"'
            ),
            (
                "python scripts/build_source_tex_map.py visual-line-split "
                f'--visual-report "{visual_json}" --working "{working}" '
                f'--paragraph-plan "{plan_path}" --out "{split_map}"'
            ),
            (
                "powershell -File scripts/repair_visual_line_split.ps1 "
                f'-InputDocx "{working}" -OutputDocx "{repaired}" '
                f'-MapPath "{split_map}" -ReportPath "{repair_report}"'
            ),
            (
                "python scripts/validate_local_visual_line_repair.py "
                f'--baseline "{working}" --candidate "{repaired}" '
                f'--map "{split_map}" --out "{validation}"'
            ),
        ]
    return report


def main() -> int:
    _utf8_console()
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--working", required=True)
    ap.add_argument("--start-anchor", required=True)
    ap.add_argument("--outdir", required=True)
    args = ap.parse_args()

    report = diagnose_local_layout(
        Path(args.source),
        Path(args.working),
        args.start_anchor,
        Path(args.outdir),
    )
    out = Path(args.outdir).resolve() / "LOCAL_LAYOUT_NEXT_ACTION.json"
    _write(out, report)
    print(
        json.dumps(
            {
                "status": report["status"],
                "repair_class": report.get("repair_class"),
                "candidate_count": report.get("candidate_count"),
                "paragraph_offset": report.get("paragraph_offset"),
                "next_action": str(out),
            },
            ensure_ascii=True,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
