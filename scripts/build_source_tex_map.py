from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from omml2latex import convert_omml

from normalize_axmath_tex import canonical_prime_issues, normalize_axmath_tex, prime_orders
from source_math_structure import analyze_source_structure, classify_count_state
from source_tex_sanity import normalize_word_latex, strip_math_delimiters, word_latex_is_sane

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
}


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _style_pt(style: str | None, name: str) -> float | None:
    if not style:
        return None
    match = re.search(
        rf"(?:^|;)\s*{re.escape(name)}\s*:\s*([-+]?\d+(?:\.\d+)?)pt(?:;|$)",
        style,
        flags=re.IGNORECASE,
    )
    return float(match.group(1)) if match else None


def static_axmath_inventory(docx: Path) -> dict:
    """Fast DOCX/XML AxMath inventory without Word COM.

    Besides global/local location, capture the VML OLE shell width/height.  A
    single-section direct-body formula whose fixed OLE width exceeds the entire
    section text width is a conservative over-wide candidate: no arbitrary
    formula-width threshold is involved.
    """
    docx = Path(docx).resolve()
    with zipfile.ZipFile(docx) as z:
        root = ET.fromstring(z.read("word/document.xml"))

    body = root.find(".//w:body", NS)
    paragraphs = root.findall(".//w:p", NS)
    direct_body_ids = (
        {id(p) for p in body.findall("./w:p", NS)} if body is not None else set()
    )

    section_text_width_pt = None
    section_reason = "not_single_section"
    sects = root.findall(".//w:sectPr", NS)
    if len(sects) == 1:
        sect = sects[0]
        cols = sect.find("w:cols", NS)
        col_count = 1
        if cols is not None:
            try:
                col_count = int(cols.attrib.get("{%s}num" % NS["w"], "1"))
            except ValueError:
                col_count = 1
        pgsz = sect.find("w:pgSz", NS)
        pgmar = sect.find("w:pgMar", NS)
        if col_count == 1 and pgsz is not None and pgmar is not None:
            try:
                page_twips = float(pgsz.attrib["{%s}w" % NS["w"]])
                left_twips = float(pgmar.attrib["{%s}left" % NS["w"]])
                right_twips = float(pgmar.attrib["{%s}right" % NS["w"]])
                width = (page_twips - left_twips - right_twips) / 20.0
                if width > 0:
                    section_text_width_pt = width
                    section_reason = "single_section_single_column"
            except (KeyError, ValueError):
                section_reason = "section_measurement_failed"
        elif col_count != 1:
            section_reason = "multi_column_section"

    locations = []
    global_ordinal = 0
    for paragraph_index, paragraph in enumerate(paragraphs, 1):
        local_axmath = 0
        # Use w:object so the AxMath OLE and its VML shell stay associated.
        for obj in paragraph.findall(".//w:object", NS):
            ole = obj.find(".//o:OLEObject", NS)
            if ole is None or (ole.attrib.get("ProgID") or "").lower() != "equation.axmath":
                continue
            global_ordinal += 1
            local_axmath += 1
            shape = obj.find(".//v:shape", NS)
            style = shape.attrib.get("style") if shape is not None else None
            width_pt = _style_pt(style, "width")
            height_pt = _style_pt(style, "height")
            locations.append({
                "ordinal": global_ordinal,
                "paragraph_index": paragraph_index,
                "paragraph_axmath_index": local_axmath,
                "direct_body_paragraph": id(paragraph) in direct_body_ids,
                "width_pt": width_pt,
                "height_pt": height_pt,
            })

    overwide = []
    if section_text_width_pt is not None:
        for row in locations:
            width = row.get("width_pt")
            if (
                row.get("direct_body_paragraph")
                and width is not None
                and float(width) > float(section_text_width_pt) + 2.0
            ):
                overwide.append({
                    **row,
                    "section_text_width_pt": section_text_width_pt,
                    "overflow_vs_full_text_width_pt": float(width) - section_text_width_pt,
                })

    return {
        "paragraph_count": len(paragraphs),
        "omath_count": len(root.findall(".//m:oMath", NS)),
        "axmath_count": global_ordinal,
        "section_text_width_pt": section_text_width_pt,
        "section_measurement_reason": section_reason,
        "locations": locations,
        "conservative_overwide_candidates": overwide,
    }


def _first_math_node(root: ET.Element) -> ET.Element:
    node = root.find(".//m:oMathPara", NS)
    if node is not None:
        return node
    node = root.find(".//m:oMath", NS)
    if node is None:
        raise ValueError("DOCX fragment contains no OfficeMath")
    return node


def omml_tex_from_fragment(fragment: Path) -> str:
    with zipfile.ZipFile(fragment) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    return strip_math_delimiters(convert_omml(_first_math_node(root)))


def _source_omath_nodes(source: Path) -> list[ET.Element]:
    with zipfile.ZipFile(source) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    return root.findall(".//m:oMath", NS)


def _canonical_axmath_tex(tex: str) -> tuple[str, list[int]]:
    normalized = normalize_axmath_tex(strip_math_delimiters(tex))
    issues = canonical_prime_issues(normalized)
    if issues:
        raise ValueError(
            "source-derived TeX still violates the verified AxMath prime contract: "
            + ",".join(issues)
        )
    return normalized, prime_orders(normalized)


def _safe_source_tex(node: ET.Element) -> str:
    tex = strip_math_delimiters(convert_omml(node))
    if not tex:
        raise ValueError("OMML to LaTeX fallback returned empty output")
    normalized, _ = _canonical_axmath_tex(tex)
    return normalized


def _alignment_anchor(line: str) -> str:
    """Insert one aligned-column marker at the first top-level equality."""
    depth = 0
    for i, ch in enumerate(line):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth = max(0, depth - 1)
        elif ch == "=" and depth == 0:
            return line[:i] + "&" + line[i:]
    return "&" + line


def build_multisibling_map(source: Path, working: Path | None = None) -> dict:
    source = Path(source).resolve()
    structure = analyze_source_structure(source)
    nodes = _source_omath_nodes(source)
    rows = []
    for group in structure["multi_sibling_groups"]:
        lines = []
        for ordinal in group["source_ordinals"]:
            tex = _safe_source_tex(nodes[int(ordinal) - 1])
            _, orders = _canonical_axmath_tex(tex)
            lines.append({
                "source_ordinal": int(ordinal),
                "tex": "$" + tex + "$",
                "provenance": "frozen_source_exact_omml",
                "prime_orders": orders,
                "prime_contract": "axmath_builtin_prime_v2",
            })
        rows.append({
            "repair_class": "M1_MULTISIBLING_OMATHPARA",
            "paragraph_index": int(group["paragraph_index"]),
            "source_ordinals": [int(x) for x in group["source_ordinals"]],
            "expected_current_axmath": 1,
            "expected_after_axmath": len(lines),
            "lines": lines,
        })
    payload = {
        "schema": "axmath-repair-map/v2",
        "repair_class": "M1_MULTISIBLING_OMATHPARA",
        "source": str(source),
        "source_sha256": sha256_file(source),
        "source_raw_omath_count": int(structure["raw_omath_count"]),
        "expected_final_axmath_count": int(structure["raw_omath_count"]),
        "known_extra_nodes": int(structure["multi_sibling_extra_nodes"]),
        "rows": rows,
    }
    if working is not None:
        working = Path(working).resolve()
        inv = static_axmath_inventory(working)
        count_state = classify_count_state(
            structure, inv["axmath_count"], inv["omath_count"]
        )
        if count_state["state"] != "known_multisibling_collapse":
            raise ValueError(
                "M1 map requires the exact frozen-source multi-sibling collapse "
                f"signature; got {count_state['state']}"
            )
        payload["working"] = str(working)
        payload["working_sha256"] = sha256_file(working)
        payload["working_axmath_count"] = inv["axmath_count"]
        payload["working_omath_count"] = inv["omath_count"]
        payload["working_paragraph_count"] = inv["paragraph_count"]
        payload["count_state"] = count_state
    return payload


def build_visual_wrap_map(visual_report: Path, working: Path) -> dict:
    visual_report = Path(visual_report).resolve()
    working = Path(working).resolve()
    data = json.loads(visual_report.read_text(encoding="utf-8-sig"))
    if not data.get("source_unchanged"):
        raise ValueError("visual-line report does not prove frozen source integrity")

    source = Path(str(data.get("source") or "")).resolve()
    if not source.is_file():
        raise ValueError("visual-line report source is missing")
    structure = analyze_source_structure(source)
    inv = static_axmath_inventory(working)
    count_state = classify_count_state(
        structure, inv["axmath_count"], inv["omath_count"]
    )
    if count_state["state"] != "exact":
        raise ValueError(
            "M2 map is forbidden until formula identity/count is exact; "
            f"resolve {count_state['state']} first"
        )
    locations = {
        int(row["ordinal"]): row for row in inv["locations"]
    }
    rows = []
    for item in data.get("exports") or []:
        if not item.get("success"):
            raise ValueError("visual-line export failed for ordinal " + str(item.get("ordinal")))
        line_rows = item.get("lines") or []
        if len(line_rows) < 2:
            continue
        chosen = []
        for line in line_rows:
            word_latex = str(line.get("word_latex") or "")
            sane, reasons = word_latex_is_sane(word_latex)
            if sane:
                tex = normalize_word_latex(word_latex)
                provenance = "word_latex"
            else:
                fragment = Path(line["fragment_docx"])
                tex = omml_tex_from_fragment(fragment)
                provenance = "fragment_exact_omml_fallback"
            if not tex:
                raise ValueError(
                    "empty source TeX for ordinal "
                    + str(item.get("ordinal"))
                    + " line "
                    + str(line.get("line"))
                )
            tex, orders = _canonical_axmath_tex(tex)
            chosen.append({
                "line": int(line["line"]),
                "tex": tex,
                "provenance": provenance,
                "prime_orders": orders,
                "prime_contract": "axmath_builtin_prime_v2",
                "word_latex_rejected_reasons": [] if sane else reasons,
            })
        body = r" \\ ".join(_alignment_anchor(x["tex"]) for x in chosen)
        tex = r"$\begin{aligned} " + body + r" \end{aligned}$"
        ordinal = int(item["ordinal"])
        location = locations.get(ordinal)
        if not location:
            raise ValueError(f"working AxMath ordinal not found: {ordinal}")
        rows.append({
            "repair_class": "M2_SOURCE_VISUAL_WRAP_LOSS",
            "ordinal": ordinal,
            "working_paragraph_index": int(location["paragraph_index"]),
            "working_paragraph_axmath_index": int(location["paragraph_axmath_index"]),
            "tex": tex,
            "expected_visual_lines": len(chosen),
            "line_sources": chosen,
        })

    return {
        "schema": "axmath-repair-map/v2",
        "repair_class": "M2_SOURCE_VISUAL_WRAP_LOSS",
        "source": data.get("source"),
        "source_sha256": data.get("source_sha256_before"),
        "visual_line_report": str(visual_report),
        "working": str(working),
        "working_sha256": sha256_file(working),
        "working_axmath_count": inv["axmath_count"],
        "working_omath_count": inv["omath_count"],
        "working_paragraph_count": inv["paragraph_count"],
        "count_state": count_state,
        "rows": rows,
    }



def build_visual_line_split_map(
    visual_report: Path, working: Path, paragraph_plan: Path
) -> dict:
    """Build a hash-bound local-layout repair map.

    Unlike visual-wrap (one source OfficeMath -> one internal aligned AxMath),
    this mode preserves the source visual line structure: each proven source
    visual line becomes one editable AxMath object, with soft line breaks
    between objects inside the same Word paragraph.
    """
    visual_report = Path(visual_report).resolve()
    working = Path(working).resolve()
    paragraph_plan = Path(paragraph_plan).resolve()
    data = json.loads(visual_report.read_text(encoding="utf-8-sig"))
    plan = json.loads(paragraph_plan.read_text(encoding="utf-8-sig"))

    if not data.get("source_unchanged"):
        raise ValueError("visual-line report does not prove frozen source integrity")
    if str(plan.get("schema")) != "axmath-local-layout-plan/v1":
        raise ValueError("paragraph plan schema is not axmath-local-layout-plan/v1")

    working_sha = sha256_file(working)
    if str(plan.get("working_sha256") or "").lower() != working_sha.lower():
        raise ValueError("stale local-layout plan: working SHA does not match")
    source_sha = str(data.get("source_sha256_before") or "").lower()
    if source_sha != str(plan.get("source_sha256") or "").lower():
        raise ValueError("visual-line report source SHA does not match local-layout plan")

    inv = static_axmath_inventory(working)
    if int(plan.get("working_paragraph_count") or -1) != int(inv["paragraph_count"]):
        raise ValueError("local-layout plan paragraph count no longer matches working")
    if int(plan.get("working_axmath_count") or -1) != int(inv["axmath_count"]):
        raise ValueError("local-layout plan AxMath count no longer matches working")
    if int(inv["omath_count"]) != 0:
        raise ValueError("visual-line split repair requires working OfficeMath=0")

    exports = {
        int(item["ordinal"]): item
        for item in (data.get("exports") or [])
        if item.get("success")
    }
    rows = []
    expected_increase = 0
    for planned in plan.get("rows") or []:
        source_ordinal = int(planned["source_ordinal"])
        item = exports.get(source_ordinal)
        if item is None:
            raise ValueError(
                f"visual-line export missing source ordinal {source_ordinal}"
            )
        line_rows = item.get("lines") or []
        if len(line_rows) < 2:
            raise ValueError(
                f"source ordinal {source_ordinal} has fewer than two proven visual lines"
            )

        chosen = []
        for line in line_rows:
            word_latex = str(line.get("word_latex") or "")
            sane, reasons = word_latex_is_sane(word_latex)
            if sane:
                tex = normalize_word_latex(word_latex)
                provenance = "word_latex"
            else:
                fragment = Path(line["fragment_docx"])
                tex = omml_tex_from_fragment(fragment)
                provenance = "fragment_exact_omml_fallback"
            if not tex:
                raise ValueError(
                    f"empty source TeX for ordinal {source_ordinal} line {line.get('line')}"
                )
            tex, orders = _canonical_axmath_tex(tex)
            chosen.append(
                {
                    "line": int(line["line"]),
                    "tex": "$" + tex + "$",
                    "provenance": provenance,
                    "prime_orders": orders,
                    "prime_contract": "axmath_builtin_prime_v2",
                    "word_latex_rejected_reasons": [] if sane else reasons,
                    "fragment_docx": str(line.get("fragment_docx") or ""),
                }
            )

        target_p = int(planned["working_paragraph_index"])
        target_ax = [
            x for x in inv["locations"]
            if int(x["paragraph_index"]) == target_p
        ]
        if len(target_ax) != 1:
            raise ValueError(
                f"target paragraph {target_p} must contain exactly one current AxMath; "
                f"found {len(target_ax)}"
            )

        rows.append(
            {
                "repair_class": "M2_SOURCE_VISUAL_LINE_SPLIT",
                "working_paragraph_index": target_p,
                "source_paragraph_index": int(planned["source_paragraph_index"]),
                "source_ordinal": source_ordinal,
                "expected_current_axmath": 1,
                "expected_after_axmath": len(chosen),
                "expected_soft_breaks_added": len(chosen) - 1,
                "lines": chosen,
            }
        )
        expected_increase += len(chosen) - 1

    if not rows:
        raise ValueError("local visual-line split map has no rows")

    first_repair = min(int(x["working_paragraph_index"]) for x in rows)
    return {
        "schema": "axmath-repair-map/v3",
        "repair_class": "M2_SOURCE_VISUAL_LINE_SPLIT",
        "source": data.get("source"),
        "source_sha256": data.get("source_sha256_before"),
        "working": str(working),
        "working_sha256": working_sha,
        "working_axmath_count": int(inv["axmath_count"]),
        "working_omath_count": int(inv["omath_count"]),
        "working_paragraph_count": int(inv["paragraph_count"]),
        "start_anchor": plan.get("start_anchor"),
        "source_anchor_paragraph": plan.get("source_anchor_paragraph"),
        "working_anchor_paragraph": plan.get("working_anchor_paragraph"),
        "paragraph_offset": plan.get("paragraph_offset"),
        "first_repair_paragraph": first_repair,
        "visual_line_report": str(visual_report),
        "paragraph_plan": str(paragraph_plan),
        "expected_axmath_increase": expected_increase,
        "expected_final_axmath_count": int(inv["axmath_count"]) + expected_increase,
        "rows": rows,
    }

def main() -> int:
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="mode", required=True)

    p1 = sub.add_parser("multisibling")
    p1.add_argument("--source", required=True)
    p1.add_argument("--working")
    p1.add_argument("--out", required=True)

    p2 = sub.add_parser("visual-wrap")
    p2.add_argument("--visual-report", required=True)
    p2.add_argument("--working", required=True)
    p2.add_argument("--out", required=True)

    p3 = sub.add_parser("visual-line-split")
    p3.add_argument("--visual-report", required=True)
    p3.add_argument("--working", required=True)
    p3.add_argument("--paragraph-plan", required=True)
    p3.add_argument("--out", required=True)

    args = ap.parse_args()
    if args.mode == "multisibling":
        payload = build_multisibling_map(
            Path(args.source), Path(args.working) if args.working else None
        )
    elif args.mode == "visual-wrap":
        payload = build_visual_wrap_map(Path(args.visual_report), Path(args.working))
    else:
        payload = build_visual_line_split_map(
            Path(args.visual_report), Path(args.working), Path(args.paragraph_plan)
        )

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps({
        "repair_class": payload["repair_class"],
        "row_count": len(payload["rows"]),
        "working_sha256": payload.get("working_sha256"),
        "out": str(Path(args.out).resolve()),
    }, ensure_ascii=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
