from __future__ import annotations

import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}

def _math_text(node: ET.Element, limit: int = 240) -> str:
    return "".join(t.text or "" for t in node.findall(".//m:t", NS))[:limit]


def analyze_source_structure(docx: Path) -> dict:
    """Return static OMML structure needed by the repair state machine.

    Word exposes every m:oMath through OMaths, including multiple direct
    m:oMath children inside one m:oMathPara.  AxMath batch conversion can
    collapse those siblings into one OLE.  Record the groups before conversion
    so the later count gap is classified instead of guessed at.
    """
    docx = Path(docx).resolve()
    with zipfile.ZipFile(docx) as z:
        root = ET.fromstring(z.read("word/document.xml"))

    all_omath = root.findall(".//m:oMath", NS)
    ordinal_by_id = {id(node): i for i, node in enumerate(all_omath, 1)}
    paragraphs = root.findall(".//w:p", NS)

    groups: list[dict] = []
    grouped_ids: set[int] = set()
    display_count = 0

    for paragraph_index, paragraph in enumerate(paragraphs, 1):
        for omath_para in paragraph.findall(".//m:oMathPara", NS):
            direct = omath_para.findall("./m:oMath", NS)
            if not direct:
                continue
            display_count += 1
            grouped_ids.update(id(x) for x in direct)
            ordinals = [ordinal_by_id[id(x)] for x in direct]
            if len(direct) > 1:
                groups.append({
                    "paragraph_index": paragraph_index,
                    "source_ordinals": ordinals,
                    "source_count": len(direct),
                    "extra_raw_nodes": len(direct) - 1,
                    "source_texts": [_math_text(x) for x in direct],
                    "repair_class": "M1_MULTISIBLING_OMATHPARA",
                })

    inline_ordinals = [
        ordinal_by_id[id(node)] for node in all_omath if id(node) not in grouped_ids
    ]
    extra = sum(int(x["extra_raw_nodes"]) for x in groups)
    raw = len(all_omath)

    return {
        "schema": "axmath-source-structure/v1",
        "source": str(docx),
        "paragraph_count": len(paragraphs),
        "raw_omath_count": raw,
        "display_omathpara_count": display_count,
        "inline_or_unwrapped_omath_count": len(inline_ordinals),
        "multi_sibling_group_count": len(groups),
        "multi_sibling_extra_nodes": extra,
        "batch_collapse_signature_axmath_count": raw - extra,
        "multi_sibling_groups": groups,
    }


def classify_count_state(
    source_structure: dict, candidate_axmath: int, candidate_omath: int
) -> dict:
    raw = int(source_structure.get("raw_omath_count") or 0)
    extra = int(source_structure.get("multi_sibling_extra_nodes") or 0)
    ax = int(candidate_axmath or 0)
    om = int(candidate_omath or 0)
    gap = raw - ax

    if om:
        state = "residual_officemath"
        repair_class = "STOP_RESIDUAL_OFFICEMATH"
    elif ax == raw:
        state = "exact"
        repair_class = None
    elif extra > 0 and gap == extra:
        state = "known_multisibling_collapse"
        repair_class = "M1_MULTISIBLING_OMATHPARA"
    elif ax < raw:
        state = "unexpected_formula_loss"
        repair_class = "STOP_UNEXPLAINED_COUNT_GAP"
    else:
        state = "unexpected_formula_excess"
        repair_class = "STOP_UNEXPLAINED_COUNT_GAP"

    return {
        "state": state,
        "repair_class": repair_class,
        "source_raw_omath": raw,
        "candidate_axmath": ax,
        "candidate_omath": om,
        "raw_count_gap": gap,
        "known_multisibling_extra_nodes": extra,
        "known_multisibling_signature": state == "known_multisibling_collapse",
    }


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--candidate-axmath", type=int)
    ap.add_argument("--candidate-omath", type=int, default=0)
    ap.add_argument("--out")
    args = ap.parse_args()

    report = analyze_source_structure(Path(args.docx))
    if args.candidate_axmath is not None:
        report["count_state"] = classify_count_state(
            report, args.candidate_axmath, args.candidate_omath
        )
    text_value = json.dumps(report, ensure_ascii=False, indent=2)
    if args.out:
        out = Path(args.out)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text_value, encoding="utf-8")
    print(json.dumps(report, ensure_ascii=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
