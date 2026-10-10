from __future__ import annotations

import argparse
import difflib
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from normalize_numbering_punctuation import (
    normalize_numbering_text,
    numbering_dunhao_positions,
)
from embedded_object_inventory import scan_archive
from source_math_structure import analyze_source_structure, classify_count_state

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "o": "urn:schemas-microsoft-com:office:office",
}
W = "{%s}" % NS["w"]
M = "{%s}" % NS["m"]


def analyze(path: Path):
    path = Path(path).resolve()
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        root = ET.fromstring(z.read("word/document.xml"))
        embedded = scan_archive(z, main_root=root)
    paras = root.findall(".//w:p", NS)
    plain = []
    plain_contract = []
    residual = []
    for pi, p in enumerate(paras, 1):
        chunks = []
        contract_chunks = []

        def walk_plain(node, in_math=False, in_obj=False):
            # Preserve the pre-existing raw-text audit semantics exactly.
            if node.tag in (M + "oMath", M + "oMathPara"):
                in_math = True
            if node.tag == W + "object":
                in_obj = True
            if node.tag == W + "t" and not in_math and not in_obj:
                chunks.append(node.text or "")
            for child in list(node):
                walk_plain(child, in_math, in_obj)

        def walk_contract(node):
            if node.tag in (M + "oMath", M + "oMathPara", W + "object", W + "drawing"):
                # Match the normalizer's hard semantic boundary so a number
                # before a formula/OLE/drawing cannot become a false list label.
                contract_chunks.append("\ufffc")
                return
            if node.tag == W + "t":
                contract_chunks.append(node.text or "")
            elif node.tag in (W + "br", W + "cr"):
                contract_chunks.append("\n")
            elif node.tag == W + "tab":
                contract_chunks.append("\t")
            for child in list(node):
                walk_contract(child)

        walk_plain(p)
        walk_contract(p)
        plain.append("".join(chunks))
        plain_contract.append("".join(contract_chunks))
        for om in p.findall(".//m:oMath", NS):
            residual.append({
                "p": pi,
                "text": "".join(t.text or "" for t in om.findall(".//m:t", NS))[:120],
            })

    oles = root.findall(".//o:OLEObject", NS)
    ax = sum(
        1
        for x in oles
        if (x.attrib.get("ProgID") or "").lower() == "equation.axmath"
    )
    structure = analyze_source_structure(path)
    return {
        "paragraphs": len(paras),
        "omath": len(root.findall(".//m:oMath", NS)),
        "axmath_ole": ax,
        "ole_total": len(oles),
        "ole_progid_counts": embedded["ole_progid_counts"],
        "non_axmath_ole_count": embedded["non_axmath_ole_count"],
        "non_axmath_ole_objects": embedded["non_axmath_ole_objects"],
        "nonmain_omath_count": embedded["nonmain_omath_count"],
        "nonmain_omath_by_part": embedded["nonmain_omath_by_part"],
        "embeddings": sum(
            n.startswith("word/embeddings/") and not n.endswith("/") for n in names
        ),
        "plain": plain,
        "plain_contract": plain_contract,
        "residual": residual,
        "source_structure": structure,
    }


def compare(
    source: Path,
    candidate: Path,
    *,
    numbering_scope: str = "line-start",
    source_analysis: dict | None = None,
    candidate_analysis: dict | None = None,
):
    # A conversion runner already inventories the frozen source and working
    # DOCX; passing those immutable snapshots prevents duplicate ZIP/XML
    # scans. Callers without snapshots still get the original fresh path.
    a = source_analysis if source_analysis is not None else analyze(source)
    b = candidate_analysis if candidate_analysis is not None else analyze(candidate)
    exact = a["plain"] == b["plain"]
    flat_exact = "".join(a["plain"]) == "".join(b["plain"])
    normalized_source_contract = [
        normalize_numbering_text(text, scope=numbering_scope) for text in a["plain_contract"]
    ]
    numbering_normalized_exact = normalized_source_contract == b["plain_contract"]
    numbering_normalized_flat_exact = "".join(normalized_source_contract) == "".join(
        b["plain_contract"]
    )
    numbering_expected_change_count = sum(
        len(numbering_dunhao_positions(text, scope=numbering_scope)) for text in a["plain_contract"]
    )
    nonmath_text_contract_exact = exact or numbering_normalized_exact
    diffs = []
    if not exact:
        n = max(len(a["plain"]), len(b["plain"]))
        for i in range(n):
            x = a["plain"][i] if i < len(a["plain"]) else "<MISSING>"
            y = b["plain"][i] if i < len(b["plain"]) else "<MISSING>"
            if x != y:
                diffs.append({"p": i + 1, "source": x[:180], "candidate": y[:180]})

    ca = "".join(re.findall(r"[\u3400-\u9fff]", "".join(a["plain"])))
    cb = "".join(re.findall(r"[\u3400-\u9fff]", "".join(b["plain"])))
    count_state = classify_count_state(
        a["source_structure"], b["axmath_ole"], b["omath"]
    )
    return {
        "source": {k: v for k, v in a.items() if k not in {"plain", "plain_contract"}},
        "candidate": {k: v for k, v in b.items() if k not in {"plain", "plain_contract"}},
        "source_non_axmath_ole_count": a["non_axmath_ole_count"],
        "candidate_non_axmath_ole_count": b["non_axmath_ole_count"],
        "source_nonmain_omath_count": a["nonmain_omath_count"],
        "candidate_nonmain_omath_count": b["nonmain_omath_count"],
        "paragraph_count_equal": a["paragraphs"] == b["paragraphs"],
        "nonmath_text_exact": exact,
        "nonmath_text_flat_exact": flat_exact,
        "nonmath_text_numbering_normalized_exact": numbering_normalized_exact,
        "nonmath_text_numbering_normalized_flat_exact": numbering_normalized_flat_exact,
        "nonmath_text_contract_exact": nonmath_text_contract_exact,
        "numbering_scope": numbering_scope,
        "numbering_punctuation_expected_change_count": numbering_expected_change_count,
        "nonmath_text_diff_count": len(diffs),
        "nonmath_text_diffs": diffs[:60],
        "cjk_similarity": difflib.SequenceMatcher(
            None, ca, cb, autojunk=False
        ).ratio(),
        "formula_count_state": count_state,
        "known_multisibling_merge_signature": bool(
            count_state["known_multisibling_signature"]
        ),
        "unexpected_formula_count_gap": count_state["state"]
        in {
            "unexpected_formula_loss",
            "unexpected_formula_excess",
            "residual_officemath",
        },
    }


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--candidate", required=True)
    ap.add_argument("--out")
    ap.add_argument("--numbering-scope", choices=("line-start", "anywhere"), default="line-start")
    args = ap.parse_args()
    report = compare(Path(args.source), Path(args.candidate), numbering_scope=args.numbering_scope)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
    print(json.dumps(report, ensure_ascii=True, indent=2))
