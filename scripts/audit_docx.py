from __future__ import annotations

import argparse
import difflib
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

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
    paras = root.findall(".//w:p", NS)
    plain = []
    residual = []
    for pi, p in enumerate(paras, 1):
        chunks = []

        def walk(node, in_math=False, in_obj=False):
            if node.tag in (M + "oMath", M + "oMathPara"):
                in_math = True
            if node.tag == W + "object":
                in_obj = True
            if node.tag == W + "t" and not in_math and not in_obj:
                chunks.append(node.text or "")
            for child in list(node):
                walk(child, in_math, in_obj)

        walk(p)
        plain.append("".join(chunks))
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
        "embeddings": sum(
            n.startswith("word/embeddings/") and not n.endswith("/") for n in names
        ),
        "plain": plain,
        "residual": residual,
        "source_structure": structure,
    }


def compare(source: Path, candidate: Path):
    a = analyze(source)
    b = analyze(candidate)
    exact = a["plain"] == b["plain"]
    flat_exact = "".join(a["plain"]) == "".join(b["plain"])
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
        "source": {k: v for k, v in a.items() if k != "plain"},
        "candidate": {k: v for k, v in b.items() if k != "plain"},
        "paragraph_count_equal": a["paragraphs"] == b["paragraphs"],
        "nonmath_text_exact": exact,
        "nonmath_text_flat_exact": flat_exact,
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
    args = ap.parse_args()
    report = compare(Path(args.source), Path(args.candidate))
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
    print(json.dumps(report, ensure_ascii=True, indent=2))
