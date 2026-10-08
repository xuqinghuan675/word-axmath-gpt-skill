"""Fast static OMML risk inventory; never rewrites formula semantics.

Flags derivative prime tokens and collection operators for targeted review.
A source marker is only a risk signal, never evidence that AxMath is broken.
"""
from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from normalize_axmath_tex import contains_prime_or_derivative_marker

NS = {
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
M = "{" + NS["m"] + "}"
SET_TOKENS = set("∪∩⊂⊆⊃⊇∈∉∅∖⋃⋂")


def scan_source_math(path: Path) -> dict:
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    math_nodes = root.findall(".//m:oMath", NS)
    primes, sets, standalone = [], [], []
    for ordinal, formula in enumerate(math_nodes, 1):
        text = "".join((node.text or "") for node in formula.findall(".//m:t", NS))
        # OMML may carry operators/accents in attribute values, not m:t text.
        chars = [
            node.get(M + "val") or ""
            for node in formula.iter()
            if node.tag == M + "chr"
        ]
        inspected = text + "".join(chars)
        if contains_prime_or_derivative_marker(inspected):
            primes.append({
                "ordinal": ordinal,
                "source_text": text[:180],
                "prime_symbols": sorted({c for c in inspected if c in "'′″‴⁗’"}),
                "reason": "source_prime_semantics_require_axmath_contract_review",
            })
        tokens = sorted(set(inspected) & SET_TOKENS)
        if tokens:
            sets.append({
                "ordinal": ordinal,
                "source_text": text[:180],
                "symbols": tokens,
                "reason": "collection_symbol_size_and_glyph_visual_review",
            })
    # Word may implicitly center naked standalone OMML while the resulting OLE
    # follows the paragraph's left alignment. COM x_pt can misleadingly match.
    ordinal = 0
    for pi, paragraph in enumerate(root.findall(".//w:p", NS), 1):
        equations = paragraph.findall(".//m:oMath", NS)
        if (
            len(equations) == 1
            and len(paragraph.findall("./m:oMath", NS)) == 1
            and not paragraph.findall(".//m:oMathPara", NS)
            and not paragraph.findall(".//w:t", NS)
            and not paragraph.findall(".//w:object", NS)
        ):
            standalone.append({
                "ordinal": ordinal + 1,
                "paragraph_index": pi,
                "reason": "standalone_native_math_can_center_while_axmath_ole_is_left",
            })
        ordinal += len(equations)
    return {
        "schema": "axmath-source-semantic-risks/v1",
        "source": str(Path(path).resolve()),
        "formula_count": len(math_nodes),
        "prime_semantic_candidates": primes,
        "set_symbol_visual_candidates": sets,
        "standalone_display_alignment_visual_candidates": standalone,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True)
    parser.add_argument("--out")
    args = parser.parse_args()
    result = scan_source_math(Path(args.source))
    data = json.dumps(result, ensure_ascii=False, indent=2)
    if args.out:
        Path(args.out).write_text(data, encoding="utf-8")
    print(data)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
