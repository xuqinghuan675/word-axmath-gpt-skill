"""Evidence-gated standalone display equation alignment repair.

Word may center a naked standalone m:oMath even when w:pPr/w:jc is 'left'.
After conversion its OLE is often left-aligned. The repair is explicit-only,
based on a source-vs-final visual finding; never auto-centers inline formulas.
Only target paragraph w:pPr/w:jc changes, in a NEW DOCX.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_URI = "http://schemas.openxmlformats.org/officeDocument/2006/math"
O_URI = "urn:schemas-microsoft-com:office:office"
NS = {"w": W_URI, "m": M_URI, "o": O_URI}
W = "{%s}" % W_URI
POST_JC = {
    W + name for name in (
        "textDirection", "textAlignment", "textboxTightWrap",
        "outlineLvl", "divId", "cnfStyle", "rPr",
    )
}


def _sha(p: Path) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _read_root(docx: Path):
    with zipfile.ZipFile(docx) as z:
        return etree.fromstring(z.read("word/document.xml"))


def _only_formula(node, *, source: bool) -> bool:
    # No nonmath plain text, other drawings or extra formulas.
    if node.findall(".//w:t", NS):
        return False
    if node.findall(".//w:drawing", NS):
        return False
    if source:
        return (
            len(node.findall("./m:oMath", NS)) == 1
            and not node.findall(".//w:object", NS)
            and not node.findall(".//m:oMathPara", NS)
        )
    return (
        not node.findall(".//m:oMath", NS)
        and len([
            x for x in node.findall(".//o:OLEObject", NS)
            if x.get("ProgID", "").lower() == "equation.axmath"
        ]) == 1
    )


def source_standalone_ordinals(source: Path) -> dict[int, int]:
    root = _read_root(Path(source))
    result = {}
    ordinal = 0
    for pi, para in enumerate(root.findall(".//w:p", NS), 1):
        nodes = para.findall(".//m:oMath", NS)
        if len(nodes) == 1 and _only_formula(para, source=True):
            result[ordinal + 1] = pi
        ordinal += len(nodes)
    return result


def repair(
    source: Path,
    working: Path,
    output: Path,
    ordinals: list[int],
    expected_source_sha256: str,
    expected_working_sha256: str,
) -> dict:
    source, working, output = (Path(p).resolve() for p in (source, working, output))
    if source == output or working == output or output.exists():
        raise ValueError("Refusing in-place or existing output alignment repair.")
    if not ordinals or any(int(x) <= 0 for x in ordinals) or len(set(ordinals)) != len(ordinals):
        raise ValueError("Pass distinct positive visually approved --ordinals.")
    if _sha(source) != expected_source_sha256.lower().strip():
        raise ValueError("Frozen source hash changed; refuse alignment repair.")
    if _sha(working) != expected_working_sha256.lower().strip():
        raise ValueError("Working DOCX hash changed; refuse alignment repair.")

    index = source_standalone_ordinals(source)
    if set(ordinals) - set(index):
        raise ValueError("Target ordinal is not a sole naked OfficeMath in its source paragraph.")
    tree = _read_root(working)
    paragraphs = tree.findall(".//w:p", NS)
    report = []
    for ordinal in ordinals:
        paragraph_index = index[ordinal]
        if paragraph_index > len(paragraphs):
            raise ValueError("Source/working paragraphs differ; refuse.")
        p = paragraphs[paragraph_index - 1]
        if not _only_formula(p, source=False):
            raise ValueError(f"Working paragraph {paragraph_index} is not a sole AxMath equation.")
        prop = p.find("w:pPr", NS)
        if prop is None:
            prop = etree.Element(W + "pPr")
            p.insert(0, prop)
        align = prop.find("w:jc", NS)
        old = align.get(W + "val") if align is not None else None
        if align is None:
            align = etree.Element(W + "jc")
            insert = next((i for i, c in enumerate(prop) if c.tag in POST_JC), len(prop))
            prop.insert(insert, align)
        align.set(W + "val", "center")
        report.append({"ordinal": ordinal, "paragraph": paragraph_index, "previous_alignment": old})

    document_xml = etree.tostring(tree, encoding="UTF-8", xml_declaration=True, standalone=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    fd, filename = tempfile.mkstemp(prefix=output.name+".", suffix=".tmp", dir=str(output.parent))
    os.close(fd)
    tmp = Path(filename)
    try:
        with zipfile.ZipFile(working) as zin, zipfile.ZipFile(tmp, "w") as zout:
            for info in zin.infolist():
                content = document_xml if info.filename == "word/document.xml" else zin.read(info.filename)
                zout.writestr(info, content)
        if not zipfile.is_zipfile(tmp):
            raise ValueError("Generated alignment repair isn't DOCX ZIP.")
        with zipfile.ZipFile(tmp) as z:
            if z.testzip() is not None:
                raise ValueError("Corrupted ZIP after alignment repair.")
            verified = etree.fromstring(z.read("word/document.xml"))
            ps = verified.findall(".//w:p", NS)
            for item in report:
                if ps[item["paragraph"]-1].find("./w:pPr/w:jc", NS).get(W+"val") != "center":
                    raise RuntimeError("Target paragraph alignment did not persist.")
            # All parts except document.xml MUST be byte-identical, including OLE.
            with zipfile.ZipFile(working) as old_zip:
                for part in z.namelist():
                    if part != "word/document.xml" and z.read(part) != old_zip.read(part):
                        raise RuntimeError(f"Unexpected non-document change in {part}.")
        if _sha(source) != expected_source_sha256.lower().strip() or _sha(working) != expected_working_sha256.lower().strip():
            raise RuntimeError("Source or input changed during repair.")
        os.rename(tmp, output)
    finally:
        if tmp.exists():
            tmp.unlink()
    return {
        "schema": "axmath-reviewed-display-alignment-repair/v1",
        "source_sha256": expected_source_sha256,
        "baseline_sha256": expected_working_sha256,
        "candidate_sha256": _sha(output),
        "source_standalone_ordinals": sorted(index),
        "reviewed_changes": report,
        "output": str(output),
        "mandatory_next_step": "Re-run diagnosis and strict per-page source-vs-final visual acceptance.",
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--working", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--ordinals", required=True, help="Comma-separated, source-visual-confirmed ordinals")
    ap.add_argument("--expected-source-sha256", required=True)
    ap.add_argument("--expected-working-sha256", required=True)
    args = ap.parse_args()
    ordinals = [int(x.strip()) for x in args.ordinals.split(",") if x.strip()]
    print(json.dumps(repair(
        Path(args.source), Path(args.working), Path(args.output),
        ordinals, args.expected_source_sha256, args.expected_working_sha256
    ), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
