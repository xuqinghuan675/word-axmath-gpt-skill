from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Any

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "o": "urn:schemas-microsoft-com:office:office",
}
W = "{%s}" % NS["w"]
M = "{%s}" % NS["m"]

# Only literal list/sub-item labels are normalized. A label must be at the
# beginning of a paragraph or immediately after a Word line break; ordinary
# prose such as "第1、2项" or "甲、乙" is intentionally outside this rule.
_NUMBERING_DUNHAO_RE = re.compile(
    r"(?m)^[ \t\u3000]*(?:[0-9０-９]+|[（(][0-9０-９]+[)）])、"
)


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def numbering_dunhao_positions(text: str) -> list[int]:
    """Return character offsets of label-final Chinese dunhao characters."""
    return [match.end() - 1 for match in _NUMBERING_DUNHAO_RE.finditer(text or "")]


def normalize_numbering_text(text: str) -> str:
    positions = numbering_dunhao_positions(text)
    if not positions:
        return text
    chars = list(text)
    for pos in positions:
        chars[pos] = "."
    return "".join(chars)


def _paragraph_projection(paragraph: Any) -> tuple[str, list[tuple[Any, int] | None]]:
    chars: list[str] = []
    mapping: list[tuple[Any, int] | None] = []

    def append_boundary(value: str) -> None:
        chars.append(value)
        mapping.append(None)

    def walk(node: Any) -> None:
        # Do not let numbering recognition cross formulas/OLE/drawings. They are
        # semantic object boundaries, not plain text.
        if node.tag in {W + "object", W + "drawing", M + "oMath", M + "oMathPara"}:
            append_boundary("\ufffc")
            return
        if node.tag == W + "t":
            value = node.text or ""
            for offset, ch in enumerate(value):
                chars.append(ch)
                mapping.append((node, offset))
            return
        if node.tag in {W + "br", W + "cr"}:
            append_boundary("\n")
            return
        if node.tag == W + "tab":
            append_boundary("\t")
            return
        for child in node:
            walk(child)

    walk(paragraph)
    return "".join(chars), mapping


def _replace_in_paragraph(paragraph: Any) -> int:
    text, mapping = _paragraph_projection(paragraph)
    positions = numbering_dunhao_positions(text)
    if not positions:
        return 0

    by_node: dict[Any, list[int]] = {}
    for pos in positions:
        target = mapping[pos]
        if target is None:
            raise RuntimeError("numbering punctuation target is not backed by a w:t node")
        node, offset = target
        by_node.setdefault(node, []).append(offset)

    changed = 0
    for node, offsets in by_node.items():
        value = list(node.text or "")
        for offset in offsets:
            if value[offset] != "、":
                raise RuntimeError("numbering punctuation projection drifted before mutation")
            value[offset] = "."
            changed += 1
        node.text = "".join(value)
    return changed


def _document_counts(root: Any) -> dict[str, int]:
    oles = root.findall(".//o:OLEObject", NS)
    return {
        "paragraphs": len(root.findall(".//w:p", NS)),
        "omath": len(root.findall(".//m:oMath", NS)),
        "axmath_ole": sum(
            1
            for obj in oles
            if (obj.attrib.get("ProgID") or "").lower() == "equation.axmath"
        ),
    }


def normalize_docx(input_docx: Path, output_docx: Path) -> dict:
    # Keep lxml lazy so one_click_convert.py --inspect-only can still report a
    # missing lxml dependency through environment_preflight instead of failing
    # at module import time.
    from lxml import etree as ET

    input_docx = Path(input_docx).resolve()
    output_docx = Path(output_docx).resolve()
    if input_docx == output_docx:
        raise ValueError("Refusing in-place numbering punctuation normalization.")
    if not input_docx.is_file():
        raise FileNotFoundError(input_docx)

    parser = ET.XMLParser(remove_blank_text=False, resolve_entities=False)
    with zipfile.ZipFile(input_docx, "r") as zin:
        original_xml = zin.read("word/document.xml")
        root = ET.fromstring(original_xml, parser=parser)
        before = _document_counts(root)
        replacements = sum(_replace_in_paragraph(p) for p in root.findall(".//w:p", NS))
        after = _document_counts(root)
        if before != after:
            raise RuntimeError(
                f"numbering normalization changed document structure: {before} -> {after}"
            )
        new_xml = ET.tostring(
            root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

        output_docx.parent.mkdir(parents=True, exist_ok=True)
        fd, temp_name = tempfile.mkstemp(
            prefix=output_docx.name + ".",
            suffix=".tmp",
            dir=str(output_docx.parent),
        )
        os.close(fd)
        temp_path = Path(temp_name)
        try:
            with zipfile.ZipFile(temp_path, "w") as zout:
                for info in zin.infolist():
                    data = new_xml if info.filename == "word/document.xml" else zin.read(info.filename)
                    zout.writestr(info, data)
            os.replace(temp_path, output_docx)
        finally:
            if temp_path.exists():
                temp_path.unlink()

    return {
        "schema": "axmath-numbering-punctuation-normalization/v1",
        "input": str(input_docx),
        "output": str(output_docx),
        "input_sha256": sha256_file(input_docx),
        "output_sha256": sha256_file(output_docx),
        "replacements": replacements,
        "rule": "paragraph-or-line-start Arabic-number label: 、 -> .",
        "supports": ["1、", "12、", "（1）、", "(1)、", "fullwidth Arabic digits"],
        "structure_before": before,
        "structure_after": after,
    }


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Normalize Arabic-number list label punctuation in DOCX plain text."
    )
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--report")
    args = ap.parse_args()

    report = normalize_docx(Path(args.input), Path(args.output))
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        Path(args.report).write_text(text, encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
