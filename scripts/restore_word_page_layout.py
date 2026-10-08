"""Restore only DOCX page background and section page borders from a source.

No paragraph runs, OMML, AxMath/OLE bytes, geometry or page setup are changed.
The destination must be new. Output is validated before atomic publication.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + NS["w"] + "}"
# CT_SectPr ordering: pgBorders follows paperSrc and precedes lnNumType.
AFTER_BORDERS = {
    W + name for name in (
        "lnNumType", "pgNumType", "cols", "formProt", "vAlign",
        "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter",
        "docGrid", "printerSettings", "sectPrChange",
    )
}


def _xml(data: bytes):
    return etree.fromstring(data, parser=etree.XMLParser(resolve_entities=False))


def _sections(root):
    # Ignore revision-history sectPr nested inside sectPrChange.
    return root.xpath(".//w:sectPr[not(ancestor::w:sectPrChange)]", namespaces=NS)


def _identity(node):
    return etree.tostring(node, method="c14n").decode("utf-8") if node is not None else None


def page_style_report(reference: Path, target: Path) -> dict:
    """Static, exact comparison of the page-decoration contract."""
    with zipfile.ZipFile(reference) as z:
        a = _xml(z.read("word/document.xml"))
    with zipfile.ZipFile(target) as z:
        b = _xml(z.read("word/document.xml"))
    sa, sb = _sections(a), _sections(b)
    a_bg, b_bg = _identity(a.find("w:background", NS)), _identity(b.find("w:background", NS))
    borders_equal = len(sa) == len(sb) and all(
        _identity(x.find("w:pgBorders", NS)) == _identity(y.find("w:pgBorders", NS))
        for x, y in zip(sa, sb)
    )
    return {
        "source_section_count": len(sa),
        "target_section_count": len(sb),
        "source_background_present": a_bg is not None,
        "target_background_present": b_bg is not None,
        "background_equal": a_bg == b_bg,
        "borders_equal": borders_equal,
        "equal": a_bg == b_bg and borders_equal,
    }


def _ensure_no_relationship(node):
    if node is not None:
        for elem in node.iter():
            for key in elem.attrib:
                if key.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"):
                    raise ValueError("Page decoration references a media relationship; copying its XML alone would break it.")


def _strip_page_style(root):
    background = root.find("w:background", NS)
    if background is not None:
        root.remove(background)
    for section in _sections(root):
        border = section.find("w:pgBorders", NS)
        if border is not None:
            section.remove(border)


def _insert_border(section, border):
    if border is None:
        return
    new = deepcopy(border)
    for i, child in enumerate(section):
        if child.tag in AFTER_BORDERS:
            section.insert(i, new)
            return
    section.append(new)


def copy_page_style(reference: Path, target: Path, output: Path) -> dict:
    reference, target, output = (
        Path(reference).resolve(), Path(target).resolve(), Path(output).resolve()
    )
    if output in (reference, target):
        raise ValueError("Refusing in-place page restoration or overwriting the reference DOCX.")
    if output.exists():
        raise FileExistsError(f"Refusing to overwrite an existing output: {output}")
    if not reference.is_file() or not target.is_file():
        raise FileNotFoundError("Reference and target DOCX must both exist.")

    source_bytes = target.read_bytes()
    input_sha = hashlib.sha256(source_bytes).hexdigest()
    with zipfile.ZipFile(reference) as src, zipfile.ZipFile(target) as dst:
        ref_doc = _xml(src.read("word/document.xml"))
        working_doc = _xml(dst.read("word/document.xml"))
        baseline = deepcopy(working_doc)
        source_sections, target_sections = _sections(ref_doc), _sections(working_doc)
        if len(source_sections) != len(target_sections):
            raise ValueError(
                f"Section count differs ({len(source_sections)} vs {len(target_sections)}); "
                "refusing silent zip() truncation or copying to wrong section."
            )
        background = ref_doc.find("w:background", NS)
        _ensure_no_relationship(background)
        previous = working_doc.find("w:background", NS)
        if previous is not None:
            working_doc.remove(previous)
        if background is not None:
            working_doc.insert(0, deepcopy(background))
        for old, current in zip(source_sections, target_sections):
            source_border = old.find("w:pgBorders", NS)
            _ensure_no_relationship(source_border)
            old_border = current.find("w:pgBorders", NS)
            if old_border is not None:
                current.remove(old_border)
            _insert_border(current, source_border)

        _strip_page_style(working_doc)
        _strip_page_style(baseline)
        if _identity(working_doc) != _identity(baseline):
            raise RuntimeError("Page restore unexpectedly changed non-page XML structure.")
        # Rebuild the decorated XML from the original target (not the stripped check).
        decorated = _xml(dst.read("word/document.xml"))
        old_bg = decorated.find("w:background", NS)
        if old_bg is not None:
            decorated.remove(old_bg)
        if background is not None:
            decorated.insert(0, deepcopy(background))
        for src_section, target_section in zip(source_sections, _sections(decorated)):
            existing = target_section.find("w:pgBorders", NS)
            if existing is not None:
                target_section.remove(existing)
            _insert_border(target_section, src_section.find("w:pgBorders", NS))
        new_xml = etree.tostring(
            decorated, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        output.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(
            prefix=output.name + ".", suffix=".tmp", dir=str(output.parent)
        )
        os.close(fd)
        tmp = Path(temporary)
        try:
            with zipfile.ZipFile(tmp, "w") as zipped:
                for info in dst.infolist():
                    zipped.writestr(
                        info,
                        new_xml if info.filename == "word/document.xml"
                        else dst.read(info.filename),
                    )
            if not zipfile.is_zipfile(tmp):
                raise ValueError("Output is not a valid DOCX ZIP package.")
            with zipfile.ZipFile(tmp) as zipped:
                if zipped.testzip() is not None:
                    raise ValueError("Generated DOCX failed ZIP CRC validation.")
            if not page_style_report(reference, tmp)["equal"]:
                raise RuntimeError("Page background/borders still differ after restore.")
            if hashlib.sha256(target.read_bytes()).hexdigest() != input_sha:
                raise RuntimeError("Target DOCX changed during restore; output not published.")
            os.rename(tmp, output)  # Windows rename refuses an existing output.
        finally:
            if tmp.exists():
                tmp.unlink()
    return {
        "status": "restored",
        "source_sha256": hashlib.sha256(reference.read_bytes()).hexdigest(),
        "target_sha256_before": input_sha,
        "target_sha256_after": hashlib.sha256(target.read_bytes()).hexdigest(),
        "output": str(output),
        "page_style": page_style_report(reference, output),
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--reference", required=True)
    ap.add_argument("--target", required=True)
    ap.add_argument("--output", required=True)
    args = ap.parse_args()
    print(json.dumps(
        copy_page_style(Path(args.reference), Path(args.target), Path(args.output)),
        ensure_ascii=False, indent=2,
    ))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
