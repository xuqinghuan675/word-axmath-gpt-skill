"""
Restore Word page-level layout from a reference DOCX without touching document body,
paragraphs, equations or AxMath objects.

Restores:
- w:background (document background)
- w:pgBorders (page borders)
- section page setup carried in w:sectPr (optional exact replacement mode)

Default: only document background + page borders. This avoids changing pagination.
"""
from __future__ import annotations

import argparse
import zipfile
from copy import deepcopy
from pathlib import Path
from lxml import etree

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def read_xml(z: zipfile.ZipFile, name: str):
    return etree.fromstring(z.read(name))


def write_zip_replace(src: Path, out: Path, replacements: dict[str, bytes]):
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = replacements.get(item.filename, zin.read(item.filename))
            zout.writestr(item, data)


def copy_page_style(reference: Path, target: Path, output: Path):
    with zipfile.ZipFile(reference, "r") as rz:
        ref_doc = read_xml(rz, "word/document.xml")
        ref_settings = None
        try:
            ref_settings = read_xml(rz, "word/settings.xml")
        except KeyError:
            pass

    with zipfile.ZipFile(target, "r") as tz:
        doc = read_xml(tz, "word/document.xml")

    # document background
    bg = ref_doc.find("w:background", NS)
    old_bg = doc.find("w:background", NS)
    if old_bg is not None:
        doc.remove(old_bg)
    if bg is not None:
        doc.insert(0, deepcopy(bg))

    # page borders in settings? Word stores them in section properties in many files,
    # copy only pgBorders from every section without changing section pagination.
    ref_sections = ref_doc.findall(".//w:sectPr", NS)
    dst_sections = doc.findall(".//w:sectPr", NS)
    for ref_sec, dst_sec in zip(ref_sections, dst_sections):
        ref_border = ref_sec.find("w:pgBorders", NS)
        old_border = dst_sec.find("w:pgBorders", NS)
        if old_border is not None:
            dst_sec.remove(old_border)
        if ref_border is not None:
            dst_sec.append(deepcopy(ref_border))

    replacements = {"word/document.xml": etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)}
    write_zip_replace(target, output, replacements)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--reference", required=True)
    ap.add_argument("--target", required=True)
    ap.add_argument("--output", required=True)
    args = ap.parse_args()
    copy_page_style(Path(args.reference), Path(args.target), Path(args.output))
    print(args.output)


if __name__ == "__main__":
    main()
