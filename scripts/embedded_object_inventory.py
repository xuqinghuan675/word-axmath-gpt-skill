"""Read-only inventory of DOCX formula containers and legacy OLE objects.

OfficeMath is stored as m:oMath XML; AxMath/MathType/Equation Editor are
embedded OLE objects.  Counting only the first two silently misses legacy
equations (e.g. Equation.DSMT4), so keep every non-AxMath OLE visible until
it is individually classified and handled on a new document copy.
"""
from __future__ import annotations

import collections
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
O = "{urn:schemas-microsoft-com:office:office}"

_STORY = re.compile(
    r"word/(?:document|footnotes|endnotes|comments|header[0-9]+|footer[0-9]+)\.xml$"
)


def _kind(progid: str) -> str:
    name = progid.casefold().strip()
    if name == "equation.axmath":
        return "axmath"
    if (name.startswith("equation.") or name == "equation"
            or name.startswith("mathtype.") or name == "mathtype"):
        return "legacy_equation"
    return "unclassified_ole"


def scan_archive(z: zipfile.ZipFile, main_root: ET.Element | None = None) -> dict:
    """Enumerate active equation/OLE markup in all standard Word story XML.

    Unknown embedded objects are not assumed to be mathematical, but cannot
    silently pass an all-formulas-AxMath claim. No OLE payloads are activated.
    """
    parts = sorted(
        (name for name in z.namelist() if _STORY.fullmatch(name)),
        key=lambda name: (name != "word/document.xml", name),
    )
    if "word/document.xml" not in parts:
        raise ValueError("DOCX has no word/document.xml")

    progids: collections.Counter[str] = collections.Counter()
    foreign: list[dict] = []
    nonmain_omath_by_part: dict[str, int] = {}
    object_ordinal = 0
    axmath_count = 0

    for part in parts:
        root = main_root if part == "word/document.xml" and main_root is not None else ET.fromstring(z.read(part))
        paragraphs = root.findall(".//" + W + "p")
        para_by_ole = {
            id(ole): pi
            for pi, paragraph in enumerate(paragraphs, 1)
            for ole in paragraph.findall(".//" + O + "OLEObject")
        }
        omath_count = len(root.findall(".//" + M + "oMath"))
        if part != "word/document.xml" and omath_count:
            nonmain_omath_by_part[part] = omath_count

        for ole in root.findall(".//" + O + "OLEObject"):
            object_ordinal += 1
            progid = (ole.get("ProgID") or "").strip()
            kind = _kind(progid)
            progids[progid or "<missing ProgID>"] += 1
            if kind == "axmath":
                axmath_count += 1
            else:
                foreign.append({
                    "object_ordinal": object_ordinal,
                    "story_part": part,
                    "paragraph_index": para_by_ole.get(id(ole)),
                    "progid": progid,
                    "kind": kind,
                })

    return {
        "schema": "axmath-embedded-object-inventory/v1",
        "ole_total_all_stories": object_ordinal,
        "axmath_ole_all_stories": axmath_count,
        "ole_progid_counts": dict(sorted(progids.items())),
        "non_axmath_ole_count": len(foreign),
        "non_axmath_ole_objects": foreign,
        "nonmain_omath_count": sum(nonmain_omath_by_part.values()),
        "nonmain_omath_by_part": nonmain_omath_by_part,
    }


def inspect_docx(path: Path) -> dict:
    with zipfile.ZipFile(path) as z:
        return scan_archive(z)


def main() -> int:
    import argparse
    import sys

    ap = argparse.ArgumentParser(description="Inspect embedded Word/AxMath/legacy equation objects without opening Word.")
    ap.add_argument("docx")
    args = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass
    report = inspect_docx(Path(args.docx))
    report["source"] = str(Path(args.docx).resolve())
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if not report["non_axmath_ole_count"] and not report["nonmain_omath_count"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
