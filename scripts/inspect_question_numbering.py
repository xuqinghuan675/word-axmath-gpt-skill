"""Read-only candidates for automatic numbering, fields and outline headings.

Candidates require review in Word: not every list/field/heading is a question,
and package inspection alone cannot prove the displayed question is correct.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
FIELD = re.compile(r"\b(?:LISTNUM|SEQ|AUTONUM|AUTONUMLGL|AUTONUMOUT)\b", re.I)
STORY = re.compile(r"^word/(?:document|footnotes|endnotes|comments|header\d+|footer\d+)\.xml$")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def val(element: ET.Element | None) -> str | None:
    return element.get(W + "val") if element is not None else None


def number_state(properties: ET.Element | None) -> str | None:
    if properties is None:
        return None
    number = properties.find(W + "numPr")
    if number is None:
        return None
    # A directly applied numId=0 disables an inherited list. Anything else,
    # including an incomplete numPr, is a candidate for Word review.
    return "disabled" if val(number.find(W + "numId")) == "0" else "candidate"


def inspect(docx: Path) -> dict:
    docx = docx.resolve(strict=True)
    before = sha256(docx)
    styles: dict[str, ET.Element] = {}
    default_style = None
    findings = []
    with zipfile.ZipFile(docx) as package:
        names = package.namelist()
        if "word/document.xml" not in names:
            raise ValueError("Not a readable Word document: word/document.xml is missing")
        if "word/styles.xml" in names:
            root = ET.fromstring(package.read("word/styles.xml"))
            for style in root.findall(W + "style"):
                if style.get(W + "type") != "paragraph":
                    continue
                identifier = style.get(W + "styleId")
                if identifier:
                    styles[identifier] = style
                    if style.get(W + "default") in ("1", "true", "on"):
                        default_style = identifier

        def ancestry(identifier: str | None) -> list[tuple[str, ET.Element]]:
            chain = []
            seen = set()
            while identifier and identifier not in seen and identifier in styles:
                seen.add(identifier)
                style = styles[identifier]
                chain.append((identifier, style))
                identifier = val(style.find(W + "basedOn"))
            return chain

        for name in sorted(n for n in names if STORY.fullmatch(n)):
            root = ET.fromstring(package.read(name))
            for index, paragraph in enumerate(root.iter(W + "p"), 1):
                properties = paragraph.find(W + "pPr")
                style_id = val(properties.find(W + "pStyle")) if properties is not None else None
                style_id = style_id or default_style
                chain = ancestry(style_id)
                state = number_state(properties)
                origin = "paragraph" if state is not None else None
                if state is None:
                    for inherited_id, style in chain:
                        state = number_state(style.find(W + "pPr"))
                        if state is not None:
                            origin = "style:" + inherited_id
                            break
                outline = val(properties.find(W + "outlineLvl")) if properties is not None else None
                if outline is None:
                    for _, style in chain:
                        style_properties = style.find(W + "pPr")
                        outline = val(style_properties.find(W + "outlineLvl")) if style_properties is not None else None
                        if outline is not None:
                            break
                heading_style = any(re.fullmatch(r"(?:Heading|标题)[1-9]", item[0], re.I) for item in chain)
                if not chain and style_id:
                    heading_style = bool(re.fullmatch(r"(?:Heading|标题)[1-9]", style_id, re.I))
                instructions = " ".join(
                    ["".join(x.text or "" for x in paragraph.iter(W + "instrText"))]
                    + [x.get(W + "instr", "") for x in paragraph.iter(W + "fldSimple")]
                )
                fields = sorted({match.group(0).upper() for match in FIELD.finditer(instructions)})
                outline_candidate = (outline is not None and outline != "9") or (outline is None and heading_style)
                if state == "candidate" or fields or outline_candidate:
                    findings.append({
                        "story": name,
                        "paragraph": index,
                        "style": style_id,
                        "automatic_numbering": state == "candidate",
                        "numbering_origin": origin if state == "candidate" else None,
                        "numbering_fields": fields,
                        "outline_level": outline,
                        "heading_style_candidate": heading_style,
                        "text_preview": "".join(x.text or "" for x in paragraph.iter(W + "t"))[:100],
                    })
    after = sha256(docx)
    if before != after:
        raise RuntimeError("The document changed during inspection; repeat on a frozen copy.")
    return {
        "schema": "axmath-question-numbering-candidates/v1",
        "docx": str(docx),
        "sha256": before,
        "candidates": findings,
        "styles_available": "word/styles.xml" in names,
        "note": "Review each question in Word; a candidate is not automatically an error and no candidates is not visual proof.",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx")
    args = parser.parse_args()
    print(json.dumps(inspect(Path(args.docx)), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
