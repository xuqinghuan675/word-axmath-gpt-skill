"""Read-only candidates for automatic numbering, fields and outline headings.

The inspector also reports built-in Heading style identity and style relationships.
WPS may show a collapse arrow for a Heading style even when a paragraph has a
direct outlineLvl=9 override; such paragraphs remain candidates for review.
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
HEADING_IDENTITY = re.compile(r"^(?:Heading|标题)\s*[1-9](?:\s*Char)?$", re.I)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def val(element: ET.Element | None) -> str | None:
    return element.get(W + "val") if element is not None else None


def style_name(style: ET.Element | None) -> str | None:
    return val(style.find(W + "name")) if style is not None else None


def is_heading_identity(identifier: str | None, style: ET.Element | None) -> bool:
    candidates = [identifier or "", style_name(style) or ""]
    return any(HEADING_IDENTITY.fullmatch(candidate.strip()) for candidate in candidates)


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
    paragraph_styles: dict[str, ET.Element] = {}
    all_styles: dict[str, ET.Element] = {}
    default_style = None
    findings = []

    with zipfile.ZipFile(docx) as package:
        names = package.namelist()
        if "word/document.xml" not in names:
            raise ValueError("Not a readable Word document: word/document.xml is missing")

        if "word/styles.xml" in names:
            style_root = ET.fromstring(package.read("word/styles.xml"))
            for style in style_root.findall(W + "style"):
                identifier = style.get(W + "styleId")
                if not identifier:
                    continue
                all_styles[identifier] = style
                if style.get(W + "type") == "paragraph":
                    paragraph_styles[identifier] = style
                    if style.get(W + "default") in ("1", "true", "on"):
                        default_style = identifier

        def ancestry(identifier: str | None) -> list[tuple[str, ET.Element]]:
            chain = []
            seen = set()
            while identifier and identifier not in seen and identifier in paragraph_styles:
                seen.add(identifier)
                style = paragraph_styles[identifier]
                chain.append((identifier, style))
                identifier = val(style.find(W + "basedOn"))
            return chain

        for name in sorted(n for n in names if STORY.fullmatch(n)):
            story_root = ET.fromstring(package.read(name))
            for index, paragraph in enumerate(story_root.iter(W + "p"), 1):
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
                outline_source = "paragraph" if outline is not None else None
                if outline is None:
                    for inherited_id, style in chain:
                        style_properties = style.find(W + "pPr")
                        inherited_outline = (
                            val(style_properties.find(W + "outlineLvl"))
                            if style_properties is not None else None
                        )
                        if inherited_outline is not None:
                            outline = inherited_outline
                            outline_source = "style:" + inherited_id
                            break

                heading_style = (
                    is_heading_identity(style_id, all_styles.get(style_id))
                    or any(
                        is_heading_identity(item_id, item_style)
                        for item_id, item_style in chain
                    )
                )

                style_chain = []
                style_relationships = []
                heading_relationships = []
                for chain_id, chain_style in chain:
                    chain_properties = chain_style.find(W + "pPr")
                    style_chain.append({
                        "style_id": chain_id,
                        "name": style_name(chain_style),
                        "based_on": val(chain_style.find(W + "basedOn")),
                        "next": val(chain_style.find(W + "next")),
                        "link": val(chain_style.find(W + "link")),
                        "outline_level": (
                            val(chain_properties.find(W + "outlineLvl"))
                            if chain_properties is not None else None
                        ),
                        "numbering": number_state(chain_properties),
                    })
                    for relationship in ("basedOn", "next", "link"):
                        target_id = val(chain_style.find(W + relationship))
                        if not target_id:
                            continue
                        target_style = all_styles.get(target_id)
                        target_info = {
                            "from_style": chain_id,
                            "relationship": relationship,
                            "target_style": target_id,
                            "target_name": style_name(target_style),
                            "target_is_heading": is_heading_identity(target_id, target_style),
                        }
                        style_relationships.append(target_info)
                        if target_info["target_is_heading"]:
                            heading_relationships.append(target_info)

                instructions = " ".join(
                    ["".join(x.text or "" for x in paragraph.iter(W + "instrText"))]
                    + [x.get(W + "instr", "") for x in paragraph.iter(W + "fldSimple")]
                )
                fields = sorted({match.group(0).upper() for match in FIELD.finditer(instructions)})
                outline_candidate = (outline is not None and outline != "9") or (
                    outline is None and heading_style
                )
                heading_reference_candidate = bool(heading_relationships)

                if state == "candidate" or fields or outline_candidate or heading_style or heading_reference_candidate:
                    findings.append({
                        "story": name,
                        "paragraph": index,
                        "style": style_id,
                        "style_name": style_name(all_styles.get(style_id)),
                        "automatic_numbering": state == "candidate",
                        "numbering_origin": origin if state == "candidate" else None,
                        "numbering_fields": fields,
                        "outline_level": outline,
                        "outline_source": outline_source,
                        "heading_style_candidate": heading_style,
                        "heading_style_reference_candidate": heading_reference_candidate,
                        "style_chain": style_chain,
                        "style_relationships": style_relationships,
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
        "note": (
            "Read-only candidates require review. WPS may retain a collapse arrow from "
            "built-in Heading style identity even when outline_level is 9; inspect style "
            "identity and basedOn/next/link references, then verify in Word and WPS."
        ),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx")
    args = parser.parse_args()
    print(json.dumps(inspect(Path(args.docx)), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
