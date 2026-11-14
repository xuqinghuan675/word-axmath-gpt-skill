from __future__ import annotations

import tempfile
import zipfile
from pathlib import Path

from inspect_question_numbering import inspect

ROOT = Path(__file__).resolve().parents[1]
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _write_docx(path: Path) -> None:
    styles = f"""<w:styles xmlns:w="{W}">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:name w:val="Normal"/>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Heading2">
        <w:name w:val="heading 2"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:link w:val="Heading2Char"/>
        <w:pPr><w:outlineLvl w:val="1"/></w:pPr>
      </w:style>
      <w:style w:type="character" w:styleId="Heading2Char">
        <w:name w:val="Heading 2 Char"/>
        <w:link w:val="Heading2"/>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Heading3">
        <w:name w:val="Heading 3"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr><w:outlineLvl w:val="2"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="PlainTextStyle2">
        <w:name w:val="题目普通样式2"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr><w:outlineLvl w:val="9"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="QuestionChild">
        <w:name w:val="题目正文"/>
        <w:basedOn w:val="Heading2"/>
      </w:style>
      <w:style w:type="paragraph" w:styleId="PlainNextHeading">
        <w:name w:val="题目自定义正文"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Heading2"/>
        <w:pPr><w:outlineLvl w:val="9"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="NumberedQuestion">
        <w:name w:val="编号题目"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr>
      </w:style>
    </w:styles>"""
    document = f"""<w:document xmlns:w="{W}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Heading2"/><w:outlineLvl w:val="9"/></w:pPr><w:r><w:t>(2) heading identity</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="PlainTextStyle2"/></w:pPr><w:r><w:t>(3) plain style</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="QuestionChild"/></w:pPr><w:r><w:t>(4) based on heading</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="NumberedQuestion"/></w:pPr><w:r><w:t>auto numbered</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:instrText> SEQ Question </w:instrText><w:t>1</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Heading3"/><w:outlineLvl w:val="9"/></w:pPr><w:r><w:t>(6) heading without linked style</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="PlainNextHeading"/><w:outlineLvl w:val="9"/></w:pPr><w:r><w:t>(7) custom style next points to Heading</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Heading4"/><w:outlineLvl w:val="9"/></w:pPr><w:r><w:t>(8) built-in Heading style id without style definition</w:t></w:r></w:p>
    </w:body></w:document>"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as package:
        package.writestr("word/document.xml", document.encode("utf-8"))
        package.writestr("word/styles.xml", styles.encode("utf-8"))


def test_skill_routes_and_contract() -> None:
    skill = (ROOT / "SKILL.md").read_text(encoding="utf-8")
    layout = (ROOT / "docs" / "MATH_LAYOUT_REFERENCE.md").read_text(encoding="utf-8")
    contract = (ROOT / "docs" / "NUMBERING_AND_OUTLINE.md").read_text(encoding="utf-8")
    assert "docs/NUMBERING_AND_OUTLINE.md" in skill
    assert "NUMBERING_AND_OUTLINE.md" in layout
    for phrase in (
        "普通文字",
        "同级编号统一左缩进",
        "保留不同层级原有的视觉缩进差异",
        "w:outlineLvl",
        "PlainTextStyle1",
        "basedOn",
        "next",
        "link",
        "Word/WPS",
    ):
        assert phrase in contract, f"missing contract rule: {phrase}"


def test_inspector_detects_heading_identity_even_with_outline_nine() -> None:
    with tempfile.TemporaryDirectory() as temp:
        docx = Path(temp) / "numbering.docx"
        _write_docx(docx)
        report = inspect(docx)
        candidates = {item["paragraph"]: item for item in report["candidates"]}

        assert 1 in candidates
        heading = candidates[1]
        assert heading["outline_level"] == "9"
        assert heading["heading_style_candidate"] is True
        assert heading["heading_style_reference_candidate"] is True
        assert any(
            relation["relationship"] == "link"
            and relation["target_style"] == "Heading2Char"
            for relation in heading["style_relationships"]
        )

        assert 2 not in candidates
        assert 3 in candidates and candidates[3]["heading_style_candidate"] is True
        assert candidates[4]["automatic_numbering"] is True
        assert candidates[5]["numbering_fields"] == ["SEQ"]
        # Heading identity must still be detected with outlineLvl=9 and no link.
        assert 6 in candidates
        assert candidates[6]["outline_level"] == "9"
        assert candidates[6]["heading_style_candidate"] is True
        assert candidates[6]["heading_style_reference_candidate"] is False
        assert 7 in candidates and candidates[7]["heading_style_reference_candidate"] is True
        assert any(
            relation["relationship"] == "next"
            and relation["target_style"] == "Heading2"
            for relation in candidates[7]["style_relationships"]
        )
        assert 8 in candidates and candidates[8]["heading_style_candidate"] is True


if __name__ == "__main__":
    test_skill_routes_and_contract()
    test_inspector_detects_heading_identity_even_with_outline_nine()
    print("NUMBERING_CONTRACT_OK")
