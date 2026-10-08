from __future__ import annotations

import tempfile
import zipfile
from pathlib import Path

from lxml import etree as ET

from audit_docx import compare as audit_compare
from normalize_numbering_punctuation import normalize_docx, normalize_numbering_text

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _write_docx(path: Path, document_xml: str) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", document_xml.encode("utf-8"))


def _plain_paragraphs(path: Path) -> list[str]:
    ns = {"w": W}
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    return [
        "".join(node.text or "" for node in p.findall(".//w:t", ns))
        for p in root.findall(".//w:p", ns)
    ]


def _replace_document_xml(source: Path, destination: Path, old: str, new: str) -> None:
    with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(destination, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/document.xml":
                text_value = data.decode("utf-8")
                if old not in text_value:
                    raise AssertionError(f"test fixture text not found: {old!r}")
                data = text_value.replace(old, new, 1).encode("utf-8")
            zout.writestr(info, data)


def main() -> int:
    assert normalize_numbering_text("1、定义") == "1.定义"
    assert normalize_numbering_text("12、定义") == "12.定义"
    assert normalize_numbering_text("  （1）、定义") == "  （1）.定义"
    assert normalize_numbering_text("(2)、定义") == "(2).定义"
    assert normalize_numbering_text("正文\n\t（３）、子项") == "正文\n\t（３）.子项"
    assert normalize_numbering_text("第1、2项") == "第1、2项"
    assert normalize_numbering_text("甲、乙") == "甲、乙"
    assert normalize_numbering_text("正文中的（1）、不是行首") == "正文中的（1）、不是行首"
    assert normalize_numbering_text("正文中的（1）、是编号", scope="anywhere") == "正文中的（1）.是编号"
    assert normalize_numbering_text("第1、2项", scope="anywhere") == "第1.2项"
    assert normalize_numbering_text("3、第一项", scope="anywhere") == "3.第一项"

    xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}">
  <w:body>
    <w:p>
      <w:r><w:t>（</w:t></w:r>
      <w:r><w:t>1</w:t></w:r>
      <w:r><w:t>）、</w:t></w:r>
      <w:r><w:t>内容</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>普通、顿号</w:t></w:r></w:p>
    <w:p>
      <w:r><w:t>正文</w:t></w:r>
      <w:r><w:br/></w:r>
      <w:r><w:t>2</w:t></w:r>
      <w:r><w:t>、子项</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>1</w:t></w:r>
      <w:object><w:r><w:t>OBJECT</w:t></w:r></w:object>
      <w:r><w:t>、不能跨对象</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>注：（3）、子项</w:t></w:r></w:p>
  </w:body>
</w:document>
"""

    with tempfile.TemporaryDirectory() as td:
        source = Path(td) / "source.docx"
        output = Path(td) / "output.docx"
        _write_docx(source, xml)
        report = normalize_docx(source, output)
        assert report["replacements"] == 2
        assert report["structure_before"] == report["structure_after"]
        paragraphs = _plain_paragraphs(output)
        assert paragraphs[0] == "（1）.内容"
        assert paragraphs[1] == "普通、顿号"
        assert paragraphs[2] == "正文2.子项"
        # The plain-text reader sees fallback text inside w:object, but the
        # normalizer deliberately treats that object as a hard boundary.
        assert paragraphs[3] == "1OBJECT、不能跨对象"

        audit = audit_compare(source, output)
        assert audit["nonmath_text_exact"] is False
        assert audit["nonmath_text_numbering_normalized_exact"] is True
        assert audit["nonmath_text_contract_exact"] is True
        assert audit["numbering_punctuation_expected_change_count"] == 2

        # A document requiring no changes must skip ZIP recompression and
        # preserve all embedded OLE/media bytes and package metadata exactly.
        untouched = Path(td) / "no-op.docx"
        no_op = normalize_docx(output, untouched)
        assert no_op["replacements"] == 0
        assert output.read_bytes() == untouched.read_bytes()

        drift = Path(td) / "drift.docx"
        _replace_document_xml(output, drift, "普通、顿号", "普通、顿号X")
        drift_audit = audit_compare(source, drift)
        assert drift_audit["nonmath_text_contract_exact"] is False

        aggressive = Path(td) / "anywhere.docx"
        aggressive_report = normalize_docx(source, aggressive, scope="anywhere")
        assert aggressive_report["scope"] == "anywhere"
        assert audit_compare(source, aggressive, numbering_scope="anywhere")["nonmath_text_contract_exact"]
        # An aggressive normalization must not silently pass the safer default audit.
        assert not audit_compare(source, aggressive)["nonmath_text_contract_exact"]

    print("NUMBERING_PUNCTUATION_TEST_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
