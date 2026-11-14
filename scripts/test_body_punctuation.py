from __future__ import annotations

import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))
from normalize_body_punctuation import normalize_body_punctuation

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W_NS, "m": M_NS}


def _make_docx(path: Path) -> None:
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}" xmlns:m="{M_NS}">
  <w:body>
    <w:p><w:r><w:t>正文。句号</w:t></w:r></w:p>
    <w:p><m:oMath><m:r><m:t>公式。</m:t></m:r></m:oMath></w:p>
    <w:p><w:object><w:t>OLE。fallback</w:t></w:object></w:p>
    <w:p><w:drawing><w:t>图中文字。</w:t></w:drawing></w:p>
    <w:p><w:fldSimple><w:r><w:t>字段结果。</w:t></w:r></w:fldSimple></w:p>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText> PAGE </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>复杂字段结果。</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:body>
</w:document>'''
    header = f'''<?xml version="1.0" encoding="UTF-8"?>
<w:hdr xmlns:w="{W_NS}"><w:p><w:r><w:t>页眉。</w:t></w:r></w:p></w:hdr>'''
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", document.encode("utf-8"))
        archive.writestr("word/header1.xml", header.encode("utf-8"))
        archive.writestr("word/media/image1.png", b"image-payload")
        archive.writestr("word/embeddings/oleObject1.bin", b"ole-payload")


def test_only_ordinary_main_body_text_is_changed() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        source = root / "source.docx"
        output = root / "output.docx"
        report = root / "report.json"
        _make_docx(source)
        before = source.read_bytes()
        state = normalize_body_punctuation(source, output, report)

        assert state["replaced_count"] == 1
        assert state["remaining_ordinary_body_count"] == 0
        assert source.read_bytes() == before

        with zipfile.ZipFile(source) as src, zipfile.ZipFile(output) as dst:
            assert src.namelist() == dst.namelist()
            for name in src.namelist():
                if name != "word/document.xml":
                    assert src.read(name) == dst.read(name), name
            root_xml = ET.fromstring(dst.read("word/document.xml"))
            ordinary = root_xml.find(".//w:p/w:r/w:t", NS)
            assert ordinary is not None and ordinary.text == "正文.句号"
            math_text = root_xml.find(".//m:t", NS)
            assert math_text is not None and math_text.text == "公式。"
            ole_text = root_xml.find(".//w:object/w:t", NS)
            assert ole_text is not None and ole_text.text == "OLE。fallback"
            drawing_text = root_xml.find(".//w:drawing/w:t", NS)
            assert drawing_text is not None and drawing_text.text == "图中文字。"
            field_text = root_xml.find(".//w:fldSimple//w:t", NS)
            assert field_text is not None and field_text.text == "字段结果。"
            complex_field_paragraphs = [
                paragraph for paragraph in root_xml.findall(".//w:p", NS)
                if paragraph.find(".//w:fldChar", NS) is not None
            ]
            complex_field_texts = [
                node.text for paragraph in complex_field_paragraphs
                for node in paragraph.findall(".//w:t", NS)
            ]
            assert "复杂字段结果。" in complex_field_texts


if __name__ == "__main__":
    test_only_ordinary_main_body_text_is_changed()
    print("BODY_PUNCTUATION_TESTS_OK")
