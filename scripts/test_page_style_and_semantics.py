from __future__ import annotations

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from restore_word_page_layout import copy_page_style, page_style_report
from source_semantic_risks import scan_source_math

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def doc_xml(*, background=False, borders=False, extra_section=False, math=False, relationship=False):
    bg = f'<w:background w:color="FFF0AA"/>' if background else ""
    border = (
        f'<w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="10"/>'
        f'<w:bottom w:val="single" w:sz="10"/></w:pgBorders>'
    ) if borders else ""
    if relationship and borders:
        border = '<w:pgBorders r:id="rIdMissing"/>'
    section = (
        f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        f'<w:pgMar w:left="1440" w:right="1440"/>{border}<w:cols w:num="1"/></w:sectPr>'
    )
    prefix = '<w:p><w:r><w:t>保留的正文</w:t></w:r></w:p>'
    if math:
        prefix += (
            f'<w:p><m:oMath><m:r><m:t>x′</m:t></m:r></m:oMath></w:p>'
            f'<w:p><m:oMath><m:r><m:t>A∪B</m:t></m:r></m:oMath></w:p>'
            f'<w:p><m:oMath><m:r><m:t>z+y</m:t></m:r></m:oMath></w:p>'
        )
    if extra_section:
        prefix += f'<w:p><w:pPr>{section}</w:pPr><w:r><w:t>中间分节</w:t></w:r></w:p>'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}" xmlns:m="{M}" xmlns:r="{REL}">'
        f'{bg}<w:body>{prefix}{section}</w:body></w:document>'
    ).encode("utf-8")


def write_doc(path: Path, **kw):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", doc_xml(**kw))
        z.writestr("word/embeddings/oleObject1.bin", b"OLE_TEST_MUST_NOT_CHANGE")
        z.writestr("word/media/image1.png", b"IMAGE_TEST_MUST_NOT_CHANGE")


def main():
    with tempfile.TemporaryDirectory() as folder:
        root = Path(folder)
        src, dst, out = (root / n for n in ("src.docx", "dst.docx", "out.docx"))
        write_doc(src, background=True, borders=True, extra_section=True, math=True)
        write_doc(dst, extra_section=True, math=True)
        before = dst.read_bytes()
        assert not page_style_report(src, dst)["equal"]
        report = copy_page_style(src, dst, out)
        assert report["status"] == "restored"
        assert page_style_report(src, out)["equal"]
        assert before == dst.read_bytes(), "page restore modified its target"
        with zipfile.ZipFile(out) as z:
            tree = etree.fromstring(z.read("word/document.xml"))
            ns = {"w": W}
            sections = tree.xpath(".//w:sectPr", namespaces=ns)
            assert len(sections) == 2
            for section in sections:
                assert [etree.QName(x).localname for x in section] == ["pgSz", "pgMar", "pgBorders", "cols"]
            assert z.read("word/embeddings/oleObject1.bin") == b"OLE_TEST_MUST_NOT_CHANGE"
        try:
            copy_page_style(src, dst, dst)
        except ValueError:
            pass
        else:
            raise AssertionError("in-place overwrite was allowed")
        try:
            copy_page_style(src, dst, out)
        except FileExistsError:
            pass
        else:
            raise AssertionError("existing output overwrite was allowed")

        mismatch = root / "sections-mismatch.docx"
        write_doc(mismatch)
        try:
            copy_page_style(src, mismatch, root / "invalid.docx")
        except ValueError:
            pass
        else:
            raise AssertionError("section mismatch did not fail closed")

        referenced = root / "referenced.docx"
        write_doc(referenced, borders=True, relationship=True)
        try:
            copy_page_style(referenced, dst, root / "bad-relationship.docx")
        except ValueError:
            pass
        else:
            raise AssertionError("unsafe media relationship not rejected")

        sem = scan_source_math(src)
        assert sem["formula_count"] == 3
        assert [x["ordinal"] for x in sem["prime_semantic_candidates"]] == [1]
        assert [x["ordinal"] for x in sem["set_symbol_visual_candidates"]] == [2]
        assert [x["ordinal"] for x in sem["standalone_display_alignment_visual_candidates"]] == [1, 2, 3]

    print("PAGE_STYLE_AND_SEMANTICS_TEST_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
