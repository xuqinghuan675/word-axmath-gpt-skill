"""Standalone Word OMML -> AxMath display repair safety contract."""
from __future__ import annotations

import hashlib
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from repair_standalone_display_alignment import repair, source_standalone_ordinals
from source_semantic_risks import scan_source_math

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
O = "urn:schemas-microsoft-com:office:office"


def write_docx(path, math):
    body = (
        f'<m:oMath><m:r><m:t>x+y</m:t></m:r></m:oMath>'
        if math else '<w:object><o:OLEObject ProgID="Equation.AxMath"/></w:object>'
    )
    xml = (
        f'<w:document xmlns:w="{W}" xmlns:m="{M}" xmlns:o="{O}"><w:body>'
        '<w:p><w:r><w:t>Intro</w:t></w:r></w:p>'
        f'<w:p>{body}</w:p>'
        '<w:sectPr><w:pgSz w:w="11906"/></w:sectPr>'
        '</w:body></w:document>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", xml.encode("utf-8"))
        z.writestr("word/embeddings/oleObject1.bin", b"PREEXISTING_OLE_BYTE_CONTRACT")


def main():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        src, working, repaired = (root / a for a in ("src.docx", "work.docx", "fixed.docx"))
        write_docx(src, math=True)
        write_docx(working, math=False)
        assert source_standalone_ordinals(src) == {1: 2}
        sem = scan_source_math(src)
        assert [x["ordinal"] for x in sem["standalone_display_alignment_visual_candidates"]] == [1]
        sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
        old_src, old_work = sha(src), sha(working)
        report = repair(src, working, repaired, [1], old_src, old_work)
        assert report["reviewed_changes"][0]["paragraph"] == 2
        assert sha(src) == old_src and sha(working) == old_work
        with zipfile.ZipFile(repaired) as z:
            doc = etree.fromstring(z.read("word/document.xml"))
            ns = {"w": W}
            assert doc.xpath(".//w:p[2]/w:pPr/w:jc/@w:val", namespaces=ns) == ["center"]
            assert z.read("word/embeddings/oleObject1.bin") == b"PREEXISTING_OLE_BYTE_CONTRACT"
        for ords, sh in [([2], old_src), ([1], "0" * 64)]:
            try:
                repair(src, working, root / ("invalid-" + str(ords[0]) + ".docx"), ords, sh, old_work)
            except ValueError:
                pass
            else:
                raise AssertionError("Unsafe standalone alignment repair was permitted")
        try:
            repair(src, working, repaired, [1], old_src, old_work)
        except ValueError:
            pass
        else:
            raise AssertionError("Existing finished DOCX was overwritten")
    print("DISPLAY_ALIGNMENT_REPAIR_TEST_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
