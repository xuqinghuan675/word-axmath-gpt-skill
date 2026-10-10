"""Legacy Equation Editor / MathType regression without launching Word.

Run: python scripts/test_legacy_ole_gate.py
"""
from __future__ import annotations

import hashlib
import json
import tempfile
import zipfile
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

import diagnose_after_conversion
import one_click_convert
import run_skill
import strict_final_compare
from audit_docx import analyze, compare
from embedded_object_inventory import inspect_docx

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
O = "urn:schemas-microsoft-com:office:office"


def _ole(progid: str | None) -> str:
    attr = f' ProgID="{progid}"' if progid is not None else ""
    return '<w:r><w:object><o:OLEObject' + attr + '/></w:object></w:r>'


def make_docx(
    path: Path, *, native=True, axmath=False,
    extras: tuple[str | None, ...] = (), header_math=False, header_ole=None,
) -> None:
    para = '<w:r><w:t>正文</w:t></w:r>'
    if native:
        para += '<m:oMath><m:r><m:t>x+y</m:t></m:r></m:oMath>'
    if axmath:
        para += _ole("Equation.AxMath")
    para += "".join(_ole(progid) for progid in extras)
    decl = f'xmlns:w="{W}" xmlns:m="{M}" xmlns:o="{O}"'
    xml = f'<w:document {decl}><w:body><w:p>{para}</w:p></w:body></w:document>'
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", xml)
        if header_math or header_ole is not None:
            header = '<m:oMath><m:r><m:t>y</m:t></m:r></m:oMath>' if header_math else ""
            if header_ole is not None:
                header += _ole(header_ole)
            z.writestr("word/header1.xml", f'<w:hdr {decl}><w:p>{header}</w:p></w:hdr>')


def _mock_snapshots(path: Path, *_args, **_kwargs) -> dict:
    is_source = _args[-1] == "source"
    return {
        "status": "ready", "docx_stable": True,
        "pages": [{"page": 1, "path": "mock.png"}],
        "inventory": {
            "omath": [{"ordinal": 1, "paragraph_format": {}}] if is_source else [],
            "axmath": [] if is_source else [{"ordinal": 1, "paragraph_format": {}}],
        },
        "paragraph_count": 1,
    }


def main() -> int:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        src = root / "legacy_source.docx"
        make_docx(src, extras=("Equation.DSMT4", "Equation.3", "Package", None))
        inv = inspect_docx(src)
        assert inv["ole_total_all_stories"] == 4, inv
        assert inv["non_axmath_ole_count"] == 4, inv
        assert [x["kind"] for x in inv["non_axmath_ole_objects"]] == [
            "legacy_equation", "legacy_equation", "unclassified_ole", "unclassified_ole"
        ], inv
        assert all(x["paragraph_index"] == 1 for x in inv["non_axmath_ole_objects"])
        assert all(x["story_part"] == "word/document.xml" for x in inv["non_axmath_ole_objects"])
        state = one_click_convert.inspect_source(src)
        assert state["state"] == "blocked_unhandled_embedded_math", state
        assert state["non_axmath_ole_count"] == 4

        preflight_stdout = StringIO()
        with (
            mock.patch.object(one_click_convert, "environment_preflight", return_value={"ready_for_end_to_end": True}),
            mock.patch.object(one_click_convert, "word_pids", return_value=[]),
            mock.patch("sys.argv", ["one_click_convert.py", "--input", str(src)]),
            redirect_stdout(preflight_stdout),
        ):
            assert one_click_convert.main() == 1
        assert json.loads(preflight_stdout.getvalue())["status"] == "blocked_unhandled_embedded_math"
        assert not (root / (src.stem + "_AxMath-workspace")).exists(), "blocked input must not create a run"

        header_src = root / "header.docx"
        make_docx(header_src, header_math=True, header_ole="Equation.DSMT4")
        header_inv = inspect_docx(header_src)
        assert header_inv["nonmain_omath_count"] == 1, header_inv
        assert header_inv["non_axmath_ole_objects"][0]["story_part"] == "word/header1.xml"
        assert one_click_convert.inspect_source(header_src)["state"] == "blocked_unhandled_embedded_math"

        clean = root / "clean.docx"
        make_docx(clean)
        assert one_click_convert.inspect_source(clean)["state"] == "ready_officemath"
        working = root / "working.docx"
        make_docx(working, native=False, axmath=True, extras=("Equation.DSMT4",))
        audited = compare(clean, working)
        assert audited["formula_count_state"]["state"] == "exact", audited["formula_count_state"]
        assert audited["candidate_non_axmath_ole_count"] == 1

        review_dir = root / "review"
        report = diagnose_after_conversion.diagnose(clean, working, review_dir)
        assert report["status"] == "blocked_unhandled_embedded_math", report
        assert report["repair_class"] == "STOP_UNHANDLED_EMBEDDED_MATH"
        assert report["blocking_objects"]["working"][0]["progid"] == "Equation.DSMT4"

        removed_legacy = root / "removed_legacy.docx"
        make_docx(removed_legacy, native=False, axmath=True)
        report = diagnose_after_conversion.diagnose(src, removed_legacy, root / "lost-source-review")
        assert report["status"] == "blocked_unhandled_embedded_math", report
        assert report["embedded_math_gate"]["source_non_axmath_ole_count"] == 4

        output = root / "should-not-be-created.docx"
        try:
            run_skill._execute(SimpleNamespace(input=str(src), output=str(output), resume=False, overwrite_output=False))
        except RuntimeError as exc:
            assert "unhandled embedded" in str(exc)
        else:
            raise AssertionError("direct converter did not reject MathType")
        assert not output.exists(), "no output on blocked source"

        with (
            mock.patch.object(strict_final_compare, "snapshot", side_effect=_mock_snapshots),
            mock.patch.object(strict_final_compare, "page_style_report", return_value={"equal": True}),
            mock.patch.object(strict_final_compare, "_build_side_by_side", return_value=[{"page": 1, "path": "mock.png", "sha256": "fake"}]),
        ):
            source_sha = hashlib.sha256(clean.read_bytes()).hexdigest()
            failed = strict_final_compare.compare(clean, working, root / "strict-fail", source_sha)
            assert not failed["hard_content_pass"], failed
            assert failed["final_non_axmath_ole_count"] == 1

            clean_final = strict_final_compare.compare(clean, removed_legacy, root / "strict-clean", source_sha)
            assert clean_final["final_non_axmath_ole_count"] == 0
            assert clean_final["hard_content_pass"], clean_final

        assert analyze(clean)["non_axmath_ole_count"] == 0
    print("LEGACY_OLE_GATE_REGRESSION_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
