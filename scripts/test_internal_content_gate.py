"""Regression: equal OLE counts cannot approve an empty editable AxMath object.

The export reports are fixtures; real AMSAM2TeX remains a Word/AxMath smoke gate.
"""
from __future__ import annotations

import json
import tempfile
import zipfile
from pathlib import Path
from unittest import mock

from audit_docx import compare as content_compare
from class_e_review import prepare as prepare_class_e_review, validate as validate_class_e_review
from internal_content_audit import sha256_file, validate_full_internal_audit
import diagnose_after_conversion
import finalize_visual_review
import strict_final_compare

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
O = "urn:schemas-microsoft-com:office:office"


def make_docx(path: Path, *, converted: bool) -> None:
    def formula(text: str) -> str:
        if converted:
            return '<w:r><w:object><o:OLEObject ProgID="Equation.AxMath"/></w:object></w:r>'
        if text == "a":
            return (
                '<m:oMath><m:sSub><m:e><m:acc><m:accPr><m:chr m:val="̂"/></m:accPr>'
                '<m:e><m:r><m:t>a</m:t></m:r></m:e></m:acc></m:e>'
                '<m:sub><m:r><m:t>t</m:t></m:r></m:sub></m:sSub></m:oMath>'
            )
        return f"<m:oMath><m:r><m:t>{text}</m:t></m:r></m:oMath>"
    xml = (
        f'<w:document xmlns:w="{W}" xmlns:m="{M}" xmlns:o="{O}"><w:body>'
        f'<w:p><w:r><w:t>对任意的正整数</w:t></w:r>{formula("n")}</w:p>'
        f'<w:p><w:r><w:t>另一个短公式</w:t></w:r>{formula("a")}</w:p>'
        "</w:body></w:document>"
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", xml.encode("utf-8"))


def write_inspection(path: Path, docx: Path, exports: list[str], *, all_objects: bool) -> Path:
    digest = sha256_file(docx)
    path.write_text(json.dumps({
        "schema": "axmath-internal-tex-diagnostic/v1",
        "target_scope": "all" if all_objects else "selected",
        "input": str(docx.resolve()),
        "input_sha256_before": digest,
        "input_sha256_after": digest,
        "input_unchanged": True,
        "diagnostic_only": True,
        "success": True,
        "integrity_pass": all(value.strip() for value in exports),
        "axmath_count": 2,
        "failed": [],
        "exports": [
            {"ordinal": i, "success": True, "axmath_exported_tex": value}
            for i, value in enumerate(exports, 1)
        ],
    }), encoding="utf-8")
    return path


def main() -> int:
    with tempfile.TemporaryDirectory() as folder:
        root = Path(folder)
        source, final = root / "frozen.docx", root / "converted.docx"
        make_docx(source, converted=False)
        make_docx(final, converted=True)
        counts = content_compare(source, final)["formula_count_state"]
        assert counts["state"] == "exact", counts

        empty = write_inspection(root / "empty-internal.json", final, ["\r", r"$\hat a_t$"], all_objects=True)
        invalid = validate_full_internal_audit(empty, final, 2)
        assert not invalid["passed"] and invalid["defects"] == [
            {"ordinal": 1, "reason": "empty_internal_export"}
        ], invalid
        with mock.patch.object(strict_final_compare, "snapshot", side_effect=AssertionError("render must not start")):
            try:
                strict_final_compare.compare(
                    source, final, root / "compare", sha256_file(source), internal_audit=empty
                )
            except ValueError as exc:
                assert "internal AxMath audit failed" in str(exc)
            else:
                raise AssertionError("equal AxMath counts approved an empty internal formula")

        complete = write_inspection(root / "complete.json", final, ["$n$", r"$\hat a_t$"], all_objects=True)
        assert validate_full_internal_audit(complete, final, 2)["passed"]
        targeted = write_inspection(root / "targeted.json", final, ["\r", r"$\hat a_t$"], all_objects=False)
        review_path = root / "class-e-review.json"
        prepare_class_e_review(source, final, targeted, [1, 2], review_path)
        review = json.loads(review_path.read_text(encoding="utf-8"))
        assert "acc" in review["rows"][1]["source_omml_xml"]
        assert r"\hat" in review["rows"][1]["source_tex"]
        review["rows"][0].update(decision="verified_match", evidence="count looks correct")
        review["rows"][1].update(decision="verified_match", evidence="frozen hat a sub t matches AxMath export")
        review_path.write_text(json.dumps(review, ensure_ascii=False), encoding="utf-8")
        assert not validate_class_e_review(review_path, source, final)["valid"], "empty export was approved"
        review["rows"][0].update(decision="needs_rebuild", evidence="frozen n; AxMath export is empty")
        review_path.write_text(json.dumps(review, ensure_ascii=False), encoding="utf-8")
        valid_review = validate_class_e_review(review_path, source, final)
        assert valid_review["valid"] and valid_review["verified_ordinals"] == [2], valid_review
        assert valid_review["rebuild_ordinals"] == [1], valid_review
        with mock.patch.object(diagnose_after_conversion, "_fast_static_triage", return_value={
            "semantic_tiny_shell_candidates": [{"ordinal": 2}],
            "m2_visual_wrap_loss_candidates": [],
            "overwide_source_singleline_review": [],
        }):
            diagnosed = diagnose_after_conversion.diagnose(
                source, final, root / "diagnosis", class_e_review=review_path
            )
        assert diagnosed["status"] == "repair_review_required", diagnosed
        assert diagnosed["queues"]["CLASS_E_SEMANTIC_REBUILD"] == [1], diagnosed

        wrong = write_inspection(root / "wrong-internal.json", final, ["$m$", r"$\hat a_t$"], all_objects=False)
        wrong_review_path = root / "wrong-class-e-review.json"
        prepare_class_e_review(source, final, wrong, [1], wrong_review_path)
        wrong_review = json.loads(wrong_review_path.read_text(encoding="utf-8"))
        wrong_review["rows"][0].update(
            decision="needs_rebuild", evidence="frozen OMML is n; internal AxMath exports m"
        )
        wrong_review_path.write_text(json.dumps(wrong_review, ensure_ascii=False), encoding="utf-8")
        assert validate_class_e_review(wrong_review_path, source, final)["rebuild_ordinals"] == [1]

        image = root / "page.png"
        image.write_bytes(b"reviewed page evidence")
        compare_path = root / "STRICT_FINAL_COMPARE.json"
        compare_path.write_text(json.dumps({
            "schema": "axmath-strict-final-compare/v3",
            "source": str(source), "final": str(final),
            "source_sha256": sha256_file(source), "final_sha256": sha256_file(final),
            "internal_audit": str(complete), "internal_audit_sha256": sha256_file(complete),
            "internal_audit_pass": True, "final_axmath_count": 2,
            "hard_content_pass": True, "visual_evidence_complete": True,
            "side_by_side_inventory": [{"page": 1, "path": str(image), "sha256": sha256_file(image)}],
        }), encoding="utf-8")
        visual_path = root / "VISUAL_REVIEW.json"
        visual = {
            "schema": "axmath-visual-review/v2",
            "source_sha256": sha256_file(source), "final_sha256": sha256_file(final),
            "pages": [{"page": 1, "image_sha256": sha256_file(image), "passed": True}],
            "overall_passed": True,
            "cross_page_content_anchors_passed": None,
            "cross_page_content_anchors_evidence": "",
            "small_inline_closeups_passed": None,
            "small_inline_closeups_evidence": "",
        }
        visual_path.write_text(json.dumps(visual), encoding="utf-8")
        assert not finalize_visual_review.finalize(compare_path, visual_path)["acceptance_pass"]
        visual.update(
            cross_page_content_anchors_passed=True,
            cross_page_content_anchors_evidence="question 20(2), source p1 to final p1",
            small_inline_closeups_passed=True,
            small_inline_closeups_evidence="ordinal 1 n legible at close zoom",
        )
        visual_path.write_text(json.dumps(visual), encoding="utf-8")
        assert finalize_visual_review.finalize(compare_path, visual_path)["acceptance_pass"]

        final.write_bytes(final.read_bytes() + b"stale hash")
        assert not validate_full_internal_audit(complete, final, 2)["passed"]
        assert not validate_class_e_review(review_path, source, final)["valid"]
    print("INTERNAL_CONTENT_GATE_REGRESSION_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
