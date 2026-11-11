"""Fast regression for the slim happy path. No Word installation required."""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from unittest import mock
from xml.sax.saxutils import escape

import one_click_convert
from audit_docx import compare


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
O = "urn:schemas-microsoft-com:office:office"


def make_docx(path: Path, *, label: str, formula: str = "x+y", converted: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    equation = (
        '<w:r><w:object><o:OLEObject ProgID="Equation.AxMath"/></w:object></w:r>'
        if converted
        else f'<m:oMath><m:r><m:t>{escape(formula)}</m:t></m:r></m:oMath>'
    )
    xml = (
        f'<w:document xmlns:w="{W}" xmlns:m="{M}" xmlns:o="{O}"><w:body>'
        f'<w:p><w:r><w:t>{escape(label)}</w:t></w:r>{equation}</w:p>'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:left="1440" w:right="1440"/></w:sectPr>'
        '</w:body></w:document>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", xml.encode("utf-8"))


def simulate_conversion(label: str, formula: str = "x+y", audit_scope: str = "preserve"):
    def run(cmd, **_kwargs):
        assert "--quiet" in cmd
        input_path = Path(cmd[cmd.index("--input") + 1])
        output_path = Path(cmd[cmd.index("--output") + 1])
        make_docx(output_path, label=label, formula=formula, converted=True)
        analysis = compare(input_path, output_path, numbering_scope=audit_scope)
        assert analysis["formula_count_state"]["state"] == "exact"
        report = {
            "returncode": 0,
            "conversion_report_fresh": True,
            "source_unchanged": True,
            "audit": analysis,
            "performance": {"batch_count": 1},
        }
        Path(str(output_path) + ".skill-report.json").write_text(
            json.dumps(report, ensure_ascii=False), encoding="utf-8"
        )
        return subprocess.CompletedProcess(cmd, 0, "", "")
    return run


def diagnose_cli(source: Path, working: Path, outdir: Path) -> dict:
    process = subprocess.run(
        [sys.executable, str(Path(__file__).parent / "diagnose_after_conversion.py"),
         "--source", str(source), "--working", str(working),
         "--outdir", str(outdir)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=20,
    )
    assert process.returncode == 0, process.stderr or process.stdout
    return json.loads((outdir / "NEXT_ACTION.json").read_text(encoding="utf-8"))


def test_happy_path(folder: Path):
    src = folder / "unlabeled.docx"
    make_docx(src, label="普通段落")
    run_dir = folder / "unlabeled-run"
    with (
        mock.patch.object(one_click_convert.subprocess, "run", side_effect=simulate_conversion("普通段落")),
        mock.patch.object(one_click_convert, "analyze", side_effect=AssertionError("redundant scan")),
    ):
        report = one_click_convert.run_conversion(src, run_dir)
    assert report["status"] == "converted_ready_for_gpt_review"
    assert report["working_docx"] == report["conversion_working_docx"]
    assert report["numbering_punctuation_normalization"]["skipped_by_preserve_policy"]
    assert not (run_dir / "logs").exists()
    assert len(list((run_dir / "working").glob("*.docx"))) == 1
    assert report["conversion_log"] is None
    assert report["working_axmath"] == report["source_omath"] == 1
    frozen = Path(report["frozen_source"])
    working = Path(report["working_docx"])
    review = run_dir / "review"
    diagnosed = diagnose_cli(frozen, working, review)
    assert diagnosed["status"] == "ready_for_strict_final_compare"
    assert diagnosed["queues"]["CLASS_E_PRIME_RISK"] == []
    assert diagnosed["fast_triage"] is None
    assert diagnosed["semantic_risks"] is None
    assert [p.name for p in review.iterdir() if p.suffix == ".json"] == ["NEXT_ACTION.json"]


def test_numbered_case(folder: Path, label: str, scope: str, expected_text: str):
    src = folder / f"numbered-{scope}.docx"
    make_docx(src, label=label)
    with mock.patch.object(
        one_click_convert.subprocess, "run", side_effect=simulate_conversion(label)
    ):
        report = one_click_convert.run_conversion(
            src, folder / f"numbered-run-{scope}", numbering_scope=scope
        )
    assert report["status"] == "converted_ready_for_gpt_review"
    assert report["numbering_punctuation_normalization"]["replacements"] == 1
    assert report["working_docx"] != report["conversion_working_docx"]
    with zipfile.ZipFile(report["working_docx"]) as z:
        assert expected_text in z.read("word/document.xml").decode("utf-8")


def test_preserve_numbering_contract(folder: Path):
    label = "（1）、保留原小问编号"
    source = folder / "preserve.docx"
    make_docx(source, label=label)
    with mock.patch.object(one_click_convert.subprocess, "run", side_effect=simulate_conversion(label)):
        report = one_click_convert.run_conversion(source, folder / "preserve-run")
    assert report["status"] == "converted_ready_for_gpt_review"
    assert report["numbering_scope"] == "preserve"
    assert report["working_docx"] == report["conversion_working_docx"]
    assert report["numbering_punctuation_normalization"]["skipped_by_preserve_policy"]
    assert compare(Path(report["frozen_source"]), Path(report["working_docx"]), numbering_scope="preserve")["nonmath_text_contract_exact"]

    # Simulate a child reporting success under its default, more permissive
    # audit while its output has changed a subquestion label. Preserve must
    # perform its own comparison and reject that result.
    with mock.patch.object(one_click_convert.subprocess, "run", side_effect=simulate_conversion("（1）.保留原小问编号", audit_scope="line-start")):
        drift = one_click_convert.run_conversion(source, folder / "preserve-drift-run", numbering_scope="preserve")
    assert drift["status"] == "conversion_failed"
    assert not drift["nonmath_text_contract_exact"]


def test_resume_rejects_changed_numbering_policy(folder: Path):
    source = folder / "resume-source.docx"
    make_docx(source, label="（1）、原题")
    run_dir = folder / "resume-policy-run"
    run_dir.mkdir()
    state = {
        "numbering_scope": "preserve",
        "documents": [{"source": str(source.resolve()), "sha256": one_click_convert.sha256_file(source)}],
    }
    (run_dir / "SOURCE_STATE.json").write_text(json.dumps(state), encoding="utf-8")
    args = ["one_click_convert.py", "--input", str(source), "--resume-run", str(run_dir),
            "--numbering-scope", "line-start"]
    with (
        mock.patch.object(sys, "argv", args),
        mock.patch.object(one_click_convert, "discover_docs", return_value=[source]),
        mock.patch.object(one_click_convert, "inspect_source", return_value={"source": str(source), "state": "ready_officemath"}),
        mock.patch.object(one_click_convert, "environment_preflight", return_value={"ready_for_end_to_end": True}),
        mock.patch.object(one_click_convert, "word_pids", return_value=[]),
        mock.patch.object(one_click_convert, "run_conversion", side_effect=AssertionError("must not resume")),
    ):
        try:
            one_click_convert.main()
        except RuntimeError as error:
            assert "numbering policy differs" in str(error)
        else:
            raise AssertionError("Changed policy was accepted on resume")


def test_failure_preserves_diagnostics(folder: Path):
    src = folder / "broken.docx"
    make_docx(src, label="普通文字")
    result = subprocess.CompletedProcess([], 1, stdout="", stderr="converted failed")
    with mock.patch.object(one_click_convert.subprocess, "run", return_value=result):
        report = one_click_convert.run_conversion(src, folder / "broken-run")
    assert report["status"] == "conversion_failed"
    assert Path(report["conversion_log"]).is_file()


def test_semantic_risks(folder: Path):
    for formula, warning, expected_status in (
        ("f′(x)", "CLASS_E_PRIME_RISK", "repair_review_required"),
        ("A∪B", "SET_SYMBOL_VISUAL_REVIEW", "ready_for_strict_final_compare"),
    ):
        src, cand = folder / (formula + "-src.docx"), folder / (formula + "-cand.docx")
        make_docx(src, label="正文", formula=formula)
        make_docx(cand, label="正文", formula=formula, converted=True)
        review_dir = folder / (formula + "-review")
        report = diagnose_cli(src, cand, review_dir)
        assert report["status"] == expected_status
        assert report["queues"][warning] == [1]
        assert (review_dir / "SOURCE_SEMANTIC_RISKS.json").exists()
        assert (review_dir / "NEXT_ACTION.json").exists()
        assert not (review_dir / "FAST_SOURCE_MULTILINE_PROBE.json").exists()
        if warning == "SET_SYMBOL_VISUAL_REVIEW":
            assert any("collection symbol" in x for x in report["next_actions"])
        else:
            assert (review_dir / "PRIME_REVIEW_TEMPLATE.json").is_file()


def main():
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        test_happy_path(root)
        test_numbered_case(root, "1、编号", "line-start", "1.编号")
        test_numbered_case(root, "第1、2项", "anywhere", "第1.2项")
        test_preserve_numbering_contract(root)
        test_resume_rejects_changed_numbering_policy(root)
        test_failure_preserves_diagnostics(root)
        test_semantic_risks(root)
    print("SLIM_PIPELINE_REGRESSION_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
