from __future__ import annotations

import ast
from pathlib import Path

from normalize_numbering_punctuation import normalize_numbering_text
from source_math_structure import classify_count_state
from source_tex_sanity import word_latex_is_sane

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"


def require(text: str, needles: list[str], where: str, errors: list[str]) -> None:
    for needle in needles:
        if needle not in text:
            errors.append(f"{where} missing contract: {needle}")


def main() -> int:
    errors: list[str] = []

    # Every Python file must remain parseable in CI without importing optional
    # Windows/Word dependencies.
    for path in sorted(SCRIPTS.glob("*.py")):
        try:
            ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        except Exception as exc:
            errors.append(f"Python syntax error in {path.name}: {exc}")

    required_files = [
        "source_math_structure.py",
        "source_tex_sanity.py",
        "diagnose_after_conversion.py",
        "diagnose_local_layout.py",
        "build_local_layout_plan.py",
        "build_source_tex_map.py",
        "export_source_visual_lines.ps1",
        "repair_visual_line_split.ps1",
        "validate_local_layout_repair.py",
        "validate_local_visual_line_repair.py",
        "axmath_prime_contract.ps1",
        "normalize_axmath_tex.py",
        "normalize_numbering_punctuation.py",
        "probe_axmath_prime_contract.ps1",
        "test_prime_normalization.py",
        "test_numbering_punctuation.py",
        "export_source_visual_lines.ps1",
        "repair_multisibling_groups.ps1",
        "inspect_axmath_tex.ps1",
        "repair_axmath_from_approved_tex.ps1",
        "strict_final_compare.py",
        "finalize_visual_review.py",
    ]
    for name in required_files:
        if not (SCRIPTS / name).is_file():
            errors.append(f"required production tool missing: {name}")

    skill = (ROOT / "SKILL.md").read_text(encoding="utf-8")
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    requirements = (ROOT / "requirements.txt").read_text(encoding="utf-8")
    one_click = (SCRIPTS / "one_click_convert.py").read_text(encoding="utf-8")
    runner = (SCRIPTS / "run_skill.py").read_text(encoding="utf-8")
    word_runtime = (SCRIPTS / "word_runtime.py").read_text(encoding="utf-8")
    converter = (SCRIPTS / "convert_officemath_to_axmath.ps1").read_text(encoding="utf-8")
    audit = (SCRIPTS / "audit_docx.py").read_text(encoding="utf-8")
    structure = (SCRIPTS / "source_math_structure.py").read_text(encoding="utf-8")
    diagnosis = (SCRIPTS / "diagnose_after_conversion.py").read_text(encoding="utf-8")
    snapshot = (SCRIPTS / "snapshot_docx.py").read_text(encoding="utf-8")
    geo = (SCRIPTS / "formula_geometry_audit.py").read_text(encoding="utf-8")
    map_builder = (SCRIPTS / "build_source_tex_map.py").read_text(encoding="utf-8")
    tex_sanity = (SCRIPTS / "source_tex_sanity.py").read_text(encoding="utf-8")
    prime_normalizer = (SCRIPTS / "normalize_axmath_tex.py").read_text(encoding="utf-8")
    numbering_normalizer = (SCRIPTS / "normalize_numbering_punctuation.py").read_text(encoding="utf-8")
    prime_contract_ps = (SCRIPTS / "axmath_prime_contract.ps1").read_text(encoding="utf-8")
    prime_probe = (SCRIPTS / "probe_axmath_prime_contract.ps1").read_text(encoding="utf-8")
    prime_test = SCRIPTS / "test_prime_normalization.py"
    local_diag = (SCRIPTS / "diagnose_local_layout.py").read_text(encoding="utf-8")
    local_plan = (SCRIPTS / "build_local_layout_plan.py").read_text(encoding="utf-8")
    split_repair = (SCRIPTS / "repair_visual_line_split.ps1").read_text(encoding="utf-8")
    split_validate = (SCRIPTS / "validate_local_layout_repair.py").read_text(encoding="utf-8")
    split_wrapper = (SCRIPTS / "validate_local_visual_line_repair.py").read_text(encoding="utf-8")
    visual_lines = (SCRIPTS / "export_source_visual_lines.ps1").read_text(encoding="utf-8")
    m1_repair = (SCRIPTS / "repair_multisibling_groups.ps1").read_text(encoding="utf-8")
    internal_diag = (SCRIPTS / "inspect_axmath_tex.ps1").read_text(encoding="utf-8")
    approved_tex = (SCRIPTS / "repair_axmath_from_approved_tex.ps1").read_text(encoding="utf-8")
    inline_roundtrip = (SCRIPTS / "repair_axmath_inline_roundtrip.ps1").read_text(encoding="utf-8")
    rebuild = (SCRIPTS / "rebuild_axmath_baselines.ps1").read_text(encoding="utf-8")
    calibrate = (SCRIPTS / "calibrate_axmath_boxes.py").read_text(encoding="utf-8")
    source_latex = (SCRIPTS / "export_source_word_latex.ps1").read_text(encoding="utf-8")
    strict = (SCRIPTS / "strict_final_compare.py").read_text(encoding="utf-8")
    finalizer = (SCRIPTS / "finalize_visual_review.py").read_text(encoding="utf-8")

    # New-chat hard entry and deterministic state machine.
    require(
        skill,
        [
            "A new chat must read this file",
            "Mandatory entry protocol",
            "diagnose_after_conversion.py",
            "NEXT_ACTION.json",
            "M1_MULTISIBLING_OMATHPARA",
            "M2_SOURCE_VISUAL_WRAP_LOSS" if "M2_SOURCE_VISUAL_WRAP_LOSS" in skill else "M2 — source visual wrap",
            "M2-B",
            "inspect_axmath_tex.ps1",
            "source-reading from AxMath mutation",
            "acceptance_pass=true",
        ],
        "SKILL.md",
        errors,
    )
    require(
        readme,
        [
            "diagnose_after_conversion.py",
            "known_multisibling_collapse",
            "超长公式",
            "M2-B：按原稿视觉数学行恢复多个 AxMath",
            "inspect_axmath_tex.ps1",
            "99.02%",
        ],
        "README.md",
        errors,
    )

    # Environment is preflighted before the expensive conversion.
    require(
        one_click,
        [
            "_utf8_console",
            "environment_preflight",
            "missing_python_modules",
            "omml2latex",
            "ready_for_end_to_end",
            "diagnose_after_conversion.py",
            "known_multisibling_collapse",
            "source_unchanged",
            "normalize_docx",
            "numbering_punctuation_normalization",
            "nonmath_text_contract_exact",
        ],
        "one_click_convert.py",
        errors,
    )
    if "omml2latex==0.1.1" not in requirements:
        errors.append("requirements.txt missing pinned omml2latex fallback dependency")

    # M1: raw OfficeMath count gap must be structurally explained, never guessed.
    require(
        structure,
        [
            "multi_sibling_groups",
            "multi_sibling_extra_nodes",
            "batch_collapse_signature_axmath_count",
            "known_multisibling_collapse",
            "STOP_UNEXPLAINED_COUNT_GAP",
        ],
        "source_math_structure.py",
        errors,
    )
    require(
        audit,
        [
            "formula_count_state",
            "known_multisibling_merge_signature",
            "unexpected_formula_count_gap",
            "nonmath_text_numbering_normalized_exact",
            "nonmath_text_contract_exact",
        ],
        "audit_docx.py",
        errors,
    )
    require(
        numbering_normalizer,
        [
            "numbering_dunhao_positions",
            "normalize_numbering_text",
            "normalize_docx",
            "Refusing in-place numbering punctuation normalization.",
            "paragraph-or-line-start Arabic-number label",
        ],
        "normalize_numbering_punctuation.py",
        errors,
    )
    for source, expected in [
        ("1、定义", "1.定义"),
        ("（1）、定义", "（1）.定义"),
        ("(2)、定义", "(2).定义"),
        ("第1、2项", "第1、2项"),
        ("甲、乙", "甲、乙"),
    ]:
        actual = normalize_numbering_text(source)
        if actual != expected:
            errors.append(
                f"numbering punctuation normalization regression: {source!r} -> {actual!r}"
            )
    require(
        m1_repair,
        [
            "M1_MULTISIBLING_OMATHPARA",
            "working_sha256",
            "expected_final_axmath",
            "expected_increase",
            "word_pid_owned",
            "Refusing to overwrite the input DOCX.",
        ],
        "repair_multisibling_groups.ps1",
        errors,
    )

    # M2: evidence is source multiline + measured right-boundary overflow.
    require(
        snapshot,
        [
            "_paragraph_text_bounds",
            "table_cell_requires_visual_review",
            "multi_column_section_requires_visual_review",
            "right_overflow_pt",
            "overflows_text_right",
            'profile: str = "full"',
            'choices=("full", "geometry")',
        ],
        "snapshot_docx.py",
        errors,
    )
    require(
        geo,
        [
            "visual_wrap_loss_candidates",
            "M2_SOURCE_VISUAL_WRAP_LOSS",
            "source_multiline_working_ole_crosses_measured_text_boundary",
            "overflows_text_right",
            '"auto_apply": False',
        ],
        "formula_geometry_audit.py",
        errors,
    )
    require(
        visual_lines,
        [
            "Get-VisualLineRanges",
            "Same-VisualLine",
            "Non-monotonic OfficeMath layout",
            "one-task-owned-word-process-per-ordinal",
            "source_unchanged",
            "MaxRetries",
        ],
        "export_source_visual_lines.ps1",
        errors,
    )
    if "AxMath.dotm" in visual_lines or "AMSTeX2AM" in visual_lines:
        errors.append(
            "export_source_visual_lines.ps1 must remain a pure source-reading phase without AxMath"
        )
    require(
        map_builder,
        [
            "word_latex_is_sane",
            "fragment_exact_omml_fallback",
            "omml2latex",
            r"\begin{aligned}",
            "M2_SOURCE_VISUAL_WRAP_LOSS",
            "working_sha256",
            "working_paragraph_index",
            "M2 map is forbidden until formula identity/count is exact",
        ],
        "build_source_tex_map.py",
        errors,
    )
    require(
        tex_sanity,
        [
            "embedded_line_control",
            "unicode_math_alphanumeric",
            "word_linear_artifact",
            "unbalanced_brace",
        ],
        "source_tex_sanity.py",
        errors,
    )
    require(
        approved_tex,
        [
            "Stale approved map",
            "M2_SOURCE_VISUAL_WRAP_LOSS",
            r"\begin\{aligned\}",
            "every M2/Class-E row must remain exactly one AxMath object",
            "axmath_after",
            "omath_after",
            "paragraphs_after",
            "$doc.Save()",
        ],
        "repair_axmath_from_approved_tex.ps1",
        errors,
    )

    # Class D/E: inspect internal OLE on a copy; frozen source remains semantic authority.
    require(
        internal_diag,
        [
            "diagnostic_only",
            "OpenNoRepairDialog",
            "AMSAM2TeX",
            "temporary copies",
            "input_unchanged",
        ],
        "inspect_axmath_tex.ps1",
        errors,
    )
    require(
        geo,
        [
            "semantic_rebuild_candidates",
            "roundtrip_semantic_risk_candidates",
            "source_contains_prime_or_derivative_marker",
        ],
        "formula_geometry_audit.py",
        errors,
    )
    require(
        source_latex,
        [
            "source_sha256_before",
            "source_sha256_after",
            "source_unchanged",
            "Dialogs.Item(2844)",
            "word_pid_owned",
        ],
        "export_source_word_latex.ps1",
        errors,
    )

    # Mandatory dispatcher must stop structural/count drift before geometry.
    require(
        diagnosis,
        [
            "blocked_content_drift",
            "known_multisibling_collapse",
            "M1_MULTISIBLING_REPAIR_MAP.json",
            "blocked_unexplained_formula_count",
            "fast_static_targeted",
            "_fast_static_triage",
            "FAST_SOURCE_MULTILINE_PROBE.json",
            "conservative_overwide_candidates",
            "--deep-geometry",
            "ready_for_strict_final_compare",
            "Do not tune page count directly.",
        ],
        "diagnose_after_conversion.py",
        errors,
    )
    try:
        fast_pos = diagnosis.index("fast = _fast_static_triage")
        deep_pos = diagnosis.index("if not deep_geometry:")
        snapshot_pos = diagnosis.index("source_snapshot = snapshot(")
        if not (fast_pos < deep_pos < snapshot_pos):
            errors.append(
                "diagnose_after_conversion.py no longer keeps full geometry behind the deep-geometry gate"
            )
    except ValueError:
        errors.append("diagnose_after_conversion.py fast/deep dispatcher ordering is incomplete")

    require(
        map_builder,
        [
            "static_axmath_inventory",
            "conservative_overwide_candidates",
            "single_section_single_column",
            "working_paragraph_index",
            "working_paragraph_axmath_index",
        ],
        "build_source_tex_map.py",
        errors,
    )

    require(
        word_runtime,
        [
            "refused_unowned_attach",
            "ambiguous_created_pids",
            "Could not prove a distinct task-owned WINWORD process",
            "if app is not None and self.meta.owned_pid",
            "Refusing to attach to or mutate a pre-existing Word instance",
        ],
        "word_runtime.py",
        errors,
    )
    ownership_check = word_runtime.find("if owned_pid is None:")
    first_mutation = word_runtime.find("self.app.Visible = self.visible")
    if ownership_check < 0 or first_mutation < 0 or ownership_check > first_mutation:
        errors.append("OwnedWord mutates Word before proving distinct PID ownership")

    # PowerShell COM collections are not reliably enumerable with foreach.
    # Production regression: 1521 InlineShapes / 1520 AxMath by Count+Item,
    # but foreach over InlineShapes yielded zero AxMath objects.
    for ps_name in [
        "repair_multisibling_groups.ps1",
        "repair_axmath_from_approved_tex.ps1",
        "repair_axmath_inline_roundtrip.ps1",
        "inspect_axmath_tex.ps1",
        "rebuild_axmath_baselines.ps1",
    ]:
        ps_text = (SCRIPTS / ps_name).read_text(encoding="utf-8")
        for banned in [
            "foreach($s in $doc.InlineShapes)",
            "foreach($s in $d.InlineShapes)",
            "foreach($s in $RangeOrDoc.InlineShapes)",
            "return ,@($arr)",
        ]:
            if banned in ps_text:
                errors.append(f"{ps_name} regressed to unsafe COM/array enumeration: {banned}")
        if "$collection.Item($i)" not in ps_text:
            errors.append(f"{ps_name} no longer uses explicit Count/Item COM enumeration")

    for ps_name in [
        "repair_multisibling_groups.ps1",
        "repair_axmath_from_approved_tex.ps1",
        "inspect_axmath_tex.ps1",
        "export_source_visual_lines.ps1",
        "export_source_word_latex.ps1",
    ]:
        ps_text = (SCRIPTS / ps_name).read_text(encoding="utf-8")
        require(
            ps_text,
            ["Get-SharedSha256", "[IO.FileShare]::ReadWrite", "[IO.FileShare]::Delete"],
            ps_name,
            errors,
        )

    # M2-B local-layout split is a first-class, hash-bound repair route.
    require(
        local_diag,
        [
            "M2_SOURCE_VISUAL_LINE_SPLIT",
            "start_anchor",
            "paragraph_offset",
            "alignment_mismatches",
            "export_source_visual_lines.ps1",
            "repair_visual_line_split.ps1",
        ],
        "diagnose_local_layout.py",
        errors,
    )
    require(
        local_plan,
        [
            "axmath-local-layout-plan/v1",
            "start_anchor",
            "paragraph_offset",
            "working_current_axmath_count",
        ],
        "build_local_layout_plan.py",
        errors,
    )
    require(
        map_builder,
        [
            "build_visual_line_split_map",
            "M2_SOURCE_VISUAL_LINE_SPLIT",
            "expected_axmath_increase",
            "expected_final_axmath_count",
            "working_paragraph_index",
            "axmath_builtin_prime_v2",
        ],
        "build_source_tex_map.py",
        errors,
    )
    require(
        split_repair,
        [
            "M2_SOURCE_VISUAL_LINE_SPLIT",
            "first_repair_paragraph",
            "expected_axmath_increase",
            "expected_final_axmath",
            "[string][char]11",
            "Could not prove distinct task-owned Word PID",
            "Verify-AxMathPrimeDonor",
        ],
        "repair_visual_line_split.ps1",
        errors,
    )
    require(
        split_validate,
        [
            "axmath-local-layout-repair-ledger/v1",
            "prefix_semantic_unchanged",
            "nonmath_text_flat_equal",
            "expected_candidate_axmath",
            "expected_axmath_increase",
        ],
        "validate_local_layout_repair.py",
        errors,
    )
    require(
        split_wrapper,
        ["validate_local_layout_repair", "--baseline", "--candidate", "--map"],
        "validate_local_visual_line_repair.py",
        errors,
    )
    require(
        strict,
        [
            "--repair-ledger",
            "repair_ledger_ok",
            "formula_count_contract_ok",
            "paragraph_contract_ok",
            "nonmath_text_contract_ok",
        ],
        "strict_final_compare.py",
        errors,
    )

    # Prime semantics are a semantic contract, independent of geometry.
    require(
        geo,
        [
            "prime_semantic_candidates",
            "prime_semantic_ordinals",
            "contains_prime_or_derivative_marker",
            "source_contains_prime_or_derivative_marker",
        ],
        "formula_geometry_audit.py",
        errors,
    )
    require(
        prime_normalizer,
        [
            "PRIME_SOURCE_MARKERS",
            "NONCANONICAL_PRIME_LITERALS",
            "normalize_axmath_tex",
            "canonical_prime_issues",
            "_OMML_FUSED_PRIME_EXP_RE",
            "fused_prime_exponent",
            "ungrouped_primed_atom_exponent",
            "axmath_builtin_prime_v2",
        ],
        "normalize_axmath_tex.py",
        errors,
    )
    require(
        prime_contract_ps,
        [
            "Assert-CanonicalPrimeTeX",
            "Get-CanonicalPrimeOrders",
            "Assert-AxMathPrimeRoundTrip",
            "Verify-AxMathPrimeDonor",
            "noncanonical_prime_literal",
            "ungrouped_primed_atom_exponent",
        ],
        "axmath_prime_contract.ps1",
        errors,
    )
    if not prime_test.is_file():
        errors.append("test_prime_normalization.py missing")
    require(
        approved_tex,
        [
            "axmath_prime_contract.ps1",
            "Verify-AxMathPrimeDonor",
            "axmath_builtin_prime_v2",
        ],
        "repair_axmath_from_approved_tex.ps1",
        errors,
    )
    require(
        m1_repair,
        ["axmath_prime_contract.ps1", "Verify-AxMathPrimeDonor"],
        "repair_multisibling_groups.ps1",
        errors,
    )
    if "foreach($s in $d.InlineShapes)" in prime_probe or "return ,@($arr)" in prime_probe:
        errors.append("probe_axmath_prime_contract.ps1 regressed to unsafe COM enumeration")

    # Preserve the existing class guards.
    require(
        inline_roundtrip,
        [
            "targets.Count -gt 12",
            "reviewed_target_limit=12",
            "word_pid_owned",
            "Refusing to overwrite the input DOCX.",
        ],
        "repair_axmath_inline_roundtrip.ps1",
        errors,
    )
    require(
        rebuild,
        [
            "targets.Count -gt 3",
            "probe_only",
            "word_pid_owned",
            "Refusing to overwrite the input DOCX.",
        ],
        "rebuild_axmath_baselines.ps1",
        errors,
    )
    require(
        calibrate,
        ["Refusing in-place OLE calibration", "working_sha256", "--ordinals"],
        "calibrate_axmath_boxes.py",
        errors,
    )

    # Phase A performance/safety must never regress.
    require(
        converter,
        [
            "save_seconds",
            "macro_seconds_total",
            "macro_share_percent",
            "crash_dumps_new",
            "batch_limit_owner",
            "WaitingConvert -Value 0",
            "$doc.Save()",
            "output_readonly_cleared",
        ],
        "convert_officemath_to_axmath.ps1",
        errors,
    )
    if "Refusing background AxMath conversion because Word is already running" in converter:
        errors.append("converter regressed to blocking on a pre-existing user Word process")
    if "Word is already running; refusing unattended conversion." in runner:
        errors.append("run_skill.py regressed to blocking on a pre-existing user Word process")
    require(
        runner,
        [
            "source_unchanged",
            "_performance_summary",
            "owned_word_pid",
            'conversion.get("word_pid_owned")',
        ],
        "run_skill.py",
        errors,
    )

    # Full final acceptance remains hash-bound and visual.
    require(
        strict,
        [
            "hard_content_pass",
            "VISUAL_REVIEW_TEMPLATE.json",
            "expected-source-sha256",
            'snapshot(source_docx, outdir / "source", "source")',
            'snapshot(final_docx, outdir / "final", "final")',
        ],
        "strict_final_compare.py",
        errors,
    )
    if 'return 0 if report["structural_pass"] else 1' in strict:
        errors.append("strict_final_compare.py still treats structural_pass as final acceptance")
    require(
        finalizer,
        ["acceptance_pass", "image_sha256", "source_sha256", "final_sha256"],
        "finalize_visual_review.py",
        errors,
    )

    # Word linear export regression: non-empty output is not sufficient.
    sane, reasons = word_latex_is_sane(r"\frac{1}{2}")
    if not sane or reasons:
        errors.append("source TeX sanity gate rejects ordinary LaTeX")
    for sample, expected_reason in [
        ("\\𝑙𝑒\r0\r1", "unicode_math_alphanumeric"),
        ("x\r1", "embedded_line_control"),
        ("▒x", "word_linear_artifact:▒"),
        ("\\begin(x)", "word_linear_artifact:\\begin("),
        ("{x", "unbalanced_brace"),
    ]:
        ok, sample_reasons = word_latex_is_sane(sample)
        if ok or expected_reason not in sample_reasons:
            errors.append(
                f"source TeX sanity gate missed {expected_reason}: {sample_reasons}"
            )

    # Pure-function regression for the production count signature.
    synthetic = {
        "raw_omath_count": 1574,
        "multi_sibling_extra_nodes": 54,
    }
    state = classify_count_state(synthetic, 1520, 0)
    if state["state"] != "known_multisibling_collapse":
        errors.append("M1 classifier no longer recognizes 1574 -> 1520 with 54 known extra nodes")
    if classify_count_state(synthetic, 1519, 0)["state"] != "unexpected_formula_loss":
        errors.append("M1 classifier incorrectly explains an unmatched count gap")
    if classify_count_state(synthetic, 1574, 0)["state"] != "exact":
        errors.append("count classifier no longer recognizes exact conversion")
    if classify_count_state(synthetic, 1574, 1)["state"] != "residual_officemath":
        errors.append("count classifier no longer blocks residual OfficeMath")

    if errors:
        print("SELF_CHECK_FAILED")
        for error in errors:
            print(f"- {error}")
        return 1

    print("SELF_CHECK_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
