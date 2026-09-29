from __future__ import annotations

import ast
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"


def main() -> int:
    errors: list[str] = []

    for path in sorted(SCRIPTS.glob("*.py")):
        try:
            ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        except Exception as exc:
            errors.append(f"Python syntax error in {path.name}: {exc}")

    skill = (ROOT / "SKILL.md").read_text(encoding="utf-8")
    strict = (SCRIPTS / "strict_final_compare.py").read_text(encoding="utf-8")
    finalizer = SCRIPTS / "finalize_visual_review.py"
    geo = (SCRIPTS / "formula_geometry_audit.py").read_text(encoding="utf-8")
    calibrate = (SCRIPTS / "calibrate_axmath_boxes.py").read_text(encoding="utf-8")
    one_click = (SCRIPTS / "one_click_convert.py").read_text(encoding="utf-8")
    runner = (SCRIPTS / "run_skill.py").read_text(encoding="utf-8")
    converter = (SCRIPTS / "convert_officemath_to_axmath.ps1").read_text(encoding="utf-8")
    rebuild = (SCRIPTS / "rebuild_axmath_baselines.ps1").read_text(encoding="utf-8")
    inline_roundtrip = (SCRIPTS / "repair_axmath_inline_roundtrip.ps1").read_text(encoding="utf-8")
    source_latex_export = (SCRIPTS / "export_source_word_latex.ps1").read_text(encoding="utf-8")
    approved_tex_repair = (SCRIPTS / "repair_axmath_from_approved_tex.ps1").read_text(encoding="utf-8")
    prime_normalizer = (SCRIPTS / "normalize_axmath_tex.py").read_text(encoding="utf-8")
    prime_test = SCRIPTS / "test_prime_normalization.py"
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    snapshot = (SCRIPTS / "snapshot_docx.py").read_text(encoding="utf-8")

    for needle in [
        "Visual layout is the final acceptance authority.",
        "acceptance_pass=true",
        "hard_content_pass=true",
        "VISUAL_REVIEW_TEMPLATE.json",
        "64–66",
        "probe-only",
        "source-semantic",
    ]:
        if needle not in skill:
            errors.append(f"SKILL.md missing contract: {needle}")

    for needle in ["64~66", "99.02%", "WaitingConvert"]:
        if needle not in readme:
            errors.append(f"README.md missing Phase A performance contract: {needle}")

    if 'return 0 if report["structural_pass"] else 1' in strict:
        errors.append("strict_final_compare.py still treats structural_pass as final acceptance")
    for needle in ["hard_content_pass", "VISUAL_REVIEW_TEMPLATE.json", "expected-source-sha256"]:
        if needle not in strict:
            errors.append(f"strict_final_compare.py missing: {needle}")
    if not finalizer.is_file():
        errors.append("finalize_visual_review.py missing")

    if '"auto_apply": False' not in geo or "requires_visual_confirmation" not in geo:
        errors.append("formula_geometry_audit.py may still auto-apply geometry recommendations")
    if 'report["unresolved_same_line_breaks"] = list(report["same_line_breaks"])' not in geo:
        errors.append("geometry audit still risks treating a planned repair as already resolved")
    for needle in [
        "semantic_rebuild_candidates",
        "prime_semantic_candidates",
        "prime_semantic_ordinals",
        "roundtrip_semantic_risk_candidates",
        "nontrivial_source_in_tiny_axmath_shell",
        "source_contains_prime_or_derivative_marker",
        "contains_prime_or_derivative_marker",
    ]:
        if needle not in geo:
            errors.append(f"formula_geometry_audit.py missing repair triage contract: {needle}")

    for needle in ["Refusing in-place OLE calibration", "working_sha256", "--ordinals"]:
        if needle not in calibrate:
            errors.append(f"calibrate_axmath_boxes.py missing guard: {needle}")

    for name in [
        "convert_officemath_to_axmath.ps1",
        "rebuild_axmath_baselines.ps1",
        "repair_axmath_inline_roundtrip.ps1",
    ]:
        text = (SCRIPTS / name).read_text(encoding="utf-8")
        for needle in ["Refusing to overwrite the input DOCX.", "OverwriteOutput"]:
            if needle not in text:
                errors.append(f"{name} missing safety guard: {needle}")

    for name, text in [("one_click_convert.py", one_click), ("run_skill.py", runner)]:
        if "source_unchanged" not in text:
            errors.append(f"{name} does not verify source integrity")
    if "Refusing to overwrite the input DOCX." not in runner:
        errors.append("run_skill.py does not reject input=output")
    for stale in ["blocked_word_running", "blocked_locked"]:
        if stale in one_click:
            errors.append(f"one_click_convert.py regressed to legacy blocker: {stale}")
    if "Word is already running; refusing unattended conversion." in runner:
        errors.append("run_skill.py regressed to blocking on user Word")
    if "Refusing background AxMath conversion because Word is already running" in converter:
        errors.append("convert_officemath_to_axmath.ps1 regressed to blocking on user Word")
    if "output_readonly_cleared" not in converter:
        errors.append("convert_officemath_to_axmath.ps1 no longer clears read-only on the working copy")
    for name, text in [("snapshot_docx.py", snapshot), ("formula_geometry_audit.py", geo)]:
        if "require_clean=False" not in text:
            errors.append(f"{name} regressed to requiring all user Word processes to be closed")
    for needle in ['profile: str = "full"', 'choices=("full", "geometry")', 'if profile == "full"']:
        if needle not in snapshot:
            errors.append(f"snapshot_docx.py missing geometry/full profile contract: {needle}")
    for needle in [
        'snapshot(source_docx, outdir / "source", "source")',
        'snapshot(final_docx, outdir / "final", "final")',
    ]:
        if needle not in strict:
            errors.append(f"strict_final_compare.py no longer uses default full snapshots: {needle}")
    if "--profile geometry" not in skill:
        errors.append("SKILL.md missing intermediate geometry snapshot fast path")

    for needle in [
        "save_seconds",
        "macro_seconds_total",
        "macro_share_percent",
        "crash_dumps_new",
        "batch_limit_owner",
        "WaitingConvert -Value 0",
        "$doc.Save()",
    ]:
        if needle not in converter:
            errors.append(f"convert_officemath_to_axmath.ps1 missing performance/safety contract: {needle}")

    for needle in ["targets.Count -gt 3", "probe_only", "preexisting_word_pids", "word_pid_owned"]:
        if needle not in rebuild:
            errors.append(f"rebuild_axmath_baselines.ps1 missing probe/ownership guard: {needle}")
    if "Refusing visible AxMath rebuild because Word is already running" in rebuild:
        errors.append("rebuild_axmath_baselines.ps1 regressed to blocking on user Word")

    for needle in ["targets.Count -gt 12", "reviewed_target_limit=12", "preexisting_word_pids", "word_pid_owned"]:
        if needle not in inline_roundtrip:
            errors.append(f"repair_axmath_inline_roundtrip.ps1 missing reviewed-target/ownership guard: {needle}")
    if "Refusing inline roundtrip because Word is already running" in inline_roundtrip:
        errors.append("repair_axmath_inline_roundtrip.ps1 regressed to blocking on user Word")
    for needle in ["permission to roundtrip every member", "prime/derivative", "semantic_rebuild_candidates"]:
        if needle not in skill:
            errors.append(f"SKILL.md missing targeted Class A/Class E routing contract: {needle}")

    for needle in [
        "source_sha256_before",
        "source_sha256_after",
        "source_unchanged",
        "word_pid_owned",
        "Dialogs.Item(2844)",
        "OMaths.Item($ord).Range.FormattedText",
        "OutputJson must not point to the source DOCX.",
    ]:
        if needle not in source_latex_export:
            errors.append(f"export_source_word_latex.ps1 missing frozen-source export guard: {needle}")
    for needle in [
        "Refusing to overwrite the input DOCX.",
        "AMSTeX2AM",
        "approved_tex_map",
        "word_pid_owned",
        "axmath_after",
        "omath_after",
        "paragraphs_after",
        "$doc.Save()",
        "StartsWith('$$')",
        "approved TeX map",
        "Assert-CanonicalPrimeTeX",
        "noncanonical_prime_literal",
        "ungrouped_primed_atom_exponent",
        "Assert-AxMathPrimeRoundTrip",
        "axmath_builtin_prime_v2",
    ]:
        if needle not in approved_tex_repair:
            errors.append(f"repair_axmath_from_approved_tex.ps1 missing approved-map safety contract: {needle}")
    for needle in [
        "export_source_word_latex.ps1",
        "normalize_axmath_tex.py",
        "repair_axmath_from_approved_tex.ps1",
        "GPT-approved map",
        "prime_semantic_candidates",
        "first prime",
    ]:
        if needle not in skill:
            errors.append(f"SKILL.md missing production Class E contract: {needle}")

    for needle in [
        "PRIME_SOURCE_MARKERS",
        "NONCANONICAL_PRIME_LITERALS",
        "normalize_axmath_tex",
        "canonical_prime_issues",
        "ungrouped_primed_atom_exponent",
        "axmath_builtin_prime_v2",
    ]:
        if needle not in prime_normalizer:
            errors.append(f"normalize_axmath_tex.py missing prime contract: {needle}")
    if not prime_test.is_file():
        errors.append("test_prime_normalization.py missing")

    for needle in ["_performance_summary", '"performance": _performance_summary(conversion)']:
        if needle not in runner:
            errors.append(f"run_skill.py missing performance summary contract: {needle}")
    for needle in ["owned_word_pid", 'conversion.get("word_pid_owned")', "psutil.pid_exists(owned_word_pid)"]:
        if needle not in runner:
            errors.append(f"run_skill.py missing owned-Word leak guard: {needle}")
    if '"performance": performance' not in one_click:
        errors.append("one_click_convert.py does not propagate performance summary")

    if errors:
        print("SELF_CHECK_FAILED")
        for error in errors:
            print(f"- {error}")
        return 1

    print("SELF_CHECK_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
