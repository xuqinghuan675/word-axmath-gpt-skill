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

    for needle in [
        "Visual layout is the final acceptance authority.",
        "acceptance_pass=true",
        "hard_content_pass=true",
        "VISUAL_REVIEW_TEMPLATE.json",
    ]:
        if needle not in skill:
            errors.append(f"SKILL.md missing contract: {needle}")

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

    if errors:
        print("SELF_CHECK_FAILED")
        for error in errors:
            print(f"- {error}")
        return 1

    print("SELF_CHECK_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
