from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def main() -> int:
    skill = (ROOT / "SKILL.md").read_text(encoding="utf-8")
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    workflow = (ROOT / "docs" / "AXMATH_CONVERSION_WORKFLOW.md").read_text(encoding="utf-8")
    numbering = (ROOT / "docs" / "NUMBERING_AND_OUTLINE.md").read_text(encoding="utf-8")
    contract = (ROOT / "scripts" / "math_layout_contract.py").read_text(encoding="utf-8")
    compare = (ROOT / "scripts" / "strict_final_compare.py").read_text(encoding="utf-8")
    cli = (ROOT / "axmath.py").read_text(encoding="utf-8")
    punctuation = (ROOT / "scripts" / "normalize_body_punctuation.py").read_text(encoding="utf-8")
    powershell = (ROOT / "scripts" / "axmath_prime_contract.ps1").read_text(encoding="utf-8")

    assert "首轮生成硬规则" in skill
    assert workflow.index("首轮生成硬规则") < workflow.index("## 1. Mandatory entry protocol")
    for text in (skill, readme, workflow):
        for token in ("SOURCE_FINAL_FORMULA_MAP", "source_formula_visual", "added_body_math", "AMSMML2AM", "AMSTeX2AM"):
            assert token in text, (token, "missing from first-pass docs")
        assert "\\underset" in text
        assert "\\smallint" in text

    for token in ("w:numPr", "w:numId", "w:ilvl", "LISTNUM", "SEQ", "AUTONUM", "w:outlineLvl", "悬挂缩进", "制表位"):
        assert token in numbering, token

    for token in (
        "FORMULA_MAP_SCHEMA",
        "source_formula_visual",
        "expected_tex",
        "source_formula_semantic_token_mismatch",
        "added_target_semantic_token_mismatch",
        "smallint_forbidden",
        "mapping_pass",
    ):
        assert token in contract, token

    assert "--formula-map" in compare
    assert "--formula-map" in cli
    assert "word/document.xml" in punctuation
    assert "fldChar" in punctuation
    assert "EXCLUDED_ANCESTORS" in punctuation
    assert "replaced_count" in punctuation
    assert "smallint_forbidden" in powershell
    assert "lim_subscript_requires_underset" in powershell

    active_docs = [skill, readme, workflow, numbering]
    for text in active_docs:
        assert "\u200b" not in text and "\u200c" not in text and "\ufeff" not in text

    print("FIRST_PASS_CONTRACT_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
