from __future__ import annotations

import hashlib
import json
import tempfile
import zipfile
from pathlib import Path

import build_source_tex_map as source_tex_map
from math_layout_contract import (
    build_formula_map_template,
    static_math_inventory,
    validate_math_layout_contract,
)
from normalize_axmath_tex import (
    limit_operator_issues,
    limit_operator_placements,
    normalize_axmath_tex,
    smallint_issues,
)
from source_semantic_risks import scan_source_math

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
O = "urn:schemas-microsoft-com:office:office"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _write_docx(path: Path, document_xml: str) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", document_xml.encode("utf-8"))


def _source_docx(path: Path) -> None:
    xml = (
        f'<w:document xmlns:w="{W}" xmlns:m="{M}"><w:body>'
        '<w:p><m:oMath><m:limLow>'
        '<m:e><m:r><m:t>lim</m:t></m:r></m:e>'
        '<m:lim><m:r><m:t>x→0</m:t></m:r></m:lim>'
        '</m:limLow></m:oMath></w:p>'
        '<w:p><m:oMath><m:r><m:t>f</m:t></m:r>'
        '<m:r><m:chr m:val="′"/></m:r>'
        '<m:r><m:t>(x)</m:t></m:r></m:oMath></w:p>'
        '<w:p><m:oMath><m:sSup>'
        '<m:e><m:r><m:t>lim</m:t></m:r></m:e>'
        '<m:sup><m:r><m:t>n→∞</m:t></m:r></m:sup>'
        '</m:sSup></m:oMath></w:p>'
        '<w:p><m:oMath><m:sSubSup>'
        '<m:e><m:r><m:t>lim</m:t></m:r></m:e>'
        '<m:sub><m:r><m:t>x→0</m:t></m:r></m:sub>'
        '<m:sup><m:r><m:t>+</m:t></m:r></m:sup>'
        '</m:sSubSup></m:oMath></w:p>'
        '<w:p><m:oMath>'
        '<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>'
        '<m:r><m:t>+</m:t></m:r>'
        '<m:rad><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>'
        '<m:r><m:t>+</m:t></m:r>'
        '<m:m><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e></m:mr></m:m>'
        '<m:r><m:t>+</m:t></m:r>'
        '<m:eqArr><m:e><m:r><m:t>x≤1</m:t></m:r></m:e>'
        '<m:e><m:r><m:t>0</m:t></m:r><m:r><m:t>x>1</m:t></m:r></m:e></m:eqArr>'
        '<m:r><m:t>+</m:t></m:r>'
        '<m:d><m:dPr><m:begChr m:val="|"/><m:endChr m:val="|"/></m:dPr>'
        '<m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>'
        '</m:oMath></w:p>'
        '</w:body></w:document>'
    )
    _write_docx(path, xml)


def _final_docx(path: Path, count: int = 5) -> None:
    objects = "".join(
        f'<w:p><w:r><w:object><o:OLEObject ProgID="Equation.AxMath" r:id="rId{i}"/>'
        '</w:object></w:r></w:p>'
        for i in range(1, count + 1)
    )
    xml = (
        f'<w:document xmlns:w="{W}" xmlns:o="{O}" xmlns:r="{R}">'
        f'<w:body>{objects}</w:body></w:document>'
    )
    _write_docx(path, xml)


def _internal_audit(final: Path, *texes: str) -> Path:
    digest = hashlib.sha256(final.read_bytes()).hexdigest()
    path = final.with_suffix(".internal-audit.json")
    payload = {
        "schema": "axmath-internal-tex-diagnostic/v1",
        "input": str(final.resolve()),
        "input_sha256_before": digest,
        "input_sha256_after": digest,
        "input_unchanged": True,
        "diagnostic_only": True,
        "target_scope": "all",
        "success": True,
        "integrity_pass": True,
        "axmath_count": len(texes),
        "omath_count": 0,
        "failed": [],
        "exports": [
            {"ordinal": i, "success": True, "axmath_exported_tex": tex}
            for i, tex in enumerate(texes, 1)
        ],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _formula_map(
    source: Path,
    final: Path,
    *,
    source_to_final: dict[int, list[int]] | None = None,
    added_targets: list[dict] | None = None,
) -> Path:
    source_inventory = static_math_inventory(source)
    source_count = int(source_inventory["axmath_count"] or source_inventory["omath_count"] or 0)
    final_count = int(static_math_inventory(final)["axmath_count"])
    mapping = source_to_final or {ordinal: [ordinal] for ordinal in range(1, source_count + 1)}
    source_texes = {
        1: r"\lim_{x\to0}f(x)",
        2: r"f\prime(x)",
        3: r"\lim^{n\to\infty}",
        4: r"\lim_{x\to0}^{+}",
        5: r"\frac{a}{b}+\sqrt[3]{x}+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\begin{cases}x&x\leq1\\0&x>1\end{cases}+\left|x\right|",
    }
    targets = [
        {
            "target_id": f"SRC-{source_ordinal:04d}",
            "kind": "source_formula",
            "source_ordinal": source_ordinal,
            "source_anchor": f"Frozen source formula #{source_ordinal}",
            "expected_tex": source_texes[source_ordinal],
            "final_ordinals": ordinals,
        }
        for source_ordinal, ordinals in sorted(mapping.items())
    ]
    targets.extend(added_targets or [])
    payload = {
        "schema": "axmath-source-final-map/v1",
        "source": {
            "path": str(source.resolve()),
            "sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
            "formula_count": source_count,
        },
        "final": {
            "path": str(final.resolve()),
            "sha256": hashlib.sha256(final.read_bytes()).hexdigest(),
            "axmath_count": final_count,
        },
        "targets": targets,
    }
    path = final.with_suffix(".formula-map.json")
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def main() -> int:
    assert normalize_axmath_tex(r"$\lim_{x\to0}f(x)$") == r"$\underset{x\to0}{\lim}f(x)$"
    assert normalize_axmath_tex(r"$\lim\limits_{x\to0}f(x)$") == r"$\underset{x\to0}{\lim}f(x)$"
    assert not limit_operator_issues(normalize_axmath_tex(r"$\lim_{x\to0}f(x)$"))
    assert normalize_axmath_tex(r"$\lim^{n\to\infty}a_n$") == r"$\lim^{n\to\infty}a_n$"
    assert normalize_axmath_tex(r"$\limsup_{n\to\infty}a_n$") == r"$\limsup_{n\to\infty}a_n$"
    assert normalize_axmath_tex(r"$\liminf_{n\to\infty}a_n$") == r"$\liminf_{n\to\infty}a_n$"
    assert limit_operator_placements(r"\underset{x\to0}{\lim}^{+}") == [["below", "above"]]
    assert limit_operator_issues(r"\lim_{x\to0}f(x)") == ["lim_subscript_requires_underset"]
    assert "lim_limits_command_not_allowed" in limit_operator_issues(r"\lim\limits_{x\to0}f(x)")
    assert normalize_axmath_tex(r"\smallint f(x)") == r"\int f(x)"
    assert smallint_issues(r"\smallint f(x)") == ["smallint_forbidden"]

    with tempfile.TemporaryDirectory() as folder:
        root = Path(folder)
        source = root / "source.docx"
        final = root / "final.docx"
        _source_docx(source)
        _final_docx(final)

        template_path = root / "SOURCE_FINAL_FORMULA_MAP.json"
        template = build_formula_map_template(source, template_path)
        assert template["source"]["formula_count"] == 5
        assert len(template["targets"]) == 5
        assert all(row["kind"] == "source_formula" for row in template["targets"])

        source_contract = scan_source_math(source)
        assert source_contract["limit_subscript_ordinals"] == [1, 3, 4]
        assert source_contract["limit_subscript_candidates"][0]["placements"] == ["below"]
        assert source_contract["limit_subscript_candidates"][1]["placements"] == ["above"]
        assert source_contract["limit_subscript_candidates"][2]["placements"] == ["below", "above"]
        assert source_contract["prime_semantic_candidates"][0]["prime_orders"] == [1]
        assert source_contract["math_structure_ordinals"] == [1, 3, 4, 5]

        original_converter = source_tex_map.convert_omml
        try:
            source_tex_map.convert_omml = lambda _node: r"\lim_{x\to0}f(x)"
            donor_map = source_tex_map.build_limit_layout_map(source, final, [1])
        finally:
            source_tex_map.convert_omml = original_converter
        assert donor_map["repair_class"] == "LIMIT_OPERATOR_LAYOUT_CONTRACT"
        assert donor_map["rows"][0]["provenance"] == "frozen_source_exact_omml"
        assert donor_map["rows"][0]["tex"] == r"$\underset{x\to0}{\lim}f(x)$"

        formula_map = _formula_map(source, final)
        audit = _internal_audit(
            final,
            r"\underset{x\to0}{\lim}f(x)",
            r"f\prime(x)",
            r"\lim^{n\to\infty}",
            r"\underset{x\to0}{\lim}^{+}",
            r"\frac{a}{b}+\sqrt[3]{x}+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\begin{cases}x&x\leq1\\0&x>1\end{cases}+\left|x\right|",
        )
        passed = validate_math_layout_contract(source, final, audit, formula_map=formula_map)
        assert passed["passed"], passed
        assert passed["mapping_pass"]
        assert passed["source_formula_count"] == passed["final_axmath_count"] == 5

        source_axmath = root / "source_axmath.docx"
        final_axmath = root / "final_axmath.docx"
        _final_docx(source_axmath)
        _final_docx(final_axmath)
        baseline_texes = (
            r"\underset{x\to0}{\lim}f(x)",
            r"f\prime(x)",
            r"\lim^{n\to\infty}",
            r"\underset{x\to0}{\lim}^{+}",
            r"\frac{a}{b}+\sqrt[3]{x}+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\begin{cases}x&x\leq1\\0&x>1\end{cases}+\left|x\right|",
        )
        source_baseline = _internal_audit(source_axmath, *baseline_texes)
        final_baseline = _internal_audit(final_axmath, *baseline_texes)
        source_axmath_map = _formula_map(source_axmath, final_axmath)
        existing_axmath_pass = validate_math_layout_contract(
            source_axmath,
            final_axmath,
            final_baseline,
            formula_map=source_axmath_map,
            source_internal_audit=source_baseline,
        )
        assert existing_axmath_pass["passed"], existing_axmath_pass

        bad_limit_audit = _internal_audit(
            final,
            r"\lim_{x\to0}f(x)",
            r"f\prime(x)",
            r"\lim^{n\to\infty}",
            r"\underset{x\to0}{\lim}^{+}",
            r"\frac{a}{b}+\sqrt[3]{x}+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\begin{cases}x&x\leq1\\0&x>1\end{cases}+\left|x\right|",
        )
        bad_limit = validate_math_layout_contract(source, final, bad_limit_audit, formula_map=formula_map)
        assert not bad_limit["passed"]
        assert any(x["reason"] == "lim_subscript_requires_underset" for x in bad_limit["defects"])

        bad_source_symbol_audit = _internal_audit(
            final,
            r"\underset{y\to0}{\lim}f(x)",
            r"f\prime(x)",
            r"\lim^{n\to\infty}",
            r"\underset{x\to0}{\lim}^{+}",
            r"\frac{a}{b}+\sqrt[3]{x}+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\begin{cases}x&x\leq1\\0&x>1\end{cases}+\left|x\right|",
        )
        bad_source_symbol = validate_math_layout_contract(
            source, final, bad_source_symbol_audit, formula_map=formula_map
        )
        assert not bad_source_symbol["passed"]
        assert any(
            x["reason"] == "source_formula_semantic_token_mismatch"
            for x in bad_source_symbol["defects"]
        )

        bad_prime_audit = _internal_audit(
            final,
            r"\underset{x\to0}{\lim}f(x)", "f′(x)",
            r"\lim^{n\to\infty}",
            r"\underset{x\to0}{\lim}^{+}",
            r"\frac{a}{b}+\sqrt[3]{x}+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\begin{cases}x&x\leq1\\0&x>1\end{cases}+\left|x\right|",
        )
        bad_prime = validate_math_layout_contract(source, final, bad_prime_audit, formula_map=formula_map)
        assert not bad_prime["passed"]
        assert any(x["reason"] == "noncanonical_axmath_prime_syntax" for x in bad_prime["defects"])

        wrong_prime_order_audit = _internal_audit(
            final,
            r"\underset{x\to0}{\lim}f(x)", "f''(x)",
            r"\lim^{n\to\infty}",
            r"\underset{x\to0}{\lim}^{+}",
            r"\frac{a}{b}+\sqrt[3]{x}+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\begin{cases}x&x\leq1\\0&x>1\end{cases}+\left|x\right|",
        )
        wrong_prime_order = validate_math_layout_contract(
            source, final, wrong_prime_order_audit, formula_map=formula_map
        )
        assert not wrong_prime_order["passed"]
        assert any(x["reason"] == "prime_order_mismatch" for x in wrong_prime_order["defects"])

        missing_structure_audit = _internal_audit(
            final,
            r"\underset{x\to0}{\lim}f(x)",
            r"f\prime(x)",
            r"\lim^{n\to\infty}",
            r"\underset{x\to0}{\lim}^{+}",
            r"a/b+\sqrt{x}+a+b+c+d+x\leq1+|x|",
        )
        missing_structure = validate_math_layout_contract(
            source, final, missing_structure_audit, formula_map=formula_map
        )
        assert not missing_structure["passed"]
        assert any(
            x["reason"] == "fractions_structure_count_mismatch"
            for x in missing_structure["defects"]
        )

        smallint_audit = _internal_audit(
            final,
            r"\underset{x\to0}{\lim}f(x)",
            r"f\prime(x)",
            r"\lim^{n\to\infty}",
            r"\underset{x\to0}{\lim}^{+}",
            r"\smallint_a^b f(x)\,dx",
        )
        smallint_result = validate_math_layout_contract(
            source, final, smallint_audit, formula_map=formula_map
        )
        assert not smallint_result["passed"]
        assert any(x["reason"] == "smallint_forbidden" for x in smallint_result["defects"])

        # One source formula is split into two AxMath objects, and a new body
        # variable/expression is added. Total source/final counts intentionally differ.
        split_final = root / "final_split.docx"
        _final_docx(split_final, count=8)
        split_audit = _internal_audit(
            split_final,
            r"\underset{x\to0}{\lim}",
            r"f(x)",
            r"f\prime(x)",
            r"\lim^{n\to\infty}",
            r"\underset{x\to0}{\lim}^{+}",
            r"\frac{a}{b}+\sqrt[3]{x}+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\begin{cases}x&x\leq1\\0&x>1\end{cases}+\left|x\right|",
            r"x^2+y",
            r"\int_0^1 f(x)\,dx",
        )
        split_map = _formula_map(
            source,
            split_final,
            source_to_final={1: [1, 2], 2: [3], 3: [4], 4: [5], 5: [6]},
            added_targets=[{
                "target_id": "BODY-001",
                "kind": "added_body_math",
                "source_anchor": "Paragraph 2, after the ordinary text 求",
                "expected_tex": "x^2+y",
                "final_ordinals": [7],
            }, {
                "target_id": "PDF-001",
                "kind": "source_formula_visual",
                "source_anchor": "PDF page 3, scanned equation region",
                "expected_tex": r"\int_0^1 f(x)\,dx",
                "final_ordinals": [8],
            }],
        )
        split_result = validate_math_layout_contract(
            source, split_final, split_audit, formula_map=split_map
        )
        assert split_result["passed"], split_result
        assert split_result["mapping_pass"]
        assert split_result["source_formula_count"] == 5
        assert split_result["final_axmath_count"] == 8
        assert split_result["additional_target_count"] == 2
        assert split_result["source_final_mapping"][0]["final_ordinals"] == [1, 2]

        bad_added_audit = _internal_audit(
            split_final,
            r"\underset{x\to0}{\lim}",
            r"f(x)",
            r"f\prime(x)",
            r"\lim^{n\to\infty}",
            r"\underset{x\to0}{\lim}^{+}",
            r"\frac{a}{b}+\sqrt[3]{x}+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\begin{cases}x&x\leq1\\0&x>1\end{cases}+\left|x\right|",
            r"y^2+y",
            r"\int_0^1 f(x)\,dx",
        )
        bad_added = validate_math_layout_contract(
            source, split_final, bad_added_audit, formula_map=split_map
        )
        assert not bad_added["passed"]
        assert any(
            x["reason"] == "added_target_semantic_token_mismatch"
            for x in bad_added["defects"]
        )

        broken_map = _formula_map(
            source,
            split_final,
            source_to_final={1: [1], 2: [3], 3: [4], 4: [5], 5: [6]},
            added_targets=[{
                "target_id": "BODY-001",
                "kind": "added_body_math",
                "source_anchor": "Paragraph 2, after the ordinary text 求",
                "expected_tex": "x^2+y",
                "final_ordinals": [7],
            }, {
                "target_id": "PDF-001",
                "kind": "source_formula_visual",
                "source_anchor": "PDF page 3, scanned equation region",
                "expected_tex": r"\int_0^1 f(x)\,dx",
                "final_ordinals": [8],
            }],
        )
        broken_result = validate_math_layout_contract(
            source, split_final, split_audit, formula_map=broken_map
        )
        assert not broken_result["passed"]
        assert not broken_result["mapping_pass"]
        assert any("unassigned=[2]" in x for x in broken_result["mapping_errors"])

    print("MATH_LAYOUT_CONTRACT_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
