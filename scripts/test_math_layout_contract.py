from __future__ import annotations

import hashlib
import json
import tempfile
import zipfile
from pathlib import Path

import build_source_tex_map as source_tex_map
from math_layout_contract import validate_math_layout_contract
from normalize_axmath_tex import (
    limit_operator_issues,
    limit_operator_placements,
    normalize_axmath_tex,
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
        '<m:rad><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>'
        '<m:m><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e></m:mr></m:m>'
        '<m:eqArr><m:e><m:r><m:t>x≤1</m:t></m:r></m:e></m:eqArr>'
        '<m:d><m:dPr><m:begChr m:val="|"/><m:endChr m:val="|"/></m:dPr>'
        '<m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>'
        '</m:oMath></w:p>'
        '</w:body></w:document>'
    )
    _write_docx(path, xml)


def _final_docx(path: Path) -> None:
    xml = (
        f'<w:document xmlns:w="{W}" xmlns:o="{O}" xmlns:r="{R}">'
        '<w:body>'
        '<w:p><w:r><w:object><o:OLEObject ProgID="Equation.AxMath" r:id="rId1"/>'
        '</w:object></w:r></w:p>'
        '<w:p><w:r><w:object><o:OLEObject ProgID="Equation.AxMath" r:id="rId2"/>'
        '</w:object></w:r></w:p>'
        '<w:p><w:r><w:object><o:OLEObject ProgID="Equation.AxMath" r:id="rId3"/>'
        '</w:object></w:r></w:p>'
        '<w:p><w:r><w:object><o:OLEObject ProgID="Equation.AxMath" r:id="rId4"/>'
        '</w:object></w:r></w:p>'
        '<w:p><w:r><w:object><o:OLEObject ProgID="Equation.AxMath" r:id="rId5"/>'
        '</w:object></w:r></w:p>'
        '</w:body></w:document>'
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


def main() -> int:
    assert normalize_axmath_tex(r"$\lim_{x\to0}f(x)$") == r"$\lim\limits_{x\to0}f(x)$"
    assert not limit_operator_issues(normalize_axmath_tex(r"$\lim_{x\to0}f(x)$"))
    assert normalize_axmath_tex(r"$\lim^{n\to\infty}a_n$") == r"$\lim\limits^{n\to\infty}a_n$"
    assert normalize_axmath_tex(r"$\limsup_{n\to\infty}a_n$") == r"$\limsup_{n\to\infty}a_n$"
    assert normalize_axmath_tex(r"$\liminf_{n\to\infty}a_n$") == r"$\liminf_{n\to\infty}a_n$"
    assert limit_operator_placements(r"\lim\limits_{x\to0}^{+}f(x)") == [["below", "above"]]

    with tempfile.TemporaryDirectory() as folder:
        root = Path(folder)
        source = root / "source.docx"
        final = root / "final.docx"
        _source_docx(source)
        _final_docx(final)

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
        assert donor_map["rows"][0]["tex"] == r"$\lim\limits_{x\to0}f(x)$"

        audit = _internal_audit(
            final,
            r"\lim\limits_{x\to0}f(x)",
            r"f\prime(x)",
            r"\lim\limits^{n\to\infty}a_n",
            r"\lim\limits_{x\to0}^{+}f(x)",
            r"\frac{a}{b}+\sqrt[3]{x}+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\begin{cases}x&x\leq1\\0&x>1\end{cases}+\left|x\right|",
        )
        passed = validate_math_layout_contract(source, final, audit)
        assert passed["passed"], passed

        source_axmath = root / "source_axmath.docx"
        final_axmath = root / "final_axmath.docx"
        _final_docx(source_axmath)
        _final_docx(final_axmath)
        baseline_texes = (
            r"\lim\limits_{x\to0}f(x)",
            r"f\prime(x)",
            r"\lim\limits^{n\to\infty}a_n",
            r"\lim\limits_{x\to0}^{+}f(x)",
            r"\frac{a}{b}+\sqrt[3]{x}+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\begin{cases}x&x\leq1\\0&x>1\end{cases}+\left|x\right|",
        )
        source_baseline = _internal_audit(source_axmath, *baseline_texes)
        final_baseline = _internal_audit(final_axmath, *baseline_texes)
        existing_axmath_pass = validate_math_layout_contract(
            source_axmath,
            final_axmath,
            final_baseline,
            source_internal_audit=source_baseline,
        )
        assert existing_axmath_pass["passed"], existing_axmath_pass

        bad_limit_audit = _internal_audit(
            final,
            r"\lim_{x\to0}f(x)",
            r"f\prime(x)",
            r"\lim\limits^{n\to\infty}a_n",
            r"\lim\limits_{x\to0}^{+}f(x)",
            r"\frac{a}{b}+\sqrt[3]{x}+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\begin{cases}x&x\leq1\\0&x>1\end{cases}+\left|x\right|",
        )
        bad_limit = validate_math_layout_contract(source, final, bad_limit_audit)
        assert not bad_limit["passed"]
        assert any(x["reason"] == "limit_script_not_forced_under_over" for x in bad_limit["defects"])

        bad_prime_audit = _internal_audit(
            final,
            r"\lim\limits_{x\to0}f(x)", "f′(x)",
            r"\lim\limits^{n\to\infty}a_n",
            r"\lim\limits_{x\to0}^{+}f(x)",
            r"\frac{a}{b}+\sqrt[3]{x}+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\begin{cases}x&x\leq1\\0&x>1\end{cases}+\left|x\right|",
        )
        bad_prime = validate_math_layout_contract(source, final, bad_prime_audit)
        assert not bad_prime["passed"]
        assert any(x["reason"] == "noncanonical_axmath_prime_syntax" for x in bad_prime["defects"])

        wrong_prime_order_audit = _internal_audit(
            final,
            r"\lim\limits_{x\to0}f(x)", "f''(x)",
            r"\lim\limits^{n\to\infty}a_n",
            r"\lim\limits_{x\to0}^{+}f(x)",
            r"\frac{a}{b}+\sqrt[3]{x}+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\begin{cases}x&x\leq1\\0&x>1\end{cases}+\left|x\right|",
        )
        wrong_prime_order = validate_math_layout_contract(source, final, wrong_prime_order_audit)
        assert not wrong_prime_order["passed"]
        assert any(x["reason"] == "prime_order_mismatch" for x in wrong_prime_order["defects"])

        missing_structure_audit = _internal_audit(
            final,
            r"\lim\limits_{x\to0}f(x)",
            r"f\prime(x)",
            r"\lim\limits^{n\to\infty}a_n",
            r"\lim\limits_{x\to0}^{+}f(x)",
            r"a/b+\sqrt{x}+a+b+c+d+x\leq1+|x|",
        )
        missing_structure = validate_math_layout_contract(source, final, missing_structure_audit)
        assert not missing_structure["passed"]
        assert any(x["reason"] == "fractions_structure_missing" for x in missing_structure["defects"])

    print("MATH_LAYOUT_CONTRACT_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
