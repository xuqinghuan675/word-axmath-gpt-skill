"""Fail-closed source-to-AxMath math-layout contract.

The contract is checked before conversion (source structure inventory), then
again against a complete read-only AMSAM2TeX export before any final delivery.
It prevents the classic inline-lim side-placement defect and derivative-prime
order drift from being silently accepted.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

M_NS = {"m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}
O_NS = "{urn:schemas-microsoft-com:office:office}"


def static_math_inventory(docx: Path) -> dict:
    with zipfile.ZipFile(docx) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    axmath = sum(
        1 for node in root.iter()
        if node.tag == O_NS + "OLEObject" and node.get("ProgID") == "Equation.AxMath"
    )
    return {"omath_count": len(root.findall(".//m:oMath", M_NS)), "axmath_count": axmath}
from internal_content_audit import validate_full_internal_audit
from normalize_axmath_tex import (
    canonical_prime_issues,
    limit_operator_count,
    limit_operator_issues,
    limit_operator_placements,
    normalize_axmath_tex,
    prime_orders,
)
from source_semantic_risks import scan_source_math

SCHEMA = "axmath-math-layout-contract/v1"
INTERNAL_SCHEMA = "axmath-internal-tex-diagnostic/v1"

_FRACTION_RE = re.compile(r"\\(?:dfrac|tfrac|cfrac|frac|binom)\b|\\over\b")
_RADICAL_RE = re.compile(r"\\sqrt\b")
_MATRIX_RE = re.compile(
    r"\\begin\s*\{(?:smallmatrix|[pbBvV]?matrix|array)\}|\\matrix\b"
)
_EQUATION_ARRAY_RE = re.compile(
    r"\\begin\s*\{(?:cases|dcases|rcases|aligned|alignedat|array)\}"
)
_SUBSCRIPT_RE = re.compile(r"_(?:\{[^{}]*\}|[A-Za-z0-9])")
_SUPERSCRIPT_RE = re.compile(r"\^(?:\{[^{}]*\}|[A-Za-z0-9])")
_RELATION_TEX_PATTERNS = {
    "le": re.compile(r"\\(?:le|leq|leqslant)(?![A-Za-z])|≤"),
    "ge": re.compile(r"\\(?:ge|geq|geqslant)(?![A-Za-z])|≥"),
    "ne": re.compile(r"\\(?:ne|neq)(?![A-Za-z])|≠"),
    "lt": re.compile(r"\\lt(?![A-Za-z])|(?<!\\)<"),
    "gt": re.compile(r"\\gt(?![A-Za-z])|(?<!\\)>"),
    "eq": re.compile(r"(?<![<>=])=(?![=])"),
    "plus": re.compile(r"\+"),
    "minus": re.compile(r"(?<!\\)-|−"),
    "pm": re.compile(r"\\pm(?![A-Za-z])|±"),
    "mp": re.compile(r"\\mp(?![A-Za-z])|∓"),
    "times": re.compile(r"\\times(?![A-Za-z])|×"),
    "divide": re.compile(r"\\div(?![A-Za-z])|÷"),
    "cdot": re.compile(r"\\cdot(?![A-Za-z])|·"),
}


def _tex_structure_signature(tex: str) -> dict:
    value = str(tex or "")
    relations = {
        key: len(pattern.findall(value))
        for key, pattern in _RELATION_TEX_PATTERNS.items()
        if pattern.search(value)
    }
    return {
        "fractions": len(_FRACTION_RE.findall(value)),
        "radicals": len(_RADICAL_RE.findall(value)),
        "matrices": len(_MATRIX_RE.findall(value)),
        "equation_arrays": len(_EQUATION_ARRAY_RE.findall(value)),
        "subscripts": len(_SUBSCRIPT_RE.findall(value)),
        "superscripts": len(_SUPERSCRIPT_RE.findall(value)),
        "relations": relations,
    }


def _delimiter_char_present(tex: str, char: str, side: str) -> bool:
    value = str(tex or "")
    if char == "|":
        token = (
            r"\\(?:lvert|left|bigl|Bigl|biggl|Biggl)\s*\|"
            if side == "begin"
            else r"\\(?:rvert|right|bigr|Bigr|biggr|Biggr)\s*\|"
        )
        if re.search(token, value):
            return True
        return len(re.findall(r"(?<!\\)\|", value)) >= 2 or len(re.findall(r"\\vert(?![A-Za-z])", value)) >= 2
    if char in {"{", "}"}:
        token = r"\\\{" if char == "{" else r"\\\}"
        return bool(re.search(token, value))
    if char in {"(", ")", "[", "]"}:
        return char in value
    aliases = {"⟨": r"\\langle", "⟩": r"\\rangle", "⌊": r"\\lfloor", "⌋": r"\\rfloor", "⌈": r"\\lceil", "⌉": r"\\rceil"}
    if char in aliases:
        return bool(re.search(aliases[char], value)) or char in value
    return char in value


def _structure_defects(expected: dict, tex: str) -> list[dict]:
    actual = _tex_structure_signature(tex)
    defects: list[dict] = []
    for key in ("fractions", "radicals", "matrices", "equation_arrays", "subscripts", "superscripts"):
        want = int(expected.get(key) or 0)
        got = int(actual.get(key) or 0)
        if got < want:
            defects.append({"reason": f"{key}_structure_missing", "expected_at_least": want, "actual": got})
    want_subsup = int(expected.get("subsuperscripts") or 0)
    if want_subsup:
        paired = min(int(actual.get("subscripts") or 0), int(actual.get("superscripts") or 0))
        if paired < want_subsup:
            defects.append({"reason": "subsuperscript_structure_missing", "expected_at_least": want_subsup, "actual": paired})
    actual_relations = actual.get("relations") or {}
    for token, want_value in (expected.get("relations") or {}).items():
        want = int(want_value or 0)
        got = int(actual_relations.get(token) or 0)
        if got < want:
            defects.append({"reason": f"relation_{token}_missing", "expected_at_least": want, "actual": got})
    for delimiter in expected.get("delimiters") or []:
        begin = str(delimiter.get("begin") or "")
        end = str(delimiter.get("end") or "")
        if begin and not _delimiter_char_present(tex, begin, "begin"):
            defects.append({"reason": "opening_delimiter_missing", "delimiter": begin})
        if end and not _delimiter_char_present(tex, end, "end"):
            defects.append({"reason": "closing_delimiter_missing", "delimiter": end})
    return defects


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _read_full_audit(path: Path, docx: Path) -> dict:
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    actual_sha = sha256_file(docx)
    if data.get("schema") != INTERNAL_SCHEMA:
        raise ValueError(f"unsupported internal audit schema: {path}")
    if data.get("target_scope") != "all" or data.get("diagnostic_only") is not True:
        raise ValueError("internal audit must be a complete read-only -AllAxMath export")
    if Path(str(data.get("input") or "")).resolve() != docx.resolve():
        raise ValueError("internal audit belongs to a different DOCX")
    if data.get("input_sha256_before") != actual_sha or data.get("input_sha256_after") != actual_sha:
        raise ValueError("internal audit hash is stale or DOCX changed during export")
    if data.get("input_unchanged") is not True or data.get("success") is not True:
        raise ValueError("internal audit did not complete as a read-only successful export")
    if data.get("failed"):
        raise ValueError("internal audit contains failed exports")
    exports = data.get("exports")
    if not isinstance(exports, list):
        raise ValueError("internal audit has no exports array")
    return data


def _source_axmath_contract(source: Path, source_internal_audit: Path) -> dict:
    data = _read_full_audit(source_internal_audit, source)
    rows = data["exports"]
    limits: dict[int, int] = {}
    primes: dict[int, list[int]] = {}
    structures: dict[int, dict] = {}
    for row in rows:
        ordinal = int(row["ordinal"])
        raw = str(row.get("axmath_exported_tex") or "")
        normalized = normalize_axmath_tex(raw)
        count = limit_operator_count(normalized)
        if count:
            limits[ordinal] = count
        orders = prime_orders(normalized)
        if orders:
            primes[ordinal] = orders
        signature = _tex_structure_signature(normalized)
        if any(signature.values()):
            structures[ordinal] = signature
    return {
        "source_kind": "existing_axmath_internal_baseline",
        "source_formula_count": int(data.get("axmath_count") or len(rows)),
        "limit_targets": [
            {
                "ordinal": ordinal,
                "count": count,
                "placements": limit_operator_placements(
                    str(next(x for x in rows if int(x["ordinal"]) == ordinal).get("axmath_exported_tex") or "")
                ),
            }
            for ordinal, count in sorted(limits.items())
        ],
        "prime_targets": [
            {"ordinal": ordinal, "prime_orders": orders}
            for ordinal, orders in sorted(primes.items())
        ],
        "structure_targets": [
            {"ordinal": ordinal, "structure": signature}
            for ordinal, signature in sorted(structures.items())
        ],
        "baseline_internal_audit": str(source_internal_audit.resolve()),
        "baseline_internal_audit_sha256": sha256_file(source_internal_audit),
    }


def build_source_contract(
    source: Path, source_internal_audit: Path | None = None
) -> dict:
    source = source.resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    inventory = static_math_inventory(source)
    if int(inventory.get("axmath_count") or 0) and int(inventory.get("omath_count") or 0):
        raise ValueError(
            "mixed AxMath and native OfficeMath in the same frozen source requires a separately reviewed "
            "source-to-final mapping; refusing to silently ignore either formula inventory"
        )
    semantic = scan_source_math(source)
    source_sha = sha256_file(source)

    if int(semantic.get("formula_count") or 0) > 0:
        grouped: dict[int, list[dict]] = defaultdict(list)
        for row in semantic.get("limit_subscript_candidates") or []:
            grouped[int(row["ordinal"])].append(row)
        limit_targets = [
            {
                "ordinal": ordinal,
                "count": len(rows),
                "placements": [list(row.get("placements") or []) for row in rows],
                "source_nodes": [str(row["source_node"]) for row in rows],
                "source_scripts": [str(row.get("script_text") or "") for row in rows],
            }
            for ordinal, rows in sorted(grouped.items())
        ]
        prime_targets = [
            {
                "ordinal": int(row["ordinal"]),
                "prime_orders": list(row.get("prime_orders") or []),
                "prime_symbols": list(row.get("prime_symbols") or []),
            }
            for row in semantic.get("prime_semantic_candidates") or []
        ]
        structure_targets = [
            {
                "ordinal": int(row["ordinal"]),
                "structure": dict(row.get("structure") or {}),
            }
            for row in semantic.get("math_structure_candidates") or []
        ]
        contract = {
            "source_kind": "native_officemath_omml",
            "source_formula_count": int(semantic["formula_count"]),
            "limit_targets": limit_targets,
            "prime_targets": prime_targets,
            "structure_targets": structure_targets,
        }
    elif int(inventory.get("axmath_count") or 0) > 0:
        if source_internal_audit is None:
            raise ValueError(
                "Source is already AxMath. Export a complete read-only source "
                "baseline with inspect_axmath_tex.ps1 -AllAxMath before editing."
            )
        contract = _source_axmath_contract(source, source_internal_audit)
    else:
        contract = {
            "source_kind": "no_math",
            "source_formula_count": 0,
            "limit_targets": [],
            "prime_targets": [],
            "structure_targets": [],
        }

    return {
        "schema": SCHEMA,
        "source": str(source),
        "source_sha256": source_sha,
        "source_omath_count": int(inventory.get("omath_count") or 0),
        "source_axmath_count": int(inventory.get("axmath_count") or 0),
        "limit_target_count": len(contract["limit_targets"]),
        "limit_operator_count": sum(int(x["count"]) for x in contract["limit_targets"]),
        "prime_target_count": len(contract["prime_targets"]),
        "structure_target_count": len(contract.get("structure_targets") or []),
        **contract,
    }


def validate_math_layout_contract(
    source: Path,
    final: Path,
    internal_audit: Path,
    *,
    source_internal_audit: Path | None = None,
    repair_ledger: Path | None = None,
) -> dict:
    source = source.resolve()
    final = final.resolve()
    internal_audit = internal_audit.resolve()
    contract = build_source_contract(source, source_internal_audit)
    final_inventory = static_math_inventory(final)
    final_count = int(final_inventory.get("axmath_count") or 0)
    audit_state = validate_full_internal_audit(internal_audit, final, final_count)
    data = _read_full_audit(internal_audit, final)
    exports = {
        int(row["ordinal"]): row
        for row in data.get("exports") or []
        if isinstance(row, dict) and row.get("ordinal") is not None
    }

    errors: list[str] = []
    defects: list[dict] = []
    if not audit_state["passed"]:
        errors.extend(audit_state["errors"])

    # Global gate: no scripted lim may rely on inline/display context to decide
    # where its subscript appears. Every such operator must carry \limits.
    for ordinal, row in sorted(exports.items()):
        tex = str(row.get("axmath_exported_tex") or "")
        for issue in limit_operator_issues(tex):
            defects.append({"ordinal": ordinal, "reason": issue})
    if defects:
        errors.append("one or more lim scripts are not explicitly positioned with \\limits")

    mapping_skipped = bool(repair_ledger is not None and final_count != contract["source_formula_count"])
    if final_count != contract["source_formula_count"] and not repair_ledger:
        errors.append(
            f"source/final formula mapping is not one-to-one: "
            f"{contract['source_formula_count']} source formulas vs {final_count} final AxMath objects"
        )
    if mapping_skipped and any((
        contract["limit_targets"], contract["prime_targets"], contract.get("structure_targets")
    )):
        errors.append(
            "source/final formula count changed under a repair ledger, so ordinal-level "
            "limit/prime/structure matching is unavailable; do not waive this gate without "
            "an explicit source-to-final formula mapping"
        )

    if not mapping_skipped:
        for target in contract["limit_targets"]:
            ordinal = int(target["ordinal"])
            row = exports.get(ordinal)
            if row is None:
                defects.append({"ordinal": ordinal, "reason": "limit_target_export_missing"})
                continue
            tex = str(row.get("axmath_exported_tex") or "")
            expected = int(target["count"])
            actual = limit_operator_count(tex)
            if actual != expected:
                defects.append({
                    "ordinal": ordinal,
                    "reason": "limit_operator_count_mismatch",
                    "expected": expected,
                    "actual": actual,
                })
            expected_placements = [list(x) for x in target.get("placements") or []]
            actual_placements = limit_operator_placements(tex)
            if actual_placements != expected_placements:
                defects.append({
                    "ordinal": ordinal,
                    "reason": "limit_operator_placement_mismatch",
                    "expected": expected_placements,
                    "actual": actual_placements,
                })

        for target in contract["prime_targets"]:
            ordinal = int(target["ordinal"])
            row = exports.get(ordinal)
            if row is None:
                defects.append({"ordinal": ordinal, "reason": "prime_target_export_missing"})
                continue
            tex = str(row.get("axmath_exported_tex") or "")
            expected = list(target.get("prime_orders") or [])
            if not expected:
                defects.append({"ordinal": ordinal, "reason": "source_prime_order_unresolved"})
                continue
            issues = canonical_prime_issues(tex)
            actual = prime_orders(tex)
            if issues:
                defects.append({
                    "ordinal": ordinal,
                    "reason": "noncanonical_axmath_prime_syntax",
                    "issues": issues,
                })
            if actual != expected:
                defects.append({
                    "ordinal": ordinal,
                    "reason": "prime_order_mismatch",
                    "expected": expected,
                    "actual": actual,
                })

        for target in contract.get("structure_targets") or []:
            ordinal = int(target["ordinal"])
            row = exports.get(ordinal)
            if row is None:
                defects.append({"ordinal": ordinal, "reason": "math_structure_target_export_missing"})
                continue
            tex = str(row.get("axmath_exported_tex") or "")
            for defect in _structure_defects(dict(target.get("structure") or {}), tex):
                defects.append({"ordinal": ordinal, **defect})

    if defects and not errors:
        errors.append(f"{len(defects)} math-layout semantic defects detected")
    return {
        "schema": SCHEMA,
        "passed": not errors and not defects,
        "errors": errors,
        "defects": defects,
        "source": str(source),
        "final": str(final),
        "source_sha256": contract["source_sha256"],
        "final_sha256": sha256_file(final),
        "internal_audit": str(internal_audit),
        "internal_audit_sha256": sha256_file(internal_audit),
        "source_kind": contract["source_kind"],
        "source_formula_count": contract["source_formula_count"],
        "final_axmath_count": final_count,
        "limit_targets": contract["limit_target_count"],
        "limit_operators": contract["limit_operator_count"],
        "prime_targets": contract["prime_target_count"],
        "structure_targets": contract.get("structure_target_count", 0),
        "ordinal_mapping_skipped_due_to_repair_ledger": mapping_skipped,
        "repair_ledger": str(repair_ledger.resolve()) if repair_ledger else None,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build and enforce source-to-AxMath limit/prime layout contracts."
    )
    sub = parser.add_subparsers(dest="mode", required=True)
    preflight = sub.add_parser("preflight")
    preflight.add_argument("--source", required=True)
    preflight.add_argument("--source-internal-audit")
    preflight.add_argument("--out", required=True)

    validate = sub.add_parser("validate")
    validate.add_argument("--source", required=True)
    validate.add_argument("--final", required=True)
    validate.add_argument("--internal-audit", required=True)
    validate.add_argument("--source-internal-audit")
    validate.add_argument("--repair-ledger")
    validate.add_argument("--out", required=True)
    args = parser.parse_args()

    if args.mode == "preflight":
        payload = build_source_contract(
            Path(args.source),
            Path(args.source_internal_audit) if args.source_internal_audit else None,
        )
    else:
        payload = validate_math_layout_contract(
            Path(args.source),
            Path(args.final),
            Path(args.internal_audit),
            source_internal_audit=(
                Path(args.source_internal_audit) if args.source_internal_audit else None
            ),
            repair_ledger=Path(args.repair_ledger) if args.repair_ledger else None,
        )

    out = Path(args.out).resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "mode": args.mode,
        "passed": payload.get("passed", True),
        "source_kind": payload.get("source_kind"),
        "limit_targets": payload.get("limit_target_count", payload.get("limit_targets")),
        "prime_targets": payload.get("prime_target_count", payload.get("prime_targets")),
        "out": str(out),
        "errors": payload.get("errors", []),
    }, ensure_ascii=True, indent=2))
    return 0 if payload.get("passed", True) else 1


if __name__ == "__main__":
    raise SystemExit(main())
