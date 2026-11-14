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
    smallint_issues,
)
from source_semantic_risks import scan_source_math

SCHEMA = "axmath-math-layout-contract/v1"
INTERNAL_SCHEMA = "axmath-internal-tex-diagnostic/v1"
FORMULA_MAP_SCHEMA = "axmath-source-final-map/v1"

_FRACTION_RE = re.compile(r"\\(?:dfrac|tfrac|cfrac|frac|binom)\b|\\over\b")
_RADICAL_RE = re.compile(r"\\sqrt\b")
_MATRIX_RE = re.compile(
    r"\\begin\s*\{(?:smallmatrix|[pbBvV]?matrix|array)\}|\\matrix\b"
)
_EQUATION_ARRAY_RE = re.compile(
    r"\\begin\s*\{(?:cases|dcases|rcases|aligned|alignedat|array)\}"
)
_SUBSCRIPT_RE = re.compile(r"(?<!\\)_(?:\{[^{}]*\}|[A-Za-z0-9])")
_SUPERSCRIPT_RE = re.compile(r"(?<!\\)\^(?:\{[^{}]*\}|[A-Za-z0-9])")
_UNDERSCRIPT_RE = re.compile(r"\\underset\s*\{")
_SUBSUP_RE = re.compile(
    r"_(?:\{[^{}]*\}|[A-Za-z0-9])\s*\^(?:\{[^{}]*\}|[A-Za-z0-9])"
)
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
_OPERATOR_TEX_PATTERNS = {
    "integral": re.compile(r"\\int(?![A-Za-z])|∫"),
    "double_integral": re.compile(r"\\iint(?![A-Za-z])|∬"),
    "triple_integral": re.compile(r"\\iiint(?![A-Za-z])|∭"),
    "contour_integral": re.compile(r"\\oint(?![A-Za-z])|∮"),
    "sum": re.compile(r"\\sum(?![A-Za-z])|∑"),
    "product": re.compile(r"\\prod(?![A-Za-z])|∏"),
    "coproduct": re.compile(r"\\coprod(?![A-Za-z])|∐"),
    "union": re.compile(r"\\(?:big)?cup(?![A-Za-z])|∪"),
    "intersection": re.compile(r"\\(?:big)?cap(?![A-Za-z])|∩"),
    "membership": re.compile(r"\\in(?![A-Za-z])|∈"),
    "nonmembership": re.compile(r"\\notin(?![A-Za-z])|∉"),
    "arrow": re.compile(r"\\to(?![A-Za-z])|→"),
    "mapsto": re.compile(r"\\mapsto(?![A-Za-z])|↦"),
    "implies": re.compile(r"\\Rightarrow(?![A-Za-z])|⇒"),
    "equivalence": re.compile(r"\\Leftrightarrow(?![A-Za-z])|⇔"),
    "infinity": re.compile(r"\\infty(?![A-Za-z])|∞"),
}


def _tex_structure_signature(tex: str) -> dict:
    value = str(tex or "")
    relations = {
        key: len(pattern.findall(value))
        for key, pattern in _RELATION_TEX_PATTERNS.items()
        if pattern.search(value)
    }
    operators = {
        key: len(pattern.findall(value))
        for key, pattern in _OPERATOR_TEX_PATTERNS.items()
        if pattern.search(value)
    }
    subscripts = len(_SUBSCRIPT_RE.findall(value)) + len(_UNDERSCRIPT_RE.findall(value))
    superscripts = len(_SUPERSCRIPT_RE.findall(value))
    lim_pairs = sum(
        1 for directions in limit_operator_placements(value)
        if "below" in directions and "above" in directions
    )
    return {
        "fractions": len(_FRACTION_RE.findall(value)),
        "radicals": len(_RADICAL_RE.findall(value)),
        "matrices": len(_MATRIX_RE.findall(value)),
        "equation_arrays": len(_EQUATION_ARRAY_RE.findall(value)),
        "subscripts": subscripts,
        "superscripts": superscripts,
        "subsuperscripts": len(_SUBSUP_RE.findall(value)) + lim_pairs,
        "relations": relations,
        "operators": operators,
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


_LAYOUT_COMMANDS = {
    "frac", "dfrac", "tfrac", "sqrt", "root", "of", "begin", "end",
    "left", "right", "underset", "overset", "limits", "nolimits",
    "displaystyle", "textstyle", "scriptstyle", "scriptscriptstyle",
    "quad", "qquad", "big", "Big", "bigg", "Bigg", "bigl", "bigr",
    "Bigl", "Bigr", "biggl", "biggr", "Biggl", "Biggr",
    "mathrm", "mathbf", "mathit", "mathbb", "mathcal", "mathfrak",
    "operatorname", "text", "boldsymbol", "overline", "underline",
    "overbrace", "underbrace", "binom", "dbinom", "tbinom", "over",
    "atop", "choose", "array", "matrix", "pmatrix", "bmatrix",
    "Bmatrix", "vmatrix", "Vmatrix", "cases", "aligned", "gathered",
    "substack", "hphantom", "vphantom", "phantom", "smash", "notag", "tag",
    "thinspace", "medspace", "thickspace",
}
_SEMANTIC_REPLACEMENTS = [
    (r"\Leftrightarrow", "⇔"), (r"\Rightarrow", "⇒"),
    (r"\mapsto", "↦"), (r"\notin", "∉"), (r"\infty", "∞"),
    (r"\iiint", "∭"), (r"\iint", "∬"), (r"\oint", "∮"),
    (r"\int", "∫"), (r"\sum", "∑"), (r"\prod", "∏"),
    (r"\coprod", "∐"), (r"\bigcup", "∪"), (r"\cup", "∪"),
    (r"\bigcap", "∩"), (r"\cap", "∩"),
    (r"\leqslant", "≤"), (r"\leq", "≤"), (r"\le", "≤"),
    (r"\geqslant", "≥"), (r"\geq", "≥"), (r"\ge", "≥"),
    (r"\neq", "≠"), (r"\ne", "≠"), (r"\to", "→"),
    (r"\pm", "±"), (r"\mp", "∓"), (r"\times", "×"),
    (r"\div", "÷"), (r"\cdot", "·"), (r"\in", "∈"),
    (r"\lvert", "|"), (r"\rvert", "|"), (r"\vert", "|"),
    (r"\lVert", "‖"), (r"\rVert", "‖"), (r"\Vert", "‖"),
    (r"\langle", "⟨"), (r"\rangle", "⟩"),
    (r"\prime", " PRIME "), ("′", " PRIME "),
]
_GREEK_CANONICAL = {
    r"\alpha": "α", r"\beta": "β", r"\gamma": "γ", r"\delta": "δ",
    r"\epsilon": "ε", r"\varepsilon": "ε", r"\zeta": "ζ", r"\eta": "η",
    r"\theta": "θ", r"\vartheta": "θ", r"\iota": "ι", r"\kappa": "κ",
    r"\lambda": "λ", r"\mu": "μ", r"\nu": "ν", r"\xi": "ξ",
    r"\pi": "π", r"\varpi": "π", r"\rho": "ρ", r"\varrho": "ρ",
    r"\sigma": "σ", r"\varsigma": "σ", r"\tau": "τ", r"\upsilon": "υ",
    r"\phi": "φ", r"\varphi": "φ", r"\chi": "χ", r"\psi": "ψ",
    r"\omega": "ω", r"\Gamma": "Γ", r"\Delta": "Δ", r"\Theta": "Θ",
    r"\Lambda": "Λ", r"\Xi": "Ξ", r"\Pi": "Π", r"\Sigma": "Σ",
    r"\Upsilon": "Υ", r"\Phi": "Φ", r"\Psi": "Ψ", r"\Omega": "Ω",
}


def _semantic_tokens(tex: str) -> list[str]:
    value = normalize_axmath_tex(str(tex or ""))
    value = re.sub(r"\\(?:begin|end)\s*\{[^{}]*\}", " ", value)
    for command, canonical in _SEMANTIC_REPLACEMENTS:
        value = value.replace(command, canonical)
    for command, canonical in _GREEK_CANONICAL.items():
        value = value.replace(command, canonical)
    value = value.replace("\\\\", " ")
    tokens: list[str] = []
    for token in re.findall(r"\\[A-Za-z]+|\\.|[A-Za-z]+|\d+|[^\s{}]", value):
        if token.startswith("\\"):
            name = token[1:]
            if name in _LAYOUT_COMMANDS:
                continue
            tokens.append(name)
        else:
            tokens.append(token)
    return tokens


def _structure_defects(expected: dict, tex: str, *, exact: bool = False) -> list[dict]:
    actual = _tex_structure_signature(tex)
    defects: list[dict] = []
    numeric_keys = (
        "fractions", "radicals", "matrices", "equation_arrays",
        "subscripts", "superscripts", "subsuperscripts",
    )
    for key in numeric_keys:
        want = int(expected.get(key) or 0)
        got = int(actual.get(key) or 0)
        if (exact and got != want) or (not exact and got < want):
            defects.append({
                "reason": f"{key}_structure_count_mismatch" if exact else f"{key}_structure_missing",
                "expected": want if exact else None,
                "expected_at_least": None if exact else want,
                "actual": got,
            })
    for field in ("relations", "operators"):
        expected_counts = expected.get(field) or {}
        actual_counts = actual.get(field) or {}
        tokens = set(expected_counts) | (set(actual_counts) if exact else set())
        for token in sorted(tokens):
            want = int(expected_counts.get(token) or 0)
            got = int(actual_counts.get(token) or 0)
            if (exact and got != want) or (not exact and got < want):
                defects.append({
                    "reason": f"{field}_{token}_count_mismatch" if exact else f"{field}_{token}_missing",
                    "expected": want if exact else None,
                    "expected_at_least": None if exact else want,
                    "actual": got,
                })
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


def build_formula_map_template(source: Path, out: Path) -> dict:
    source = source.resolve()
    inventory = static_math_inventory(source)
    if inventory["omath_count"] and inventory["axmath_count"]:
        raise ValueError("mixed OfficeMath and AxMath source requires manual review before mapping")
    source_count = int(inventory["axmath_count"] or inventory["omath_count"] or 0)
    source_kind = "axmath" if inventory["axmath_count"] else ("officemath" if inventory["omath_count"] else "no_math")
    payload = {
        "schema": FORMULA_MAP_SCHEMA,
        "source": {
            "path": str(source),
            "sha256": sha256_file(source),
            "formula_count": source_count,
            "formula_kind": source_kind,
        },
        "final": {
            "path": "",
            "sha256": "",
            "axmath_count": None,
        },
        "targets": [
            {
                "target_id": f"SRC-{ordinal:04d}",
                "kind": "source_formula",
                "source_ordinal": ordinal,
                "source_anchor": "",
                "expected_tex": "",
                "final_ordinals": [],
            }
            for ordinal in range(1, source_count + 1)
        ],
        "additional_target_examples": [
            {
                "target_id": "PDF-001",
                "kind": "source_formula_visual",
                "source_anchor": "PDF page 2, equation region near the upper-right diagram",
                "expected_tex": "\\int_0^1 f(x)\\,dx",
                "final_ordinals": [],
            },
            {
                "target_id": "BODY-001",
                "kind": "added_body_math",
                "source_anchor": "Paragraph/line and nearby ordinary text locating the variable or expression",
                "expected_tex": "x",
                "final_ordinals": [],
            },
        ],
        "instructions": [
            "Create this file before generating the DOCX; fill source_anchor and expected_tex for every source_formula from the frozen source, never from the final export.",
            "Add source_formula_visual targets for PDF/image formulas not represented as source OMML/AxMath, and added_body_math targets for every body variable/expression that must become an AxMath OLE.",
            "After generation, fill final.path, final.sha256, final.axmath_count and every target.final_ordinals.",
            "A source_formula may map to multiple final ordinals when split; every final AxMath ordinal must belong to exactly one target.",
            "For visual-source and body-math targets, provide a precise source_anchor and source-derived expected_tex; keep surrounding Chinese instructions and question labels as ordinary Word text.",
        ],
    }
    out = out.resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return payload


def _load_formula_map(
    path: Path | None,
    source: Path,
    final: Path,
    source_count: int,
    final_count: int,
) -> tuple[dict, dict[int, list[int]], list[dict], list[str]]:
    if path is None:
        return {}, {}, [], ["an explicit source-to-final formula map is required"]
    path = path.resolve()
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    errors: list[str] = []
    if data.get("schema") != FORMULA_MAP_SCHEMA:
        errors.append(f"unsupported formula-map schema: {data.get('schema')!r}")
    source_meta = data.get("source") or {}
    final_meta = data.get("final") or {}
    if Path(str(source_meta.get("path") or "")).resolve() != source:
        errors.append("formula map belongs to a different source DOCX")
    if str(source_meta.get("sha256") or "").lower() != sha256_file(source).lower():
        errors.append("formula map source hash is stale")
    source_formula_count = source_meta.get("formula_count")
    if source_formula_count is None or int(source_formula_count) != source_count:
        errors.append("formula map source formula count is stale")
    if Path(str(final_meta.get("path") or "")).resolve() != final:
        errors.append("formula map belongs to a different final DOCX")
    if str(final_meta.get("sha256") or "").lower() != sha256_file(final).lower():
        errors.append("formula map final hash is stale")
    final_axmath_count = final_meta.get("axmath_count")
    if final_axmath_count is None or int(final_axmath_count) != final_count:
        errors.append("formula map final AxMath count is stale")

    targets = data.get("targets")
    if not isinstance(targets, list):
        return data, {}, [], errors + ["formula map has no targets array"]
    source_to_final: dict[int, list[int]] = {}
    added_targets: list[dict] = []
    final_owner: dict[int, str] = {}
    target_ids: set[str] = set()
    for target in targets:
        if not isinstance(target, dict):
            errors.append("formula map contains a non-object target")
            continue
        target_id = str(target.get("target_id") or "").strip()
        kind = str(target.get("kind") or "")
        if not target_id or target_id in target_ids:
            errors.append(f"formula map target_id is missing or duplicated: {target_id!r}")
        target_ids.add(target_id)
        raw_ordinals = target.get("final_ordinals")
        if not isinstance(raw_ordinals, list) or not raw_ordinals:
            errors.append(f"{target_id or '<unnamed>'} has no final_ordinals")
            continue
        try:
            ordinals = [int(x) for x in raw_ordinals]
        except (TypeError, ValueError):
            errors.append(f"{target_id} contains a non-integer final ordinal")
            continue
        if len(ordinals) != len(set(ordinals)):
            errors.append(f"{target_id} repeats a final ordinal")
        for ordinal in ordinals:
            if ordinal < 1 or ordinal > final_count:
                errors.append(f"{target_id} maps to out-of-range final AxMath ordinal {ordinal}")
            if ordinal in final_owner:
                errors.append(
                    f"final AxMath ordinal {ordinal} is assigned to both "
                    f"{final_owner[ordinal]} and {target_id}"
                )
            final_owner[ordinal] = target_id
        if kind == "source_formula":
            try:
                source_ordinal = int(target.get("source_ordinal"))
            except (TypeError, ValueError):
                errors.append(f"{target_id} has an invalid source_ordinal")
                continue
            if source_ordinal < 1 or source_ordinal > source_count:
                errors.append(f"{target_id} maps an out-of-range source formula ordinal {source_ordinal}")
            if source_ordinal in source_to_final:
                errors.append(f"source formula ordinal {source_ordinal} is mapped more than once")
            if not str(target.get("source_anchor") or "").strip():
                errors.append(f"{target_id} has no source_anchor")
            if not str(target.get("expected_tex") or "").strip():
                errors.append(f"{target_id} has no source-derived expected_tex")
            source_to_final[source_ordinal] = ordinals
        elif kind in {"source_formula_visual", "added_body_math"}:
            if not str(target.get("source_anchor") or "").strip():
                errors.append(f"{target_id} has no source_anchor")
            if not str(target.get("expected_tex") or "").strip():
                errors.append(f"{target_id} has no expected_tex")
            added_targets.append({**target, "final_ordinals": ordinals})
        else:
            errors.append(f"{target_id} has unsupported target kind {kind!r}")

    expected_sources = set(range(1, source_count + 1))
    if set(source_to_final) != expected_sources:
        missing = sorted(expected_sources - set(source_to_final))
        extra = sorted(set(source_to_final) - expected_sources)
        errors.append(f"source formula mapping coverage mismatch; missing={missing}, extra={extra}")
    expected_finals = set(range(1, final_count + 1))
    if set(final_owner) != expected_finals:
        missing = sorted(expected_finals - set(final_owner))
        extra = sorted(set(final_owner) - expected_finals)
        errors.append(f"final AxMath mapping coverage mismatch; unassigned={missing}, out_of_range={extra}")
    return data, source_to_final, added_targets, errors


def validate_math_layout_contract(
    source: Path,
    final: Path,
    internal_audit: Path,
    *,
    formula_map: Path | None = None,
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

    formula_map_data, source_to_final, added_targets, mapping_errors = _load_formula_map(
        formula_map,
        source,
        final,
        int(contract["source_formula_count"]),
        final_count,
    )
    mapping_pass = not mapping_errors
    errors: list[str] = list(mapping_errors)
    defects: list[dict] = []
    if not audit_state["passed"]:
        errors.extend(audit_state["errors"])

    for ordinal, row in sorted(exports.items()):
        tex = str(row.get("axmath_exported_tex") or "")
        for issue in limit_operator_issues(tex):
            defects.append({"ordinal": ordinal, "reason": issue})
        for issue in smallint_issues(tex):
            defects.append({"ordinal": ordinal, "reason": issue})
    if any(x.get("reason") in {"lim_subscript_requires_underset", "lim_limits_command_not_allowed"} for x in defects):
        errors.append("one or more lower lim scripts do not use the AxMath-stable \\underset form")
    if any(x.get("reason") == "smallint_forbidden" for x in defects):
        errors.append("one or more final AxMath exports use forbidden \\smallint")

    def mapped_tex(ordinals: list[int], target_id: str) -> str:
        missing = [ordinal for ordinal in ordinals if ordinal not in exports]
        if missing:
            defects.append({
                "target_id": target_id,
                "reason": "mapped_final_export_missing",
                "missing_ordinals": missing,
            })
        return "\n".join(
            str(exports[ordinal].get("axmath_exported_tex") or "")
            for ordinal in ordinals
            if ordinal in exports
        )

    if mapping_pass:
        source_target_by_ordinal = {
            int(target["source_ordinal"]): target
            for target in formula_map_data.get("targets") or []
            if isinstance(target, dict) and target.get("kind") == "source_formula"
        }
        for source_ordinal, ordinals in sorted(source_to_final.items()):
            target = source_target_by_ordinal[source_ordinal]
            expected_tex = normalize_axmath_tex(str(target.get("expected_tex") or ""))
            actual_tex = mapped_tex(ordinals, str(target.get("target_id") or f"SRC-{source_ordinal:04d}"))
            expected_tokens = _semantic_tokens(expected_tex)
            actual_tokens = _semantic_tokens(actual_tex)
            if expected_tokens != actual_tokens:
                defects.append({
                    "source_ordinal": source_ordinal,
                    "final_ordinals": ordinals,
                    "reason": "source_formula_semantic_token_mismatch",
                    "expected_tokens": expected_tokens,
                    "actual_tokens": actual_tokens,
                })

        for target in contract["limit_targets"]:
            source_ordinal = int(target["ordinal"])
            ordinals = source_to_final.get(source_ordinal, [])
            tex = mapped_tex(ordinals, f"SRC-{source_ordinal:04d}")
            expected_count = int(target["count"])
            actual_count = limit_operator_count(tex)
            if actual_count != expected_count:
                defects.append({
                    "source_ordinal": source_ordinal,
                    "final_ordinals": ordinals,
                    "reason": "limit_operator_count_mismatch",
                    "expected": expected_count,
                    "actual": actual_count,
                })
            expected_placements = [list(x) for x in target.get("placements") or []]
            actual_placements = limit_operator_placements(tex)
            if actual_placements != expected_placements:
                defects.append({
                    "source_ordinal": source_ordinal,
                    "final_ordinals": ordinals,
                    "reason": "limit_operator_placement_mismatch",
                    "expected": expected_placements,
                    "actual": actual_placements,
                })

        for target in contract["prime_targets"]:
            source_ordinal = int(target["ordinal"])
            ordinals = source_to_final.get(source_ordinal, [])
            tex = mapped_tex(ordinals, f"SRC-{source_ordinal:04d}")
            expected = list(target.get("prime_orders") or [])
            if not expected:
                defects.append({"source_ordinal": source_ordinal, "reason": "source_prime_order_unresolved"})
                continue
            issues = canonical_prime_issues(tex)
            actual = prime_orders(tex)
            if issues:
                defects.append({
                    "source_ordinal": source_ordinal,
                    "final_ordinals": ordinals,
                    "reason": "noncanonical_axmath_prime_syntax",
                    "issues": issues,
                })
            if actual != expected:
                defects.append({
                    "source_ordinal": source_ordinal,
                    "final_ordinals": ordinals,
                    "reason": "prime_order_mismatch",
                    "expected": expected,
                    "actual": actual,
                })

        for target in contract.get("structure_targets") or []:
            source_ordinal = int(target["ordinal"])
            ordinals = source_to_final.get(source_ordinal, [])
            tex = mapped_tex(ordinals, f"SRC-{source_ordinal:04d}")
            for defect in _structure_defects(
                dict(target.get("structure") or {}), tex, exact=True
            ):
                defects.append({
                    "source_ordinal": source_ordinal,
                    "final_ordinals": ordinals,
                    **defect,
                })

        for target in added_targets:
            target_id = str(target["target_id"])
            ordinals = list(target["final_ordinals"])
            expected_tex = str(target.get("expected_tex") or "")
            if smallint_issues(expected_tex):
                defects.append({"target_id": target_id, "reason": "target_uses_forbidden_smallint"})
            expected_tex = normalize_axmath_tex(expected_tex)
            actual_tex = mapped_tex(ordinals, target_id)
            expected_tokens = _semantic_tokens(expected_tex)
            actual_tokens = _semantic_tokens(actual_tex)
            if expected_tokens != actual_tokens:
                defects.append({
                    "target_id": target_id,
                    "final_ordinals": ordinals,
                    "reason": "added_target_semantic_token_mismatch",
                    "expected_tokens": expected_tokens,
                    "actual_tokens": actual_tokens,
                })
            if limit_operator_issues(actual_tex):
                defects.append({
                    "target_id": target_id,
                    "final_ordinals": ordinals,
                    "reason": "added_target_lim_placement_invalid",
                    "issues": limit_operator_issues(actual_tex),
                })
            if limit_operator_placements(expected_tex) != limit_operator_placements(actual_tex):
                defects.append({
                    "target_id": target_id,
                    "final_ordinals": ordinals,
                    "reason": "added_target_lim_placement_mismatch",
                    "expected": limit_operator_placements(expected_tex),
                    "actual": limit_operator_placements(actual_tex),
                })
            expected_primes = prime_orders(expected_tex)
            actual_primes = prime_orders(actual_tex)
            if canonical_prime_issues(actual_tex):
                defects.append({
                    "target_id": target_id,
                    "final_ordinals": ordinals,
                    "reason": "added_target_noncanonical_prime_syntax",
                    "issues": canonical_prime_issues(actual_tex),
                })
            if expected_primes != actual_primes:
                defects.append({
                    "target_id": target_id,
                    "final_ordinals": ordinals,
                    "reason": "added_target_prime_order_mismatch",
                    "expected": expected_primes,
                    "actual": actual_primes,
                })
            for defect in _structure_defects(
                _tex_structure_signature(expected_tex), actual_tex, exact=True
            ):
                defects.append({
                    "target_id": target_id,
                    "final_ordinals": ordinals,
                    **defect,
                })

    if defects and not errors:
        errors.append(f"{len(defects)} math-layout semantic defects detected")
    source_final_mapping = [
        {"source_ordinal": ordinal, "final_ordinals": ordinals}
        for ordinal, ordinals in sorted(source_to_final.items())
    ]
    return {
        "schema": SCHEMA,
        "passed": not errors and not defects,
        "mapping_pass": mapping_pass,
        "errors": errors,
        "mapping_errors": mapping_errors,
        "defects": defects,
        "source": str(source),
        "final": str(final),
        "source_sha256": contract["source_sha256"],
        "final_sha256": sha256_file(final),
        "internal_audit": str(internal_audit),
        "internal_audit_sha256": sha256_file(internal_audit),
        "formula_map": str(formula_map.resolve()) if formula_map else None,
        "formula_map_sha256": sha256_file(formula_map) if formula_map else None,
        "source_kind": contract["source_kind"],
        "source_formula_count": contract["source_formula_count"],
        "final_axmath_count": final_count,
        "mapped_source_formula_count": len(source_to_final),
        "additional_target_count": len(added_targets),
        "mapped_final_object_count": sum(len(x) for x in source_to_final.values())
            + sum(len(x["final_ordinals"]) for x in added_targets),
        "source_final_mapping": source_final_mapping,
        "limit_targets": contract["limit_target_count"],
        "limit_operators": contract["limit_operator_count"],
        "prime_targets": contract["prime_target_count"],
        "structure_targets": contract.get("structure_target_count", 0),
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

    map_template = sub.add_parser("map-template")
    map_template.add_argument("--source", required=True)
    map_template.add_argument("--out", required=True)

    validate = sub.add_parser("validate")
    validate.add_argument("--source", required=True)
    validate.add_argument("--final", required=True)
    validate.add_argument("--internal-audit", required=True)
    validate.add_argument("--formula-map", required=True)
    validate.add_argument("--source-internal-audit")
    validate.add_argument("--repair-ledger")
    validate.add_argument("--out", required=True)
    args = parser.parse_args()

    if args.mode == "map-template":
        payload = build_formula_map_template(Path(args.source), Path(args.out))
        print(json.dumps({
            "mode": args.mode,
            "source_formula_count": payload["source"]["formula_count"],
            "out": str(Path(args.out).resolve()),
        }, ensure_ascii=True, indent=2))
        return 0
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
            formula_map=Path(args.formula_map),
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
        "mapping_pass": payload.get("mapping_pass"),
        "out": str(out),
        "errors": payload.get("errors", []),
    }, ensure_ascii=True, indent=2))
    return 0 if payload.get("passed", True) else 1


if __name__ == "__main__":
    raise SystemExit(main())
