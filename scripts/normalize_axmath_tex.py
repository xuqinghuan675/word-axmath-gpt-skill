from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

# AxMath 2.7.0.58 prime contract, verified on a real Word + AxMath installation:
#   first derivative  -> \prime
#   second derivative -> ''
#   third derivative  -> '''
# AxMath's own executable contains these parser tokens, and AMSAM2TeX emits the
# same forms after AMSTeX2AM creates an editable Equation.AxMath object.
#
# Raw U+2034 TRIPLE PRIME is NOT a safe AMSTeX2AM input on that version: it
# round-trips as '?'.  Therefore Unicode/typographic prime glyphs are source
# evidence only; approved donor TeX must use the verified AxMath parser tokens.
PRIME_SOURCE_MARKERS = ("'", "′", "″", "‴", "⁗", "ʹ", "ʺ", "’", "‘")
NONCANONICAL_PRIME_LITERALS = ("′", "″", "‴", "⁗", "ʹ", "ʺ", "’", "‘")

_SUPERSCRIPT_MAP = {
    "⁰": "0",
    "¹": "1",
    "²": "2",
    "³": "3",
    "⁴": "4",
    "⁵": "5",
    "⁶": "6",
    "⁷": "7",
    "⁸": "8",
    "⁹": "9",
    "⁺": "+",
    "⁻": "-",
}
_SUPERSCRIPT_RUN_RE = re.compile("[" + re.escape("".join(_SUPERSCRIPT_MAP)) + "]+")

# Canonical AxMath prime syntax after normalization.  Match the longest
# apostrophe token first so triple/double prime stay atomic for our contract.
_CANONICAL_PRIME_TOKEN_RE = re.compile(r"'''|''|\\prime\b")
_SINGLE_ASCII_APOSTROPHE_RE = re.compile(r"(?<!')'(?!')")
_UNVERIFIED_APOSTROPHE_RUN_RE = re.compile(r"'{4,}")
_UNGROUPED_PRIME_EXP_RE = re.compile(
    r"(?:\\prime\b|'{2,3})\s*\^(?:\{|[A-Za-z0-9])"
)

# omml2latex can encode a primed atom followed by another exponent as one
# superscript run, e.g. f′² -> {f}^{\prime2}. Feeding that directly to AxMath
# is ambiguous because \prime2 is not the verified first-prime token.
# Recover the semantic structure before general primed-atom grouping.
_OMML_FUSED_PRIME_EXP_RE = re.compile(
    r"(?P<atom>"
    r"\{[^{}]+\}"
    r"|\\[A-Za-z]+"
    r"|[A-Za-z0-9]"
    r")"
    r"\^\{(?P<prime>\\prime|'{2,3})\s*(?P<exp>[A-Za-z0-9+\-]+)\}"
)
_FUSED_PRIME_TOKEN_RE = re.compile(r"\\prime(?=[A-Za-z0-9])")

def _read_tex_argument(value: str, position: int) -> tuple[str, str, int, int, int]:
    """Read one TeX argument: content, raw text, start, content start, end."""
    while position < len(value) and value[position].isspace():
        position += 1
    start = position
    if position >= len(value):
        return "", "", start, start, start
    if value[position] == "{":
        content_start = position + 1
        depth = 1
        position += 1
        while position < len(value) and depth:
            if value[position] == "\\":
                position += 2
                continue
            if value[position] == "{":
                depth += 1
            elif value[position] == "}":
                depth -= 1
            position += 1
        content_end = position - 1 if depth == 0 else position
        return value[content_start:content_end], value[start:position], start, content_start, position
    if value[position] == "\\":
        match = re.match(r"\\[A-Za-z]+|\\.", value[position:])
        end = position + (len(match.group(0)) if match else 1)
        return value[position:end], value[position:end], start, position, end
    end = position + 1
    return value[position:end], value[position:end], start, position, end


def _underset_lim_wrappers(value: str) -> dict[int, dict]:
    wrappers: dict[int, dict] = {}
    for match in re.finditer(r"\\underset(?![A-Za-z])", value):
        lower, _, _, _, after_lower = _read_tex_argument(value, match.end())
        base, _, _, base_content_start, after_base = _read_tex_argument(value, after_lower)
        if base.strip() != r"\lim":
            continue
        lim_offset = base.find(r"\lim")
        wrappers[base_content_start + lim_offset] = {
            "lower": lower.strip(),
            "end": after_base,
        }
    return wrappers


def _parse_lim_scripts(value: str, position: int) -> tuple[list[dict], int, bool]:
    while position < len(value) and value[position].isspace():
        position += 1
    placement = re.match(r"\\(?:no)?limits(?![A-Za-z])", value[position:])
    had_placement = placement is not None
    if placement:
        position += len(placement.group(0))
    scripts: list[dict] = []
    while True:
        while position < len(value) and value[position].isspace():
            position += 1
        if position >= len(value) or value[position] not in "_^":
            break
        marker = value[position]
        content, raw, start, _, end = _read_tex_argument(value, position + 1)
        if end <= position + 1:
            break
        scripts.append({
            "marker": marker,
            "content": content.strip(),
            "raw": raw,
            "start": start,
            "end": end,
        })
        position = end
    return scripts, position, had_placement


def normalize_limit_operator_placement(tex: str) -> str:
    """Encode a lower lim script with underset, which survives AxMath export."""
    value = str(tex)
    wrappers = _underset_lim_wrappers(value)
    matches = list(re.finditer(r"\\lim(?![A-Za-z])", value))
    for match in reversed(matches):
        if match.start() in wrappers:
            continue
        scripts, end, had_placement = _parse_lim_scripts(value, match.end())
        lower = next((x for x in scripts if x["marker"] == "_"), None)
        upper = next((x for x in scripts if x["marker"] == "^"), None)
        if lower is None and upper is None and not had_placement:
            continue
        if lower is not None:
            replacement = r"\underset{" + lower["content"] + r"}{\lim}"
            if upper is not None:
                replacement += "^{" + upper["content"] + "}"
        elif upper is not None:
            replacement = r"\lim^{" + upper["content"] + "}"
        else:
            replacement = r"\lim"
        value = value[:match.start()] + replacement + value[end:]
    return value


def limit_operator_issues(tex: str) -> list[str]:
    """Require lower lim scripts to use the AxMath-stable underset form."""
    value = str(tex)
    wrappers = _underset_lim_wrappers(value)
    issues: list[str] = []
    for match in re.finditer(r"\\lim(?![A-Za-z])", value):
        wrapper = wrappers.get(match.start())
        if wrapper is not None:
            if not wrapper["lower"]:
                issues.append("empty_underset_lim")
            scripts, _, had_placement = _parse_lim_scripts(value, wrapper["end"])
            if had_placement:
                issues.append("lim_limits_command_not_allowed")
            if any(x["marker"] == "_" for x in scripts):
                issues.append("duplicate_lim_subscript")
            continue
        scripts, _, had_placement = _parse_lim_scripts(value, match.end())
        if had_placement:
            issues.append("lim_limits_command_not_allowed")
        if any(x["marker"] == "_" for x in scripts):
            issues.append("lim_subscript_requires_underset")
    return sorted(set(issues))


def limit_operator_count(tex: str) -> int:
    """Count lim operators carrying a lower or upper script."""
    return len(limit_operator_placements(tex))


def limit_operator_placements(tex: str) -> list[list[str]]:
    """Return each scripted lim operator's ordered lower/upper directions."""
    value = str(tex)
    wrappers = _underset_lim_wrappers(value)
    placements: list[list[str]] = []
    for match in re.finditer(r"\\lim(?![A-Za-z])", value):
        wrapper = wrappers.get(match.start())
        if wrapper is not None:
            current = ["below"] if wrapper["lower"] else []
            scripts, _, _ = _parse_lim_scripts(value, wrapper["end"])
        else:
            current = []
            scripts, _, _ = _parse_lim_scripts(value, match.end())
        current.extend("below" if x["marker"] == "_" else "above" for x in scripts)
        if current:
            placements.append(current)
    return placements


def smallint_issues(tex: str) -> list[str]:
    """The AxMath output contract uses \\int; \\smallint is forbidden."""
    return ["smallint_forbidden"] if re.search(r"\\smallint(?![A-Za-z])", str(tex)) else []



def _normalize_omml_fused_prime_exponent(tex: str) -> str:
    def repl(match: re.Match[str]) -> str:
        atom = match.group("atom")
        if atom.startswith("{") and atom.endswith("}"):
            atom = atom[1:-1]
        return "{" + atom + match.group("prime") + "}^{"+ match.group("exp") + "}"

    previous = None
    current = str(tex)
    while previous != current:
        previous = current
        current = _OMML_FUSED_PRIME_EXP_RE.sub(repl, current)
    return current


_PRIMED_ATOM_EXP_RE = re.compile(
    r"(?P<atom>"
    r"(?:\\[A-Za-z]+|[A-Za-z0-9]|\([^()]+\))"
    r"(?:_(?:\{[^{}]+\}|\\[A-Za-z]+|[A-Za-z0-9]))?"
    r")"
    r"(?P<prime>\\prime\b|'{2,3})\s*"
    r"\^(?:\{(?P<braced>[^{}]+)\}|(?P<plain>[A-Za-z0-9]))"
)


def contains_prime_or_derivative_marker(text: str) -> bool:
    value = str(text or "")
    return any(marker in value for marker in PRIME_SOURCE_MARKERS) or bool(
        re.search(r"\\prime\b", value)
    )


def _replace_superscript_run(match: re.Match[str]) -> str:
    exponent = "".join(_SUPERSCRIPT_MAP[ch] for ch in match.group(0))
    return "^{" + exponent + "}"


def _replace_single_prime_like_runs(tex: str) -> str:
    # A run of source characters that each mean one prime is converted by
    # derivative order, not by visual glyph substitution.
    prime_like = "′ʹ’‘"
    pattern = re.compile("[" + re.escape(prime_like) + "]{1,3}")

    def repl(match: re.Match[str]) -> str:
        order = len(match.group(0))
        if order == 1:
            return r"\prime"
        if order == 2:
            return "''"
        return "'''"

    return pattern.sub(repl, tex)


def _canonicalize_prime_tokens(tex: str) -> str:
    value = str(tex)

    # Dedicated source glyphs first.
    value = value.replace("‴", "'''")
    value = value.replace("″", "''")
    value = value.replace("ʺ", "''")
    value = _replace_single_prime_like_runs(value)

    # Existing AxMath/TeX \prime is already the verified first-prime token.
    # A lone ASCII apostrophe in math is normalized to the same canonical token;
    # verified double/triple apostrophe tokens remain unchanged.
    value = _SINGLE_ASCII_APOSTROPHE_RE.sub(r"\\prime", value)
    return value


def _group_primed_atom_before_exponent(tex: str) -> str:
    # Prime itself is superscript syntax.  A second exponent must apply to the
    # already-primed atom, e.g. y′² -> {y\prime}^{2}, not y\prime^{2}.
    previous = None
    current = tex
    while previous != current:
        previous = current

        def repl(match: re.Match[str]) -> str:
            exponent = match.group("braced")
            if exponent is None:
                exponent = match.group("plain")
            return (
                "{"
                + match.group("atom")
                + match.group("prime")
                + "}^{"
                + exponent
                + "}"
            )

        current = _PRIMED_ATOM_EXP_RE.sub(repl, current)
    return current


def normalize_axmath_tex(tex: str) -> str:
    value = normalize_limit_operator_placement(str(tex))
    value = re.sub(r"\\smallint(?![A-Za-z])", r"\\int", value)
    value = _canonicalize_prime_tokens(value)

    # U+2057 QUADRUPLE PRIME is intentionally not auto-normalized.  The
    # installed AxMath binary exposes verified parser tokens only for orders
    # 1..3, so higher orders require explicit review rather than guessing.
    value = _SUPERSCRIPT_RUN_RE.sub(_replace_superscript_run, value)
    value = _normalize_omml_fused_prime_exponent(value)
    value = _group_primed_atom_before_exponent(value)
    return value


def prime_orders(tex: str) -> list[int]:
    """Return the verified AxMath prime orders encoded in canonical TeX."""
    orders: list[int] = []
    for match in _CANONICAL_PRIME_TOKEN_RE.finditer(str(tex)):
        token = match.group(0)
        if token.startswith("\\prime"):
            orders.append(1)
        else:
            orders.append(len(token))
    return orders


def canonical_prime_issues(tex: str) -> list[str]:
    value = str(tex)
    issues: list[str] = []

    if any(ch in value for ch in NONCANONICAL_PRIME_LITERALS):
        issues.append("noncanonical_prime_literal")
    if _SUPERSCRIPT_RUN_RE.search(value):
        issues.append("unicode_superscript_literal")
    if _UNVERIFIED_APOSTROPHE_RUN_RE.search(value):
        issues.append("unverified_prime_order")
    if _SINGLE_ASCII_APOSTROPHE_RE.search(value):
        issues.append("noncanonical_single_prime_apostrophe")
    if _FUSED_PRIME_TOKEN_RE.search(value):
        issues.append("fused_prime_exponent")
    if _UNGROUPED_PRIME_EXP_RE.search(value):
        issues.append("ungrouped_primed_atom_exponent")
    return issues


def _ensure_single_dollar(tex: str) -> str:
    value = str(tex).strip()
    if value.startswith("$$") and value.endswith("$$") and len(value) >= 4:
        value = "$" + value[2:-2] + "$"
    elif not (value.startswith("$") and value.endswith("$")):
        value = "$" + value.strip("$") + "$"
    return value


def normalize_export_json(input_path: Path) -> list[dict]:
    payload = json.loads(input_path.read_text(encoding="utf-8-sig"))
    if isinstance(payload, dict):
        rows = payload.get("exports", [])
    elif isinstance(payload, list):
        rows = payload
    else:
        rows = []

    out: list[dict] = []
    for row in rows:
        if not isinstance(row, dict) or not row.get("success", True):
            continue
        ordinal = int(row["ordinal"])
        raw = str(row.get("word_latex") or row.get("tex") or "")
        wrapped = _ensure_single_dollar(raw)
        normalized = normalize_axmath_tex(wrapped)
        out.append(
            {
                "ordinal": ordinal,
                "tex": normalized,
                "source_word_latex": raw,
                "normalization_changed": normalized != wrapped,
                "prime_orders": prime_orders(normalized),
                "prime_contract": "axmath_builtin_prime_v2",
                "limit_operator_count": limit_operator_count(normalized),
                "limit_operator_issues": limit_operator_issues(normalized),
            }
        )
    return out


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Normalize Word/AxMath TeX to AxMath 2.7.0.58 verified built-in "
            "prime syntax: \\prime, '', '''."
        )
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--tex")
    group.add_argument("--export-json")
    parser.add_argument("--out")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()

    if args.tex is not None:
        normalized = normalize_axmath_tex(args.tex)
        print(normalized)
        if args.check:
            return 0 if not canonical_prime_issues(normalized) else 1
        return 0

    rows = normalize_export_json(Path(args.export_json).resolve())
    encoded = json.dumps(rows, ensure_ascii=False, indent=2)
    if args.out:
        Path(args.out).resolve().write_text(encoded, encoding="utf-8")
    else:
        print(encoded)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
