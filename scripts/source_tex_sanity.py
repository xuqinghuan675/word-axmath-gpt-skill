from __future__ import annotations

import re
import unicodedata


def strip_math_delimiters(value: str) -> str:
    s = (value or "").strip()
    while len(s) >= 2 and s.startswith("$") and s.endswith("$"):
        s = s[1:-1].strip()
    return s


def word_latex_is_sane(value: str) -> tuple[bool, list[str]]:
    """Conservative gate for Word linear-math/LaTeX export.

    Word can report a successful representation switch while returning a
    non-empty string that is not safe input for AxMath. Observed production
    artifacts include controls, mathematical-alphanumeric Unicode glyphs,
    placeholders, and malformed begin constructs. Reject those and fall back
    to exact frozen OMML.
    """
    s = value or ""
    reasons: list[str] = []
    if not s.strip():
        reasons.append("empty")
    if re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", s):
        reasons.append("control_character")
    if any(ch in s for ch in ("\r", "\n", "\t")):
        reasons.append("embedded_line_control")
    if any(0x1D400 <= ord(ch) <= 0x1D7FF for ch in s):
        reasons.append("unicode_math_alphanumeric")
    for bad in ("▒", "〖", "〗", "\\begin("):
        if bad in s:
            reasons.append("word_linear_artifact:" + bad)

    depth = 0
    for ch in s:
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth < 0:
                reasons.append("unbalanced_brace")
                break
    if depth != 0 and "unbalanced_brace" not in reasons:
        reasons.append("unbalanced_brace")
    return not reasons, reasons


def normalize_word_latex(value: str) -> str:
    s = unicodedata.normalize("NFKC", value or "")
    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return strip_math_delimiters(s)
