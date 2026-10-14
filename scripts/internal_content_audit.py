"""Validate a complete, read-only AxMath internal TeX export.

Object count and page rendering cannot detect an AxMath OLE whose editable
content is empty. The PowerShell inspector performs the actual AxMath macro
exports; this module checks their coverage and binds the result to the DOCX.
"""
from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path

CONTROL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
SCHEMA = "axmath-internal-tex-diagnostic/v1"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def validate_full_internal_audit(
    audit_path: Path, final_docx: Path, expected_count: int
) -> dict:
    """Fail closed unless every AxMath ordinal has a valid internal export."""
    audit_path = Path(audit_path).resolve()
    final_docx = Path(final_docx).resolve()
    data = json.loads(audit_path.read_text(encoding="utf-8-sig"))
    actual_sha = sha256_file(final_docx)
    errors: list[str] = []
    if data.get("schema") != SCHEMA:
        errors.append("unsupported internal audit schema")
    if data.get("target_scope") != "all":
        errors.append("internal audit did not cover all AxMath objects")
    if Path(str(data.get("input") or "")).resolve() != final_docx:
        errors.append("internal audit input is not the final DOCX")
    if data.get("input_sha256_before") != actual_sha or data.get("input_sha256_after") != actual_sha:
        errors.append("internal audit DOCX hash is stale or changed during inspection")
    if data.get("input_unchanged") is not True or data.get("diagnostic_only") is not True:
        errors.append("internal audit does not prove read-only inspection")
    if data.get("success") is not True:
        errors.append("internal audit did not complete successfully")
    if data.get("failed"):
        errors.append("internal audit contains failed exports")
    try:
        audited_count = int(data.get("axmath_count"))
    except (TypeError, ValueError):
        audited_count = -1
    if audited_count != expected_count:
        errors.append(f"internal audit count {audited_count} != final AxMath count {expected_count}")

    exports = data.get("exports")
    if not isinstance(exports, list):
        exports = []
    if len(exports) != expected_count:
        errors.append(f"internal audit exported {len(exports)} of {expected_count} AxMath objects")
    ordinals: list[int] = []
    defects: list[dict] = []
    for row in exports:
        try:
            ordinal = int(row.get("ordinal"))
        except (AttributeError, TypeError, ValueError):
            defects.append({"ordinal": None, "reason": "invalid_export_row"})
            continue
        ordinals.append(ordinal)
        tex = row.get("axmath_exported_tex")
        if row.get("success") is not True or not isinstance(tex, str):
            defects.append({"ordinal": ordinal, "reason": "export_failed"})
        elif not tex.strip():
            defects.append({"ordinal": ordinal, "reason": "empty_internal_export"})
        elif CONTROL.search(tex):
            defects.append({"ordinal": ordinal, "reason": "control_character_in_internal_export"})
    if sorted(ordinals) != list(range(1, expected_count + 1)):
        errors.append("internal audit has missing, duplicate, or out-of-range ordinals")
    if defects:
        errors.append(f"{len(defects)} internal AxMath exports are empty, corrupt, or failed")
    if data.get("integrity_pass") is not True:
        errors.append("internal audit did not mark internal integrity passed")
    return {
        "passed": not errors,
        "errors": errors,
        "defects": defects,
        "source": str(audit_path),
        "audit_sha256": sha256_file(audit_path),
        "final_sha256": actual_sha,
        "expected_count": expected_count,
        "audited_count": audited_count,
    }
