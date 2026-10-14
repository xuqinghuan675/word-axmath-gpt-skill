"""Hash-bound review of selected Class E AxMath internal exports.

The frozen OMML remains the semantic source. A review records the exact source
formula, exported AxMath TeX, and a human conclusion for the current DOCX hash.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from build_source_tex_map import _safe_source_tex, static_axmath_inventory
from internal_content_audit import CONTROL, sha256_file

SCHEMA = "axmath-class-e-review/v1"
MATH_NS = {"m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}


def _source_nodes(source: Path) -> list[ET.Element]:
    with zipfile.ZipFile(source) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    return root.findall(".//m:oMath", MATH_NS)


def _inspection_rows(inspection: Path, working: Path) -> dict[int, dict]:
    data = json.loads(inspection.read_text(encoding="utf-8-sig"))
    actual_sha = sha256_file(working)
    if (
        data.get("schema") != "axmath-internal-tex-diagnostic/v1"
        or Path(str(data.get("input") or "")).resolve() != working.resolve()
        or data.get("input_sha256_before") != actual_sha
        or data.get("input_sha256_after") != actual_sha
        or data.get("input_unchanged") is not True
        or data.get("success") is not True
        or data.get("failed")
        or data.get("diagnostic_only") is not True
    ):
        raise ValueError("Class E inspection is failed, stale, or not bound to this working DOCX")
    rows = data.get("exports") or []
    by_ordinal = {int(row["ordinal"]): row for row in rows}
    if len(by_ordinal) != len(rows):
        raise ValueError("Class E inspection contains duplicate ordinals")
    return by_ordinal


def _contract_rows(source: Path, working: Path, inspection: Path, ordinals: list[int]) -> list[dict]:
    nodes = _source_nodes(source)
    actual_count = static_axmath_inventory(working)["axmath_count"]
    if len(nodes) != actual_count:
        raise ValueError("Ordinal Class E review requires one AxMath per frozen source formula")
    exports = _inspection_rows(inspection, working)
    rows: list[dict] = []
    for ordinal in ordinals:
        if ordinal < 1 or ordinal > len(nodes) or ordinal not in exports:
            raise ValueError(f"Class E inspection/source has no ordinal {ordinal}")
        node = nodes[ordinal - 1]
        xml = ET.tostring(node, encoding="utf-8")
        source_text = "".join((part.text or "") for part in node.findall(".//m:t", MATH_NS))
        try:
            source_tex = "$" + _safe_source_tex(node) + "$"
            source_tex_error = None
        except Exception as exc:
            source_tex = None
            source_tex_error = str(exc)
        exported = exports[ordinal].get("axmath_exported_tex")
        if not isinstance(exported, str) or exports[ordinal].get("success") is not True:
            raise ValueError(f"Class E inspection failed to export ordinal {ordinal}")
        rows.append({
            "ordinal": ordinal,
            "source_omml_sha256": hashlib.sha256(xml).hexdigest(),
            "source_omml_xml": xml.decode("utf-8"),
            "source_omml_text": source_text,
            "source_tex": source_tex,
            "source_tex_error": source_tex_error,
            "axmath_exported_tex": exported,
            "export_empty": not exported.strip(),
            "export_contains_control": bool(CONTROL.search(exported)),
            "decision": None,
            "evidence": "",
        })
    return rows


def prepare(source: Path, working: Path, inspection: Path, ordinals: list[int], out: Path) -> dict:
    source, working, inspection, out = map(Path, (source, working, inspection, out))
    source, working, inspection, out = (p.resolve() for p in (source, working, inspection, out))
    if out.exists():
        raise FileExistsError(out)
    selected = sorted(set(ordinals))
    if not selected or len(selected) > 24:
        raise ValueError("Class E review requires 1..24 selected ordinals")
    rows = _contract_rows(source, working, inspection, selected)
    payload = {
        "schema": SCHEMA,
        "source": str(source),
        "source_sha256": sha256_file(source),
        "working": str(working),
        "working_sha256": sha256_file(working),
        "inspection": str(inspection),
        "inspection_sha256": sha256_file(inspection),
        "rows": rows,
    }
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return payload


def validate(review: Path, source: Path, working: Path) -> dict:
    review, source, working = (Path(p).resolve() for p in (review, source, working))
    data = json.loads(review.read_text(encoding="utf-8-sig"))
    errors: list[str] = []
    if data.get("schema") != SCHEMA:
        errors.append("unsupported Class E review schema")
    if Path(str(data.get("source") or "")).resolve() != source or data.get("source_sha256") != sha256_file(source):
        errors.append("Class E review frozen source path/hash is stale")
    if Path(str(data.get("working") or "")).resolve() != working or data.get("working_sha256") != sha256_file(working):
        errors.append("Class E review working path/hash is stale")
    inspection = Path(str(data.get("inspection") or "")).resolve()
    if not inspection.is_file() or data.get("inspection_sha256") != sha256_file(inspection):
        errors.append("Class E inspection file/hash is stale")
    if errors:
        return {"valid": False, "errors": errors, "verified_ordinals": [],
                "rebuild_ordinals": [], "review": str(review)}
    rows = data.get("rows") or []
    try:
        ordinals = [int(row["ordinal"]) for row in rows]
        if not ordinals or len(ordinals) != len(set(ordinals)):
            raise ValueError("empty or duplicate ordinals")
        expected = _contract_rows(source, working, inspection, sorted(ordinals))
    except (KeyError, TypeError, ValueError) as exc:
        errors.append(f"cannot verify Class E formula contracts: {exc}")
        expected = []
    verified, rebuild = [], []
    identity_fields = (
        "ordinal", "source_omml_sha256", "source_omml_xml", "source_omml_text", "source_tex",
        "source_tex_error", "axmath_exported_tex", "export_empty", "export_contains_control",
    )
    for row, original in zip(sorted(rows, key=lambda r: int(r["ordinal"])), expected):
        ordinal = original["ordinal"]
        if any(row.get(field) != original.get(field) for field in identity_fields):
            errors.append(f"Class E formula/export evidence changed for ordinal {ordinal}")
            continue
        decision = row.get("decision")
        if decision not in ("verified_match", "needs_rebuild") or not str(row.get("evidence") or "").strip():
            errors.append(f"Class E ordinal {ordinal} lacks an evidenced conclusion")
        elif decision == "verified_match":
            if original["export_empty"] or original["export_contains_control"]:
                errors.append(f"Class E ordinal {ordinal} cannot approve empty/control internal content")
            else:
                verified.append(ordinal)
        else:
            rebuild.append(ordinal)
    return {"valid": not errors, "errors": errors, "verified_ordinals": verified, "rebuild_ordinals": rebuild, "review": str(review)}


def main() -> int:
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="command", required=True)
    create = sub.add_parser("prepare")
    create.add_argument("--source", required=True)
    create.add_argument("--working", required=True)
    create.add_argument("--inspection", required=True)
    create.add_argument("--ordinals", required=True)
    create.add_argument("--out", required=True)
    check = sub.add_parser("validate")
    check.add_argument("--source", required=True)
    check.add_argument("--working", required=True)
    check.add_argument("--review", required=True)
    args = parser.parse_args()
    if args.command == "prepare":
        result = prepare(Path(args.source), Path(args.working), Path(args.inspection),
                         [int(x.strip()) for x in args.ordinals.split(",") if x.strip()], Path(args.out))
        print(json.dumps({"review": args.out, "rows": len(result["rows"])}, ensure_ascii=False))
        return 0
    result = validate(Path(args.review), Path(args.source), Path(args.working))
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["valid"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
