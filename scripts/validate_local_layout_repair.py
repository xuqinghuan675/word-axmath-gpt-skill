from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "o": "urn:schemas-microsoft-com:office:office",
}

W = "{%s}" % NS["w"]
M = "{%s}" % NS["m"]


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _load(docx: Path):
    with zipfile.ZipFile(docx) as z:
        return ET.fromstring(z.read("word/document.xml"))


def _paragraphs(root):
    return root.findall(".//w:p", NS)


def _axmath_count(node) -> int:
    return sum(
        1
        for ole in node.findall(".//o:OLEObject", NS)
        if (ole.attrib.get("ProgID") or "").lower() == "equation.axmath"
    )


def _soft_breaks(node) -> int:
    return len(node.findall(".//w:br", NS))


def _nonmath_text(node) -> str:
    chunks: list[str] = []

    def walk(cur, in_math=False, in_object=False):
        if cur.tag in (M + "oMath", M + "oMathPara"):
            in_math = True
        if cur.tag == W + "object":
            in_object = True
        if cur.tag == W + "t" and not in_math and not in_object:
            chunks.append(cur.text or "")
        for child in list(cur):
            walk(child, in_math, in_object)

    walk(node)
    return "".join(chunks)


def _paragraph_semantic_signature(node) -> dict:
    return {
        "nonmath_text": _nonmath_text(node),
        "axmath": _axmath_count(node),
        "omath": len(node.findall(".//m:oMath", NS)),
        "soft_breaks": _soft_breaks(node),
    }


def validate(source: Path, baseline: Path, candidate: Path, repair_map: Path) -> dict:
    source = source.resolve()
    baseline = baseline.resolve()
    candidate = candidate.resolve()
    repair_map = repair_map.resolve()
    mapping = json.loads(repair_map.read_text(encoding="utf-8-sig"))

    if mapping.get("repair_class") != "M2_SOURCE_VISUAL_LINE_SPLIT":
        raise ValueError("repair map is not M2_SOURCE_VISUAL_LINE_SPLIT")

    source_sha = sha256_file(source)
    baseline_sha = sha256_file(baseline)
    candidate_sha = sha256_file(candidate)

    sr = _load(source)
    br = _load(baseline)
    cr = _load(candidate)
    bp = _paragraphs(br)
    cp = _paragraphs(cr)

    errors: list[str] = []
    if source_sha.lower() != str(mapping.get("source_sha256") or "").lower():
        errors.append("source SHA mismatch")
    if baseline_sha.lower() != str(mapping.get("working_sha256") or "").lower():
        errors.append("baseline SHA mismatch")
    if len(bp) != len(cp):
        errors.append(f"paragraph count changed: {len(bp)} -> {len(cp)}")

    source_omath = len(sr.findall(".//m:oMath", NS))
    baseline_omath = len(br.findall(".//m:oMath", NS))
    candidate_omath = len(cr.findall(".//m:oMath", NS))
    baseline_axmath = _axmath_count(br)
    candidate_axmath = _axmath_count(cr)

    if baseline_omath != 0 or candidate_omath != 0:
        errors.append(
            f"OfficeMath contract failed: baseline={baseline_omath}, candidate={candidate_omath}"
        )
    expected_final = int(mapping["expected_final_axmath_count"])
    if candidate_axmath != expected_final:
        errors.append(
            f"AxMath count mismatch: candidate={candidate_axmath}, expected={expected_final}"
        )

    first_repair = int(mapping["first_repair_paragraph"])
    prefix_ok = True
    prefix_diffs = []
    for i in range(1, min(first_repair, len(bp) + 1)):
        bs = _paragraph_semantic_signature(bp[i - 1])
        cs = _paragraph_semantic_signature(cp[i - 1])
        if bs != cs:
            prefix_ok = False
            prefix_diffs.append({"paragraph": i, "baseline": bs, "candidate": cs})
            if len(prefix_diffs) >= 20:
                break
    if not prefix_ok:
        errors.append("frozen prefix changed before first repair paragraph")

    baseline_flat = "".join(_nonmath_text(p) for p in bp)
    candidate_flat = "".join(_nonmath_text(p) for p in cp)
    nonmath_flat_equal = baseline_flat == candidate_flat
    if not nonmath_flat_equal:
        errors.append("non-math text changed across local repair")

    target_rows = []
    for row in mapping.get("rows") or []:
        pi = int(row["working_paragraph_index"])
        if pi < 1 or pi > len(cp):
            errors.append(f"target paragraph out of range: {pi}")
            continue
        para = cp[pi - 1]
        ax = _axmath_count(para)
        brks = _soft_breaks(para)
        expected_ax = int(row["expected_after_axmath"])
        expected_brks = int(row["expected_soft_breaks_added"])
        rec = {
            "paragraph": pi,
            "source_paragraph": int(row["source_paragraph_index"]),
            "source_ordinal": int(row["source_ordinal"]),
            "expected_axmath": expected_ax,
            "candidate_axmath": ax,
            "expected_soft_breaks": expected_brks,
            "candidate_soft_breaks": brks,
            "passed": ax == expected_ax and brks >= expected_brks,
        }
        target_rows.append(rec)
        if not rec["passed"]:
            errors.append(
                f"paragraph {pi} target contract failed: AxMath {ax}/{expected_ax}, "
                f"soft breaks {brks}/{expected_brks}"
            )

    return {
        "schema": "axmath-local-layout-repair-ledger/v1",
        "success": not errors,
        "source": str(source),
        "source_sha256": source_sha,
        "baseline": str(baseline),
        "baseline_sha256": baseline_sha,
        "candidate": str(candidate),
        "candidate_sha256": candidate_sha,
        "repair_map": str(repair_map),
        "repair_map_sha256": sha256_file(repair_map),
        "repair_class": mapping["repair_class"],
        "first_repair_paragraph": first_repair,
        "prefix_semantic_unchanged": prefix_ok,
        "prefix_diffs": prefix_diffs,
        "nonmath_text_flat_equal": nonmath_flat_equal,
        "baseline_paragraphs": len(bp),
        "candidate_paragraphs": len(cp),
        "source_omath": source_omath,
        "baseline_omath": baseline_omath,
        "candidate_omath": candidate_omath,
        "baseline_axmath": baseline_axmath,
        "candidate_axmath": candidate_axmath,
        "expected_candidate_axmath": expected_final,
        "expected_axmath_increase": int(mapping["expected_axmath_increase"]),
        "targets": target_rows,
        "errors": errors,
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--baseline", required=True)
    ap.add_argument("--candidate", required=True)
    ap.add_argument("--repair-map", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    report = validate(
        Path(args.source),
        Path(args.baseline),
        Path(args.candidate),
        Path(args.repair_map),
    )
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "success": report["success"],
                "prefix_semantic_unchanged": report["prefix_semantic_unchanged"],
                "nonmath_text_flat_equal": report["nonmath_text_flat_equal"],
                "candidate_axmath": report["candidate_axmath"],
                "expected_candidate_axmath": report["expected_candidate_axmath"],
                "target_count": len(report["targets"]),
                "errors": report["errors"],
                "out": str(out.resolve()),
            },
            ensure_ascii=True,
            indent=2,
        )
    )
    return 0 if report["success"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
