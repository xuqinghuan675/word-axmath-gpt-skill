from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
}
W = "{%s}" % NS["w"]


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _replace_style_dimension(style: str, name: str, value_pt: float):
    replacement = f"{name}:{value_pt:.3f}".rstrip("0").rstrip(".") + "pt"
    pattern = rf"(?:^|;){re.escape(name)}:[0-9.]+pt(?=;|$)"
    match = re.search(pattern, style or "")
    if match:
        prefix = ";" if match.group(0).startswith(";") else ""
        return style[: match.start()] + prefix + replacement + style[match.end() :]
    return (style or "").rstrip(";") + ";" + replacement


def _ensure_rpr(run):
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        rpr = etree.Element(W + "rPr")
        run.insert(0, rpr)
    return rpr


def _validate_plan_binding(input_docx: Path, plan_report: dict) -> None:
    expected_path = plan_report.get("working_docx")
    if expected_path and Path(expected_path).resolve() != input_docx.resolve():
        raise RuntimeError("Calibration plan was generated for a different working DOCX.")
    expected_hash = plan_report.get("working_sha256")
    if expected_hash and _sha256_file(input_docx) != expected_hash:
        raise RuntimeError("Calibration plan is stale: working DOCX hash no longer matches.")


def _parse_ordinals(raw: str | None) -> set[int]:
    if not raw:
        return set()
    out: set[int] = set()
    for token in raw.split(","):
        token = token.strip()
        if token:
            out.add(int(token))
    return out


def apply_plan(
    input_docx: Path,
    output_docx: Path,
    plan_report: dict,
    selected_ordinals: set[int] | None = None,
    *,
    allow_in_place: bool = False,
    overwrite: bool = False,
):
    input_docx = input_docx.resolve()
    output_docx = output_docx.resolve()
    if not input_docx.is_file():
        raise FileNotFoundError(input_docx)
    if input_docx == output_docx and not allow_in_place:
        raise RuntimeError("Refusing in-place OLE calibration; use a new output DOCX.")
    if output_docx.exists() and input_docx != output_docx and not overwrite:
        raise FileExistsError(f"Output already exists: {output_docx}")

    _validate_plan_binding(input_docx, plan_report)
    selected_ordinals = set(selected_ordinals or [])

    rows = plan_report.get("calibration_plan", [])
    plan: dict[int, dict] = {}
    for row in rows:
        ordinal = int(row["ordinal"])
        if selected_ordinals:
            if ordinal not in selected_ordinals:
                continue
        elif not row.get("auto_apply"):
            continue
        plan[ordinal] = row

    if selected_ordinals:
        missing = sorted(selected_ordinals - set(plan))
        if missing:
            raise RuntimeError(f"Selected ordinals are absent from calibration plan: {missing}")

    if not plan:
        raise RuntimeError(
            "No calibration rows selected. Geometry recommendations require explicit visual confirmation; "
            "pass --ordinals after reviewing the affected formulas."
        )

    with zipfile.ZipFile(input_docx, "r") as z:
        files = {info.filename: z.read(info.filename) for info in z.infolist()}
    root = etree.fromstring(files["word/document.xml"])

    applied = []
    ordinal = 0
    for obj in root.xpath(".//w:object", namespaces=NS):
        ole = obj.find(".//{%s}OLEObject" % NS["o"])
        if ole is None or (ole.get("ProgID") or "").lower() != "equation.axmath":
            continue
        ordinal += 1
        row = plan.get(ordinal)
        if not row:
            continue

        shape = obj.find(".//{%s}shape" % NS["v"])
        if shape is None:
            raise RuntimeError(f"AxMath ordinal {ordinal} has no v:shape")

        run = obj.getparent()
        while run is not None and etree.QName(run).localname != "r":
            run = run.getparent()
        if run is None:
            raise RuntimeError(f"AxMath ordinal {ordinal} has no parent run")

        target_width = float(row["target_width_pt"])
        target_height = float(row["target_height_pt"])
        if target_width <= 0 or target_height <= 0:
            raise RuntimeError(f"AxMath ordinal {ordinal} has invalid target dimensions")

        before = {
            "style": shape.get("style"),
            "dxaOrig": obj.get(W + "dxaOrig"),
            "dyaOrig": obj.get(W + "dyaOrig"),
            "position": None,
        }
        rpr = run.find("w:rPr", NS)
        pos = rpr.find("w:position", NS) if rpr is not None else None
        if pos is not None:
            before["position"] = pos.get(W + "val")

        style = shape.get("style") or ""
        style = _replace_style_dimension(style, "width", target_width)
        style = _replace_style_dimension(style, "height", target_height)
        shape.set("style", style)
        obj.set(W + "dxaOrig", str(int(row["target_dxa_orig"])))
        obj.set(W + "dyaOrig", str(int(row["target_dya_orig"])))

        target_pos = row.get("target_position_half_points")
        if target_pos is not None:
            rpr = _ensure_rpr(run)
            pos = rpr.find("w:position", NS)
            if pos is None:
                pos = etree.Element(W + "position")
                rpr.append(pos)
            pos.set(W + "val", str(int(target_pos)))

        applied.append({
            "ordinal": ordinal,
            "reason": row.get("reason", []),
            "scale": row.get("scale"),
            "before": before,
            "after": {
                "style": shape.get("style"),
                "dxaOrig": obj.get(W + "dxaOrig"),
                "dyaOrig": obj.get(W + "dyaOrig"),
                "position": (
                    run.find("w:rPr/w:position", NS).get(W + "val")
                    if run.find("w:rPr/w:position", NS) is not None
                    else None
                ),
            },
        })

    files["word/document.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone="yes"
    )

    input_sha256 = _sha256_file(input_docx)
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(
        prefix="axmath-calibrated-",
        suffix=".docx",
        dir=str(output_docx.parent),
    )
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in files.items():
                z.writestr(name, data)
        if input_docx == output_docx:
            os.replace(tmp, output_docx)
        else:
            if output_docx.exists() and overwrite:
                output_docx.unlink()
            os.replace(tmp, output_docx)
    finally:
        tmp.unlink(missing_ok=True)

    return {
        "input": str(input_docx),
        "input_sha256": input_sha256,
        "output": str(output_docx),
        "output_sha256": _sha256_file(output_docx),
        "applied": applied,
        "count": len(applied),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--plan", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--ordinals", help="Comma-separated, visually confirmed calibration ordinals")
    ap.add_argument("--allow-in-place", action="store_true")
    ap.add_argument("--overwrite", action="store_true")
    ap.add_argument("--report")
    args = ap.parse_args()

    plan = json.loads(Path(args.plan).read_text(encoding="utf-8-sig"))
    result = apply_plan(
        Path(args.input),
        Path(args.output),
        plan,
        _parse_ordinals(args.ordinals),
        allow_in_place=args.allow_in_place,
        overwrite=args.overwrite,
    )
    if args.report:
        Path(args.report).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "count": result["count"],
        "ordinals": [x["ordinal"] for x in result["applied"]],
        "output": result["output"],
        "output_sha256": result["output_sha256"],
    }, ensure_ascii=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
