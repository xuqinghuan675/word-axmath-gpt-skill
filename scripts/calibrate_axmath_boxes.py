from __future__ import annotations

import argparse
import json
import re
import shutil
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


def apply_plan(input_docx: Path, output_docx: Path, plan_report: dict):
    plan = {
        int(row["ordinal"]): row
        for row in plan_report.get("calibration_plan", [])
        if row.get("auto_apply")
    }
    if not plan:
        if input_docx.resolve() != output_docx.resolve():
            shutil.copy2(input_docx, output_docx)
        return {"applied": [], "count": 0, "output": str(output_docx.resolve())}

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
        style = _replace_style_dimension(style, "width", float(row["target_width_pt"]))
        style = _replace_style_dimension(style, "height", float(row["target_height_pt"]))
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

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    same = input_docx.resolve() == output_docx.resolve()
    if same:
        tmp = Path(tempfile.mkstemp(prefix="axmath-calibrated-", suffix=".docx")[1])
    else:
        tmp = output_docx
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in files.items():
                z.writestr(name, data)
        if same:
            shutil.move(str(tmp), str(output_docx))
    finally:
        if same and tmp.exists():
            tmp.unlink(missing_ok=True)

    return {
        "input": str(input_docx.resolve()),
        "output": str(output_docx.resolve()),
        "applied": applied,
        "count": len(applied),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--plan", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--report")
    args = ap.parse_args()

    plan = json.loads(Path(args.plan).read_text(encoding="utf-8-sig"))
    result = apply_plan(Path(args.input).resolve(), Path(args.output).resolve(), plan)
    if args.report:
        Path(args.report).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "count": result["count"],
        "ordinals": [x["ordinal"] for x in result["applied"]],
        "output": result["output"],
    }, ensure_ascii=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
