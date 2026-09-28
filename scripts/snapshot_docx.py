from __future__ import annotations

import argparse
import hashlib
import json
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import numpy as np
from PIL import Image

from word_runtime import OwnedWord, clean_com_text, text_anchor

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
RELNS = {"pr": "http://schemas.openxmlformats.org/package/2006/relationships"}
W = "{%s}" % NS["w"]


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _safe_text(s: str | None) -> str:
    return clean_com_text(s).replace("\r", "\n").replace("\x07", "")


def _page_info(rng):
    def info(code, default=None):
        try:
            v = rng.Information(code)
            if isinstance(v, (int, float)):
                return float(v)
            return v
        except Exception:
            return default
    return {
        "page": int(info(3, 0) or 0),
        "x_pt": info(5),
        "y_pt": info(6),
    }


def _range_visual_geometry(rng, y_tolerance_pt: float = 1.5):
    """Measure a Word range using collapsed start/end layout coordinates.

    Word's Range.Information on a non-collapsed range is not a stable start
    coordinate for complex OfficeMath. Always collapse explicit duplicates at
    both ends before measuring geometry.
    """
    try:
        start_rng = rng.Duplicate
        start_rng.Collapse(1)  # wdCollapseStart
        start = _page_info(start_rng)
    except Exception:
        start = _page_info(rng)
    try:
        end_rng = rng.Duplicate
        end_rng.Collapse(0)  # wdCollapseEnd
        end = _page_info(end_rng)
    except Exception:
        end = None

    rec = {
        "end_page": end.get("page") if end else None,
        "end_x_pt": end.get("x_pt") if end else None,
        "end_y_pt": end.get("y_pt") if end else None,
        "same_visual_line": False,
        "visual_width_pt": None,
        "visual_width_reliable": False,
    }
    if not end:
        return rec
    sx, sy = start.get("x_pt"), start.get("y_pt")
    ex, ey = end.get("x_pt"), end.get("y_pt")
    if (
        start.get("page") == end.get("page")
        and sx is not None and sy is not None
        and ex is not None and ey is not None
        and abs(float(ey) - float(sy)) <= y_tolerance_pt
        and float(ex) >= float(sx)
    ):
        width = float(ex) - float(sx)
        rec["same_visual_line"] = True
        rec["visual_width_pt"] = width
        rec["visual_width_reliable"] = width > 0.5
    return rec


def _font_size(rng):
    try:
        v = float(rng.Font.Size)
        if 0 < v < 500:
            return v
    except Exception:
        pass
    return None


def _paragraph_format(rng):
    try:
        fmt = rng.Paragraphs.Item(1).Format
    except Exception:
        return None

    def num(name):
        try:
            value = getattr(fmt, name)
            return float(value) if isinstance(value, (int, float)) else None
        except Exception:
            return None

    def integer(name):
        try:
            return int(getattr(fmt, name))
        except Exception:
            return None

    return {
        "alignment": integer("Alignment"),
        "left_indent_pt": num("LeftIndent"),
        "right_indent_pt": num("RightIndent"),
        "first_line_indent_pt": num("FirstLineIndent"),
        "space_before_pt": num("SpaceBefore"),
        "space_after_pt": num("SpaceAfter"),
        "line_spacing_pt": num("LineSpacing"),
        "line_spacing_rule": integer("LineSpacingRule"),
    }


def _source_inline_context(doc, rng):
    """Return conservative evidence that native OfficeMath is inline."""
    try:
        para = rng.Paragraphs.Item(1).Range.Duplicate
        pstart = int(para.Start)
        pend = int(para.End)
        omath_count = int(para.OMaths.Count)
    except Exception:
        return {
            "paragraph_end": None,
            "paragraph_omath_count": None,
            "source_inline_context": False,
            "source_inline_context_strong": False,
            "source_inline_reasons": [],
        }

    try:
        before = _safe_text(doc.Range(pstart, int(rng.Start)).Text)
    except Exception:
        before = ""
    try:
        after_end = max(int(rng.End), pend - 1)
        after = _safe_text(doc.Range(int(rng.End), after_end).Text)
    except Exception:
        after = ""

    def meaningful(text):
        return bool(re.sub(r"[\s\u00a0]+", "", text or ""))

    fmt = _paragraph_format(rng) or {}
    before_meaningful = meaningful(before)
    after_meaningful = meaningful(after)
    reasons = []
    if omath_count > 1:
        reasons.append("multiple_omath_same_paragraph")
    if before_meaningful:
        reasons.append("meaningful_content_before_formula")
    if after_meaningful:
        reasons.append("meaningful_content_after_formula")

    # Strong evidence deliberately ignores after-only text: a numbered display
    # equation can have a right-side number in the same paragraph. Another
    # OMath in the paragraph, or meaningful content before the formula in a
    # non-centered paragraph, is much safer evidence for automatic inline repair.
    strong = omath_count > 1 or (
        before_meaningful and fmt.get("alignment") != 1
    )
    return {
        "paragraph_end": pend,
        "paragraph_omath_count": omath_count,
        "paragraph_content_before": before[-160:],
        "paragraph_content_after": after[:160],
        "source_inline_context": bool(reasons),
        "source_inline_context_strong": bool(strong),
        "source_inline_reasons": reasons,
    }


def _collect_com_inventory(doc):
    items = {"omath": [], "axmath": []}

    for i in range(1, int(doc.OMaths.Count) + 1):
        m = doc.OMaths.Item(i)
        r = m.Range
        rec = {
            "formula_id": f"omath-{i:04d}",
            "ordinal": i,
            "start": int(r.Start),
            "end": int(r.End),
            "text": _safe_text(r.Text),
            "font_size_pt": _font_size(r),
        }
        rec.update(_page_info(r))
        rec.update(_range_visual_geometry(r))
        rec.update(text_anchor(doc, r.Start, r.End))
        try:
            rec["paragraph_start"] = int(r.Paragraphs.Item(1).Range.Start)
        except Exception:
            rec["paragraph_start"] = None
        rec.update(_source_inline_context(doc, r))
        rec["paragraph_format"] = _paragraph_format(r)
        items["omath"].append(rec)

    ax_idx = 0
    for i in range(1, int(doc.InlineShapes.Count) + 1):
        s = doc.InlineShapes.Item(i)
        try:
            prog = str(s.OLEFormat.ProgID)
        except Exception:
            continue
        if prog.lower() != "equation.axmath":
            continue
        ax_idx += 1
        r = s.Range
        rec = {
            "formula_id": f"axmath-{ax_idx:04d}",
            "ordinal": ax_idx,
            "inline_shape_index": i,
            "start": int(r.Start),
            "end": int(r.End),
            "width_pt": float(s.Width),
            "height_pt": float(s.Height),
            "font_size_pt": _font_size(r),
            "progid": prog,
        }
        rec.update(_page_info(r))
        rec.update(_range_visual_geometry(r))
        rec.update(text_anchor(doc, r.Start, r.End))
        fs = rec.get("font_size_pt")
        rec["height_font_ratio"] = (rec["height_pt"] / fs) if fs and fs > 0 else None
        try:
            rec["paragraph_start"] = int(r.Paragraphs.Item(1).Range.Start)
        except Exception:
            rec["paragraph_start"] = None
        rec["paragraph_format"] = _paragraph_format(r)
        items["axmath"].append(rec)
    return items


def _classify_plain_math_run(
    text: str,
    paragraph_text: str,
    run_start: int = 0,
    vert_align: str | None = None,
):
    raw = text or ""
    s = raw.strip()
    para = paragraph_text or ""
    if not s:
        return False, None

    normalized = unicodedata.normalize("NFKC", s)
    run_end = run_start + len(raw)
    before = para[:run_start].rstrip()
    after = para[run_end:].lstrip()

    # Preserve question/list/option labels: 1、 2. (3) （4） 例5、 A、 B、 etc.
    numbering = re.match(
        r"^\s*(?:例\s*\d+[、.．]|[（(]?\d+[)）]|[（(]?\d+[、.．]|[A-Da-d][、.．]|[一二三四五六七八九十]+[、.．])",
        para,
    )
    if numbering and run_start < numbering.end() and run_end <= numbering.end():
        return False, "paragraph_numbering"
    if re.fullmatch(r"[0-9A-Za-z一二三四五六七八九十]+", normalized):
        if after.startswith(("、", "．", ".")):
            return False, "paragraph_numbering"
        if normalized.isdigit() and after.startswith((")", "）")):
            return False, "paragraph_numbering"
        if normalized.isdigit() and before.endswith(("例", "题")):
            return False, "paragraph_numbering"
        if re.fullmatch(r"[A-Za-z]", normalized) and not before and re.match(r"^[\u4e00-\u9fff]", after):
            return False, "option_label"

    if vert_align in {"superscript", "subscript"} and re.fullmatch(
        r"(?:[+-]?\d+(?:\.\d+)?|[A-Za-zΑ-Ωα-ω]+)", normalized
    ):
        return True, f"{vert_align}_math"

    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?", normalized):
        return True, "number"
    if re.fullmatch(r"[A-Za-zΑ-Ωα-ω]+", normalized):
        return True, "letters"
    if re.fullmatch(
        r"(?:[A-Za-zΑ-Ωα-ω0-9]+\s*/\s*[A-Za-zΑ-Ωα-ω0-9]+|[½⅓⅔¼¾⅕⅖⅗⅘⅙⅚⅛⅜⅝⅞])",
        normalized,
    ):
        return True, "simple_fraction"
    if (
        re.search(
            r"[=<>≤≥±×÷∑∏∫√∞≈≠∀∃∄∈∉∋∌⊂⊆⊇⊃∪∩∧∨¬⊕⊗→←↔⇔⇒⟹⟺^_]",
            normalized,
        )
        and not re.search(r"[\u3400-\u9fff]", normalized)
    ):
        return True, "math_operator"
    if re.fullmatch(
        r"[A-Za-zΑ-Ωα-ω0-9]+(?:\s*[+\-*/=<>]\s*[A-Za-zΑ-Ωα-ω0-9]+)+",
        normalized,
    ):
        return True, "math_expression"
    return False, None


def _cambria_runs(docx: Path):
    out = []
    with zipfile.ZipFile(docx) as z:
        root = ET.fromstring(z.read("word/document.xml"))
        paras = root.findall(".//w:p", NS)
        for pi, p in enumerate(paras, 1):
            para_text = "".join(t.text or "" for t in p.findall(".//w:t", NS))
            char_offset = 0
            same_text_counts = {}
            for ri, run in enumerate(p.findall(".//w:r", NS), 1):
                text = "".join(t.text or "" for t in run.findall("w:t", NS))
                run_start = char_offset
                char_offset += len(text)
                if not text:
                    continue
                occurrence = same_text_counts.get(text, 0) + 1
                same_text_counts[text] = occurrence
                rpr = run.find("w:rPr", NS)
                if rpr is None:
                    continue
                fonts = rpr.find("w:rFonts", NS)
                names = list(fonts.attrib.values()) if fonts is not None else []
                if not any("cambria math" in (n or "").lower() for n in names):
                    continue
                va = rpr.find("w:vertAlign", NS)
                vert_align = va.attrib.get(W + "val") if va is not None else None
                sz = rpr.find("w:sz", NS)
                is_candidate, candidate_reason = _classify_plain_math_run(
                    text,
                    para_text,
                    run_start=run_start,
                    vert_align=vert_align,
                )
                out.append({
                    "paragraph": pi,
                    "run": ri,
                    "text": text,
                    "paragraph_text": para_text[:240],
                    "run_start_in_paragraph": run_start,
                    "same_text_occurrence": occurrence,
                    "vert_align": vert_align,
                    "size_half_points": int(sz.attrib.get(W + "val")) if sz is not None and (sz.attrib.get(W + "val") or "").isdigit() else None,
                    "plain_math_candidate": bool(is_candidate),
                    "candidate_reason": candidate_reason,
                })
    return out


def _plain_math_runs(docx: Path):
    """High-confidence math-like text runs that are not native OfficeMath.

    This intentionally scans all Word text runs, not only Cambria Math, because
    customers often type a lone number/letter/fraction in the surrounding body
    font. Numbering/list/option labels are excluded by _classify_plain_math_run.
    """
    out = []
    with zipfile.ZipFile(docx) as z:
        root = ET.fromstring(z.read("word/document.xml"))
        paras = root.findall(".//w:p", NS)
        for pi, p in enumerate(paras, 1):
            para_text = "".join(t.text or "" for t in p.findall(".//w:t", NS))
            char_offset = 0
            same_text_counts = {}
            for ri, run in enumerate(p.findall(".//w:r", NS), 1):
                # Runs already inside OfficeMath are not plain-text candidates.
                parent_is_math = False
                # ElementTree has no parent pointer; the paragraph-level run
                # selection above reaches m:oMath descendants too, so reject
                # runs whose XML is under an m:oMath by comparing object ids.
                for om in p.findall(".//m:oMath", NS):
                    if any(id(run) == id(x) for x in om.findall(".//w:r", NS)):
                        parent_is_math = True
                        break
                text = "".join(t.text or "" for t in run.findall("w:t", NS))
                run_start = char_offset
                char_offset += len(text)
                if not text or parent_is_math:
                    continue
                occurrence = same_text_counts.get(text, 0) + 1
                same_text_counts[text] = occurrence
                rpr = run.find("w:rPr", NS)
                fonts = []
                vert_align = None
                italic = False
                if rpr is not None:
                    rf = rpr.find("w:rFonts", NS)
                    fonts = list(rf.attrib.values()) if rf is not None else []
                    va = rpr.find("w:vertAlign", NS)
                    vert_align = va.attrib.get(W + "val") if va is not None else None
                    italic = rpr.find("w:i", NS) is not None
                is_candidate, reason = _classify_plain_math_run(
                    text,
                    para_text,
                    run_start=run_start,
                    vert_align=vert_align,
                )
                if not is_candidate:
                    continue
                out.append({
                    "paragraph": pi,
                    "run": ri,
                    "text": text,
                    "paragraph_text": para_text[:240],
                    "run_start_in_paragraph": run_start,
                    "same_text_occurrence": occurrence,
                    "candidate_reason": reason,
                    "fonts": fonts,
                    "italic": italic,
                    "vert_align": vert_align,
                })
    return out


def _preview_inventory(docx: Path, preview_dir: Path):
    preview_dir.mkdir(parents=True, exist_ok=True)
    rows = []
    with zipfile.ZipFile(docx) as z:
        names = set(z.namelist())
        root = ET.fromstring(z.read("word/document.xml"))
        relroot = ET.fromstring(z.read("word/_rels/document.xml.rels"))
        relmap = {r.attrib.get("Id"): r.attrib.get("Target") for r in relroot.findall("pr:Relationship", RELNS)}
        ax_idx = 0
        for obj in root.findall(".//w:object", NS):
            ole = obj.find(".//o:OLEObject", NS)
            if ole is None or (ole.attrib.get("ProgID") or "").lower() != "equation.axmath":
                continue
            ax_idx += 1
            image = obj.find(".//v:imagedata", NS)
            rid = image.attrib.get("{%s}id" % NS["r"]) if image is not None else None
            target = relmap.get(rid)
            ole_rid = ole.attrib.get("{%s}id" % NS["r"])
            ole_target = relmap.get(ole_rid)
            rec = {
                "ordinal": ax_idx,
                "rid": rid,
                "target": target,
                "preview_rid": rid,
                "preview_target": target,
                "ole_rid": ole_rid,
                "ole_target": ole_target,
                "preview_ok": False,
            }
            if ole_target:
                ole_internal = "word/" + ole_target.replace("\\", "/").lstrip("/")
                if ole_internal in names:
                    rec["ole_sha256"] = hashlib.sha256(z.read(ole_internal)).hexdigest()
            if target:
                internal = "word/" + target.replace("\\", "/").lstrip("/")
                if internal in names:
                    data = z.read(internal)
                    rec["preview_sha256"] = hashlib.sha256(data).hexdigest()
                    suffix = Path(target).suffix or ".bin"
                    tmp = preview_dir / f"preview-{ax_idx:04d}{suffix}"
                    tmp.write_bytes(data)
                    try:
                        im = Image.open(tmp)
                        im.load()
                        if max(im.size) > 768:
                            im.thumbnail((768, 768))
                        g = im.convert("L")
                        arr = np.asarray(g, dtype=np.float32)
                        rec.update({
                            "preview_ok": True,
                            "width_px": int(g.width),
                            "height_px": int(g.height),
                            "mean": float(arr.mean()),
                            "std": float(arr.std()),
                            "dark_fraction": float((arr < 160).mean()),
                            "ink_fraction": float((arr < 230).mean()),
                        })
                        rec["preview_suspicious"] = bool(
                            (rec["std"] < 2.0 and rec["mean"] > 245)
                            or (rec["std"] < 3.0 and 90 < rec["mean"] < 235)
                            or rec["ink_fraction"] < 0.0005
                        )
                        png = preview_dir / f"preview-{ax_idx:04d}.png"
                        g.save(png)
                        rec["preview_png"] = str(png)
                    except Exception as e:
                        rec["preview_error"] = repr(e)
            rows.append(rec)

    preview_targets = {}
    ole_targets = {}
    preview_hashes = {}
    for rec in rows:
        if rec.get("preview_target"):
            preview_targets.setdefault(rec["preview_target"], []).append(rec["ordinal"])
        if rec.get("ole_target"):
            ole_targets.setdefault(rec["ole_target"], []).append(rec["ordinal"])
        if rec.get("preview_sha256"):
            preview_hashes.setdefault(rec["preview_sha256"], []).append(rec["ordinal"])
    for rec in rows:
        preview_group = preview_targets.get(rec.get("preview_target"), [])
        ole_group = ole_targets.get(rec.get("ole_target"), [])
        hash_group = preview_hashes.get(rec.get("preview_sha256"), [])
        rec["shared_preview_target_ordinals"] = preview_group if len(preview_group) > 1 else []
        rec["shared_ole_target_ordinals"] = ole_group if len(ole_group) > 1 else []
        rec["same_preview_hash_ordinals"] = hash_group if len(hash_group) > 1 else []
    return rows


def _axmath_layout_inventory(docx: Path):
    rows = []
    with zipfile.ZipFile(docx) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    ordinal = 0
    for paragraph_index, p in enumerate(root.findall(".//w:p", NS), 1):
        run_index = 0
        for run in p.findall(".//w:r", NS):
            run_index += 1
            obj = run.find(".//w:object", NS)
            if obj is None:
                continue
            ole = obj.find(".//o:OLEObject", NS)
            if ole is None or (ole.attrib.get("ProgID") or "").lower() != "equation.axmath":
                continue
            ordinal += 1
            rpr = run.find("w:rPr", NS)
            pos = None
            if rpr is not None:
                position = rpr.find("w:position", NS)
                if position is not None:
                    raw = position.attrib.get(W + "val")
                    try:
                        pos = int(raw)
                    except (TypeError, ValueError):
                        pos = raw
            shape = obj.find(".//v:shape", NS)
            style = shape.attrib.get("style") if shape is not None else None
            dxa_orig = obj.attrib.get(W + "dxaOrig")
            dya_orig = obj.attrib.get(W + "dyaOrig")

            def as_int(value):
                try:
                    return int(value)
                except (TypeError, ValueError):
                    return None

            def style_pt(name):
                if not style:
                    return None
                match = re.search(rf"(?:^|;){name}:([0-9.]+)pt(?:;|$)", style)
                return float(match.group(1)) if match else None

            width_pt = style_pt("width")
            height_pt = style_pt("height")
            position_points = (pos / 2.0) if isinstance(pos, int) else None
            baseline_offset_ratio = (
                abs(position_points) / height_pt
                if position_points is not None and height_pt and height_pt > 0
                else None
            )
            rows.append({
                "ordinal": ordinal,
                "paragraph_index": paragraph_index,
                "run_index": run_index,
                "run_position_half_points": pos,
                "run_position_points": position_points,
                "shape_style": style,
                "shape_width_pt": width_pt,
                "shape_height_pt": height_pt,
                "dxa_orig": as_int(dxa_orig),
                "dya_orig": as_int(dya_orig),
                "baseline_offset_ratio": baseline_offset_ratio,
                "baseline_suspicious": bool(
                    baseline_offset_ratio is not None and baseline_offset_ratio > 0.70
                ),
            })
    return rows


def _inline_baseline_groups(layout_rows, threshold_ratio: float = 0.70):
    groups = {}
    for row in layout_rows:
        groups.setdefault(row["paragraph_index"], []).append(row)
    out = []
    for paragraph_index, rows in groups.items():
        suspicious = [
            r
            for r in rows
            if r.get("baseline_offset_ratio") is not None
            and float(r["baseline_offset_ratio"]) > threshold_ratio
        ]
        if not suspicious:
            continue
        numeric = [
            r["run_position_half_points"]
            for r in rows
            if isinstance(r.get("run_position_half_points"), int)
        ]
        spread = (max(numeric) - min(numeric)) if len(numeric) >= 2 else 0
        out.append({
            "paragraph_index": paragraph_index,
            "ordinals": [r["ordinal"] for r in rows],
            "suspicious_ordinals": [r["ordinal"] for r in suspicious],
            "run_positions_half_points": [r.get("run_position_half_points") for r in rows],
            "position_spread_half_points": spread,
            "position_spread_points": spread / 2.0,
            "baseline_offset_ratios": [r.get("baseline_offset_ratio") for r in rows],
            "shape_styles": [r.get("shape_style") for r in rows],
        })
    return out


def _render_pdf(pdf: Path, pages_dir: Path, zoom: float = 2.0):
    import fitz
    pages_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf)
    result = []
    for i, page in enumerate(doc, 1):
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        p = pages_dir / f"page-{i:03d}.png"
        pix.save(str(p))
        result.append({
            "page": i,
            "path": str(p),
            "width_px": pix.width,
            "height_px": pix.height,
            "width_pt": float(page.rect.width),
            "height_pt": float(page.rect.height),
        })
    doc.close()
    return result


def _crop_formula_pages(inventory, pages, crop_dir: Path):
    crop_dir.mkdir(parents=True, exist_ok=True)
    page_map = {p["page"]: p for p in pages}
    for kind, rows in inventory.items():
        for rec in rows:
            pg = int(rec.get("page") or 0)
            meta = page_map.get(pg)
            if not meta:
                continue
            x, y = rec.get("x_pt"), rec.get("y_pt")
            if x is None or y is None or x < 0 or y < 0:
                continue
            page_img = Image.open(meta["path"]).convert("RGB")
            sx = meta["width_px"] / meta["width_pt"]
            sy = meta["height_px"] / meta["height_pt"]
            if kind == "axmath":
                w = max(60.0, float(rec.get("width_pt") or 120.0))
                h = max(18.0, float(rec.get("height_pt") or 18.0))
            else:
                # Native OfficeMath crops follow Word's measured source range
                # width when it stayed on one line; fall back only when Word
                # cannot provide a reliable range geometry.
                w = max(40.0, float(rec.get("visual_width_pt") or 220.0))
                h = max(22.0, float(rec.get("font_size_pt") or 12.0) * 2.2)
            box = (
                max(0, int((x - 45) * sx)),
                max(0, int((y - 22) * sy)),
                min(page_img.width, int((x + w + 80) * sx)),
                min(page_img.height, int((y + h + 30) * sy)),
            )
            if box[2] <= box[0] or box[3] <= box[1]:
                continue
            crop = page_img.crop(box)
            out = crop_dir / f"{rec['formula_id']}.png"
            crop.save(out)
            rec["crop_png"] = str(out)
            rec["crop_box_px"] = list(box)


def snapshot(docx: Path, outdir: Path, label: str | None = None, profile: str = "full"):
    docx = docx.resolve()
    outdir = outdir.resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    pdf = outdir / "render.pdf"
    if profile not in {"full", "geometry"}:
        raise ValueError(f"unsupported snapshot profile: {profile}")
    source_sha256 = _sha256_file(docx)
    report = {
        "source": str(docx),
        "label": label or docx.stem,
        "profile": profile,
        "status": "starting",
        "docx_sha256": source_sha256,
    }

    session_meta = None
    with OwnedWord(visible=False, require_clean=False) as (word, meta):
        # Read-only snapshots must not stall on Word's interactive recovery prompt.
        doc = word.Documents.OpenNoRepairDialog(str(docx), False, True, False)
        try:
            report["pages_word"] = int(doc.ComputeStatistics(2))
            report["words"] = int(doc.ComputeStatistics(0))
            report["chars"] = int(doc.ComputeStatistics(3))
            report["paragraphs_com"] = int(doc.Paragraphs.Count)
            report["inventory"] = _collect_com_inventory(doc)
            if profile == "full":
                doc.ExportAsFixedFormat(str(pdf), 17, OpenAfterExport=False)
        finally:
            doc.Close(False)
        session_meta = meta

    report["word_session"] = session_meta.to_dict() if session_meta else None
    report["pages"] = _render_pdf(pdf, outdir / "pages", 2.0) if profile == "full" else []
    report["cambria_math_runs"] = _cambria_runs(docx)
    report["plain_math_candidates"] = _plain_math_runs(docx)
    report["preview_inventory"] = (
        _preview_inventory(docx, outdir / "formula_previews") if profile == "full" else []
    )
    report["axmath_layout_inventory"] = _axmath_layout_inventory(docx)
    report["inline_baseline_groups"] = _inline_baseline_groups(report["axmath_layout_inventory"])
    if profile == "full":
        _crop_formula_pages(report["inventory"], report["pages"], outdir / "formula_crops")

    preview_by_ord = {x["ordinal"]: x for x in report["preview_inventory"]}
    report["relationship_risks"] = {
        "shared_preview_targets": [
            {
                "target": p["preview_target"],
                "ordinals": p["shared_preview_target_ordinals"],
            }
            for p in report["preview_inventory"]
            if p.get("shared_preview_target_ordinals")
            and p["ordinal"] == min(p["shared_preview_target_ordinals"])
        ],
        "shared_ole_targets": [
            {
                "target": p["ole_target"],
                "ordinals": p["shared_ole_target_ordinals"],
            }
            for p in report["preview_inventory"]
            if p.get("shared_ole_target_ordinals")
            and p["ordinal"] == min(p["shared_ole_target_ordinals"])
        ],
    }
    anomalies = []
    for rec in report["inventory"]["axmath"]:
        flags = []
        ratio = rec.get("height_font_ratio")
        if ratio and ratio > 2.8:
            flags.append("height_vs_font")
        if rec.get("width_pt", 0) > 520:
            flags.append("very_wide")
        p = preview_by_ord.get(rec["ordinal"])
        if p and p.get("preview_suspicious"):
            flags.append("preview_suspicious")
        if flags:
            anomalies.append({
                "formula_id": rec["formula_id"],
                "ordinal": rec["ordinal"],
                "page": rec.get("page"),
                "flags": flags,
                "crop_png": rec.get("crop_png"),
                "before": rec.get("before"),
                "after": rec.get("after"),
                "preview_target": p.get("preview_target") if p else None,
                "ole_target": p.get("ole_target") if p else None,
                "shared_preview_target_ordinals": p.get("shared_preview_target_ordinals", []) if p else [],
                "shared_ole_target_ordinals": p.get("shared_ole_target_ordinals", []) if p else [],
            })
    report["anomalies"] = anomalies
    report["docx_sha256_after"] = _sha256_file(docx)
    report["docx_stable"] = report["docx_sha256_after"] == source_sha256
    report["status"] = "ready" if report["docx_stable"] else "changed_during_snapshot"
    out = outdir / "snapshot.json"
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("outdir")
    ap.add_argument("--label")
    ap.add_argument("--profile", choices=("full", "geometry"), default="full")
    args = ap.parse_args()
    rep = snapshot(Path(args.docx), Path(args.outdir), args.label, args.profile)
    print(json.dumps({
        "status": rep["status"],
        "source": rep["source"],
        "profile": rep.get("profile"),
        "docx_sha256": rep.get("docx_sha256"),
        "docx_stable": rep.get("docx_stable"),
        "pages": len(rep["pages"]),
        "omath": len(rep["inventory"]["omath"]),
        "axmath": len(rep["inventory"]["axmath"]),
        "cambria_runs": len(rep["cambria_math_runs"]),
        "anomalies": len(rep["anomalies"]),
        "inline_baseline_groups": len(rep.get("inline_baseline_groups", [])),
        "shared_preview_targets": len(rep.get("relationship_risks", {}).get("shared_preview_targets", [])),
        "shared_ole_targets": len(rep.get("relationship_risks", {}).get("shared_ole_targets", [])),
        "snapshot": str(Path(args.outdir).resolve() / "snapshot.json"),
    }, ensure_ascii=True, indent=2))
    return 0 if rep.get("status") == "ready" else 2


if __name__ == "__main__":
    raise SystemExit(main())
