from __future__ import annotations

import argparse
import hashlib
import json
import os
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = f"{{{W_NS}}}"
M = f"{{{M_NS}}}"
EXCLUDED_ANCESTORS = {
    M + "oMath",
    M + "oMathPara",
    W + "object",
    W + "drawing",
    W + "pict",
    W + "txbxContent",
    W + "fldSimple",
    W + "instrText",
}


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _eligible_body_text_nodes(root):
    body = root.find(".//" + W + "body")
    if body is None:
        raise ValueError("word/document.xml has no w:body")
    field_depth = 0
    for node in body.iter():
        if node.tag == W + "fldChar":
            field_type = node.get(W + "fldCharType")
            if field_type == "begin":
                field_depth += 1
            elif field_type == "end":
                field_depth = max(0, field_depth - 1)
            continue
        if node.tag != W + "t" or field_depth:
            continue
        if any(parent.tag in EXCLUDED_ANCESTORS for parent in node.iterancestors()):
            continue
        yield node


def _replace_body_periods(xml_bytes: bytes) -> tuple[bytes, int, int]:
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    root = etree.fromstring(xml_bytes, parser)
    replacements = 0
    for node in _eligible_body_text_nodes(root):
        text = node.text or ""
        replacements += text.count("。")
        if "。" in text:
            node.text = text.replace("。", ".")
    remaining = sum((node.text or "").count("。") for node in _eligible_body_text_nodes(root))
    output = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return output, replacements, remaining


def normalize_body_punctuation(source: Path, output: Path, report_path: Path) -> dict:
    source = source.resolve()
    output = output.resolve()
    report_path = report_path.resolve()
    if source == output or source == report_path or output == report_path:
        raise ValueError("Source DOCX, output DOCX, and report must be three different files.")
    if not source.is_file():
        raise FileNotFoundError(source)
    if output.exists():
        raise FileExistsError(f"Refusing to overwrite existing output: {output}")
    if report_path.exists():
        raise FileExistsError(f"Refusing to overwrite existing report: {report_path}")

    source_hash_before = sha256_file(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source, "r") as src:
        if "word/document.xml" not in src.namelist():
            raise ValueError("DOCX package is missing word/document.xml")
        document_xml, replacement_count, remaining_count = _replace_body_periods(
            src.read("word/document.xml")
        )
        if remaining_count:
            raise ValueError(f"Ordinary body text still contains {remaining_count} Chinese full stops")
        fd, temp_name = tempfile.mkstemp(prefix=output.stem + ".", suffix=".tmp", dir=output.parent)
        os.close(fd)
        temp_path = Path(temp_name)
        try:
            with zipfile.ZipFile(temp_path, "w") as dst:
                for item in src.infolist():
                    payload = document_xml if item.filename == "word/document.xml" else src.read(item.filename)
                    dst.writestr(item, payload)
        except Exception:
            temp_path.unlink(missing_ok=True)
            raise
    os.replace(temp_path, output)

    with zipfile.ZipFile(source, "r") as src, zipfile.ZipFile(output, "r") as dst:
        source_names = set(src.namelist())
        output_names = set(dst.namelist())
        if source_names != output_names:
            raise RuntimeError("DOCX package parts changed during punctuation normalization")
        changed_parts = [
            name for name in sorted(source_names)
            if src.read(name) != dst.read(name)
        ]
        if changed_parts != ["word/document.xml"]:
            raise RuntimeError(f"Unexpected package parts changed: {changed_parts}")
        _, _, residual = _replace_body_periods(dst.read("word/document.xml"))
        if residual:
            raise RuntimeError(f"Output still contains {residual} ordinary body Chinese full stops")

    source_hash_after = sha256_file(source)
    if source_hash_before != source_hash_after:
        raise RuntimeError("Source DOCX changed during punctuation normalization")
    report = {
        "schema": "axmath-body-punctuation-normalization/v1",
        "source": str(source),
        "source_sha256": source_hash_before,
        "source_unchanged": True,
        "output": str(output),
        "output_sha256": sha256_file(output),
        "replaced_count": replacement_count,
        "remaining_ordinary_body_count": 0,
        "changed_package_parts": ["word/document.xml"],
        "ole_media_and_other_parts_preserved": True,
    }
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Replace Chinese full stops only in ordinary main-body Word text nodes."
    )
    parser.add_argument("--source", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--report", required=True)
    args = parser.parse_args()
    report = normalize_body_punctuation(
        Path(args.source), Path(args.output), Path(args.report)
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
