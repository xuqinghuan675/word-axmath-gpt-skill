"""Reuse genuine AxMath OLE for byte-identical OfficeMath formulas.

This is an optional pre/post-processing layer around the existing official
AxMath batch converter. It never invokes a different formula conversion engine.
"""

from __future__ import annotations

import argparse
from copy import deepcopy
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import posixpath
import re
import shutil
import time
from zipfile import ZipFile

from lxml import etree

from audit_docx import analyze as audit_analyze, compare as audit_compare


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
V = "{urn:schemas-microsoft-com:vml}"
O = "{urn:schemas-microsoft-com:office:office}"
PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
SCHEMA = "axmath-exact-duplicate-cache/v1"


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def math_nodes(document: etree._Element) -> list[etree._Element]:
    return list(document.findall(".//" + M + "oMath"))


def paragraph_formula_counts(path: Path, *, converted: bool) -> list[int]:
    with ZipFile(path) as package:
        document = etree.fromstring(package.read("word/document.xml"))
    paragraphs = document.findall(".//" + W + "p")
    index = {para: i for i, para in enumerate(paragraphs)}
    counts = [0] * len(paragraphs)
    nodes = (ole for ole in document.findall(".//" + O + "OLEObject") if ole.get("ProgID") == "Equation.AxMath") if converted else iter(math_nodes(document))
    for node in nodes:
        parent = node.getparent()
        while parent is not None and parent.tag != W + "p":
            parent = parent.getparent()
        if parent is None:
            raise ValueError("Math object is not inside a Word paragraph")
        counts[index[parent]] += 1
    return counts


def repeated_formulas(source: Path) -> dict:
    with ZipFile(source) as package:
        document = etree.fromstring(package.read("word/document.xml"))
    nodes = math_nodes(document)
    first = {}
    duplicates = {}
    unique_count = 0
    for source_index, node in enumerate(nodes):
        key = etree.tostring(node, method="c14n")
        if key not in first:
            first[key] = unique_count
            unique_count += 1
        else:
            duplicates[f"AXD{source_index:05d}"] = first[key]
    return {
        "source_formulas": len(nodes),
        "unique_formulas": unique_count,
        "skipped_duplicates": len(duplicates),
        "duplicate_template_indices": duplicates,
    }


def prepare(source: Path, outdir: Path, *, min_duplicates: int = 64) -> dict:
    source = source.resolve(strict=True)
    outdir = outdir.resolve()
    if source.suffix.lower() != ".docx":
        raise ValueError("Source must be a DOCX")
    if outdir.exists() and any(outdir.iterdir()):
        raise FileExistsError(f"Preparation directory is not empty: {outdir}")
    started = time.perf_counter()
    before_hash = sha256(source)
    source_audit = audit_analyze(source)
    if source_audit["axmath_ole"] or source_audit["non_axmath_ole_count"] or source_audit["nonmain_omath_count"]:
        raise ValueError("Exact duplicate cache requires a main-document OfficeMath-only source")
    with ZipFile(source) as package:
        document = etree.fromstring(package.read("word/document.xml"))
        infos = package.infolist()
        original_data = {info.filename: package.read(info.filename) for info in infos}
    nodes = math_nodes(document)
    existing_names = {node.get(W + "name") for node in document.findall(".//" + W + "bookmarkStart")}
    existing_ids = [int(node.get(W + "id")) for node in document.findall(".//" + W + "bookmarkStart") if (node.get(W + "id") or "").isdigit()]
    next_bookmark_id = max(existing_ids, default=0) + 1
    first = {}
    mapping = {}
    unique_count = 0
    for source_index, node in enumerate(nodes):
        key = etree.tostring(node, method="c14n")
        if key not in first:
            first[key] = unique_count
            unique_count += 1
            continue
        name = f"AXD{source_index:05d}"
        if name in existing_names:
            raise ValueError(f"Source already has reserved bookmark {name}")
        parent = node.getparent()
        target = parent if parent.tag == M + "oMathPara" else node
        host = target.getparent()
        if host.tag != W + "p":
            raise ValueError(f"Formula {source_index} has unsupported container {host.tag}")
        index = host.index(target)
        host.remove(target)
        begin = etree.Element(W + "bookmarkStart")
        begin.set(W + "id", str(next_bookmark_id))
        begin.set(W + "name", name)
        end = etree.Element(W + "bookmarkEnd")
        end.set(W + "id", str(next_bookmark_id))
        host.insert(index, begin)
        host.insert(index + 1, end)
        mapping[name] = first[key]
        next_bookmark_id += 1
    if len(mapping) < min_duplicates:
        raise ValueError(f"Only {len(mapping)} exact duplicates; normal conversion is preferable below {min_duplicates}")
    outdir.mkdir(parents=True, exist_ok=True)
    frozen_dir = outdir / "source"
    prepared_dir = outdir / "prepared"
    frozen_dir.mkdir()
    prepared_dir.mkdir()
    frozen = frozen_dir / source.name
    reduced = prepared_dir / f"{source.stem}_unique.docx"
    shutil.copy2(source, frozen)
    reduced_xml = etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)
    with ZipFile(reduced, "w") as output:
        for info in infos:
            output.writestr(info, reduced_xml if info.filename == "word/document.xml" else original_data[info.filename])
    if sha256(source) != before_hash or sha256(frozen) != before_hash:
        raise RuntimeError("Source changed while preparing exact duplicate cache")
    manifest = {
        "schema": SCHEMA,
        "original_source": str(source),
        "original_source_sha256": before_hash,
        "frozen_source": str(frozen),
        "reduced_source": str(reduced),
        "reduced_source_sha256": sha256(reduced),
        "source_formulas": len(nodes),
        "unique_formulas": unique_count,
        "skipped_duplicates": len(mapping),
        "duplicate_template_indices": mapping,
        "prepare_seconds": time.perf_counter() - started,
        "next_action": "Run the official axmath.py convert on reduced_source, then restore using its READY_FOR_GPT_REVIEW.json.",
    }
    manifest_path = outdir / "CACHE_MANIFEST.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return {key: value for key, value in manifest.items() if key != "duplicate_template_indices"} | {"manifest": str(manifest_path)}


def _next_id(used: set[int]) -> int:
    number = max(used, default=0) + 1
    used.add(number)
    return number


def _existing_numbered_parts(names: set[str], folder: str, stem: str, extension: str) -> set[int]:
    pattern = re.compile(rf"word/{re.escape(folder)}/{re.escape(stem)}(\d+)\.{re.escape(extension)}")
    return {int(match.group(1)) for name in names if (match := pattern.fullmatch(name))}


def _copy_relationship(
    node: etree._Element,
    relationship_id: str,
    *,
    kind: str,
    folder: str,
    relations: etree._Element,
    by_id: dict[str, etree._Element],
    originals: dict[str, bytes],
    copied: dict[str, bytes],
    rel_ids: set[int],
    part_ids: dict[tuple[str, str], set[int]],
    content_types: etree._Element,
) -> None:
    old_relation = by_id.get(relationship_id)
    if old_relation is None or old_relation.get("TargetMode") == "External":
        raise ValueError(f"Missing or external relationship {relationship_id}")
    target = old_relation.get("Target") or ""
    old_part = posixpath.normpath(posixpath.join("word", target))
    if not old_part.startswith("word/") or old_part not in originals:
        raise ValueError(f"Missing package part {old_part}")
    extension = PurePosixPath(old_part).suffix.lstrip(".").lower()
    if kind == "oleObject" and extension != "bin":
        raise ValueError(f"Unexpected AxMath OLE extension: {extension}")
    if not extension:
        raise ValueError("Preview part has no extension")
    key = (kind, extension)
    part_ids.setdefault(key, _existing_numbered_parts(set(originals) | set(copied), folder, kind, extension))
    new_part = f"word/{folder}/{kind}{_next_id(part_ids[key])}.{extension}"
    if new_part in originals or new_part in copied:
        raise ValueError(f"Generated duplicate part {new_part}")
    new_id = f"rId{_next_id(rel_ids)}"
    relation = etree.Element(PKG_REL + "Relationship")
    relation.set("Id", new_id)
    relation.set("Type", old_relation.get("Type"))
    relation.set("Target", f"{folder}/{PurePosixPath(new_part).name}")
    relations.append(relation)
    node.set(R + "id", new_id)
    copied[new_part] = originals[old_part]
    old_override = next((x for x in content_types if x.tag == CT + "Override" and x.get("PartName") == "/" + old_part), None)
    if old_override is not None:
        override = etree.Element(CT + "Override")
        override.set("PartName", "/" + new_part)
        override.set("ContentType", old_override.get("ContentType"))
        content_types.append(override)
    elif not any(x.tag == CT + "Default" and x.get("Extension", "").lower() == extension for x in content_types):
        raise ValueError(f"No content type for .{extension}")


def clone_ole(source: Path, target: Path, mapping: dict[str, int], expected_unique: int, expected_total: int) -> dict:
    if target.exists():
        raise FileExistsError(target)
    with ZipFile(source) as package:
        infos = package.infolist()
        originals = {info.filename: package.read(info.filename) for info in infos}
    document = etree.fromstring(originals["word/document.xml"])
    relations = etree.fromstring(originals["word/_rels/document.xml.rels"])
    content_types = etree.fromstring(originals["[Content_Types].xml"])
    objects = document.findall(".//" + W + "object")
    if len(objects) != expected_unique:
        raise ValueError(f"Expected {expected_unique} converted objects, got {len(objects)}")
    templates = [obj.getparent() for obj in objects]
    if any(run.tag != W + "r" or run.find(".//" + O + "OLEObject") is None or run.find(".//" + O + "OLEObject").get("ProgID") != "Equation.AxMath" for run in templates):
        raise ValueError("Expected only genuine AxMath OLE templates")
    by_id = {relation.get("Id"): relation for relation in relations}
    rel_ids = {int(match.group(1)) for relation in relations if (match := re.fullmatch(r"rId(\d+)", relation.get("Id", "")))}
    shape_ids = {int(match.group(1)) for shape in document.findall(".//" + V + "shape") if (match := re.fullmatch(r"_x0000_i(\d+)", shape.get("id", "")))}
    object_ids = {int(match.group(1)) for ole in document.findall(".//" + O + "OLEObject") if (match := re.fullmatch(r"_(\d+)", ole.get("ObjectID", "")))}
    part_ids: dict[tuple[str, str], set[int]] = {}
    copied: dict[str, bytes] = {}
    for marker, template_index in mapping.items():
        found = [node for node in document.findall(".//" + W + "bookmarkStart") if node.get(W + "name") == marker]
        if len(found) != 1:
            raise ValueError(f"Bookmark {marker}: expected once, got {len(found)}")
        start = found[0]
        end = None
        for sibling in start.itersiblings():
            if sibling.tag == W + "bookmarkEnd" and sibling.get(W + "id") == start.get(W + "id"):
                end = sibling
                break
            if sibling.tag not in {W + "bookmarkStart", W + "bookmarkEnd"}:
                raise ValueError(f"Bookmark {marker}: unexpected content inside placeholder")
        if end is None or start.getparent().tag != W + "p":
            raise ValueError(f"Bookmark {marker}: missing boundary or paragraph")
        run = deepcopy(templates[template_index])
        shape = run.find(".//" + V + "shape")
        ole = run.find(".//" + O + "OLEObject")
        preview = run.find(".//" + V + "imagedata")
        if shape is None or ole is None or preview is None:
            raise ValueError(f"Template {template_index}: missing VML/OLE content")
        new_shape = f"_x0000_i{_next_id(shape_ids)}"
        shape.set("id", new_shape)
        ole.set("ShapeID", new_shape)
        ole.set("ObjectID", f"_{_next_id(object_ids)}")
        _copy_relationship(ole, ole.get(R + "id"), kind="oleObject", folder="embeddings", relations=relations, by_id=by_id, originals=originals, copied=copied, rel_ids=rel_ids, part_ids=part_ids, content_types=content_types)
        _copy_relationship(preview, preview.get(R + "id"), kind="image", folder="media", relations=relations, by_id=by_id, originals=originals, copied=copied, rel_ids=rel_ids, part_ids=part_ids, content_types=content_types)
        paragraph = start.getparent()
        index = paragraph.index(start)
        paragraph.remove(end)
        paragraph.remove(start)
        paragraph.insert(index, run)
    if len(document.findall(".//" + W + "object")) != expected_total:
        raise ValueError("Final AxMath count differs from original OfficeMath count")
    updates = {
        "word/document.xml": etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True),
        "word/_rels/document.xml.rels": etree.tostring(relations, xml_declaration=True, encoding="UTF-8", standalone=True),
        "[Content_Types].xml": etree.tostring(content_types, xml_declaration=True, encoding="UTF-8", standalone=True),
    }
    with ZipFile(source) as package, ZipFile(target, "w") as output:
        for info in infos:
            output.writestr(info, updates.get(info.filename, package.read(info.filename)))
        for name, data in sorted(copied.items()):
            output.writestr(name, data)
    return {"cloned_objects": len(mapping), "copied_parts": len(copied), "output_sha256": sha256(target)}


def restore(manifest_path: Path, ready_path: Path, output: Path, *, raw: bool = False, numbering_scope: str = "line-start") -> dict:
    started = time.perf_counter()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    if manifest.get("schema") != SCHEMA:
        raise ValueError("Unexpected cache manifest schema")
    original = Path(manifest["frozen_source"])
    reduced = Path(manifest["reduced_source"])
    if sha256(original) != manifest["original_source_sha256"] or sha256(reduced) != manifest["reduced_source_sha256"]:
        raise ValueError("Frozen source or prepared unique source changed")
    if sha256(Path(manifest["original_source"])) != manifest["original_source_sha256"]:
        raise ValueError("Original source changed")
    actual_repeats = repeated_formulas(original)
    if any(actual_repeats[key] != manifest[key] for key in ("source_formulas", "unique_formulas", "skipped_duplicates", "duplicate_template_indices")):
        raise ValueError("Cache map does not match frozen OfficeMath source")
    ready = json.loads(ready_path.read_text(encoding="utf-8-sig"))
    docs = ready.get("documents") or []
    if ready.get("status") != "ready" or len(docs) != 1:
        raise ValueError("Official reduced conversion is not ready")
    doc = docs[0]
    if (doc.get("status") != "converted_ready_for_gpt_review"
            or doc.get("original_source_sha256") != manifest["reduced_source_sha256"]
            or doc.get("frozen_source_sha256") != manifest["reduced_source_sha256"]):
        raise ValueError("Official conversion source hash or status mismatch")
    if (doc.get("source_omath") != manifest["unique_formulas"]
            or doc.get("working_axmath") != manifest["unique_formulas"]
            or doc.get("working_omath") != 0):
        raise ValueError("Official conversion formula count mismatch")
    raw_working = Path(doc["conversion_working_docx"])
    checkpoint = Path(str(raw_working) + ".conversion-checkpoint.json")
    state = json.loads(checkpoint.read_text(encoding="utf-8-sig"))
    if (state.get("status") != "complete" or state.get("remaining_omath") != 0
            or state.get("source_sha256") != manifest["reduced_source_sha256"]
            or sha256(raw_working) != state.get("working_sha256")):
        raise ValueError("Official conversion checkpoint is not complete or hash-bound")
    converted = raw_working if raw else Path(doc["working_docx"])
    if paragraph_formula_counts(reduced, converted=False) != paragraph_formula_counts(converted, converted=True):
        raise ValueError("Official conversion moved formulas between paragraphs")
    output = output.resolve()
    if output.exists():
        raise FileExistsError(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(output.name + ".unreviewed")
    if temporary.exists():
        raise FileExistsError(temporary)
    clone_started = time.perf_counter()
    cloned = clone_ole(converted, temporary, manifest["duplicate_template_indices"], manifest["unique_formulas"], manifest["source_formulas"])
    clone_seconds = time.perf_counter() - clone_started
    if paragraph_formula_counts(original, converted=False) != paragraph_formula_counts(temporary, converted=True):
        raise RuntimeError(f"Restored candidate moved formulas between paragraphs; inspect {temporary}")
    audit = audit_compare(original, temporary, numbering_scope=numbering_scope)
    exact = (audit.get("formula_count_state") or {}).get("state") == "exact"
    acceptable = bool(exact and audit.get("paragraph_count_equal") and audit.get("nonmath_text_contract_exact") and audit.get("candidate_non_axmath_ole_count") == 0 and audit.get("candidate_nonmain_omath_count") == 0)
    if not acceptable:
        raise RuntimeError(f"Restored candidate failed official content audit; inspect {temporary}")
    os.rename(temporary, output)
    result = {
        "schema": "axmath-exact-duplicate-restore/v1",
        "status": "ready_for_gpt_review",
        "source": str(original),
        "source_sha256": manifest["original_source_sha256"],
        "reduced_ready": str(ready_path),
        "reduced_working": str(converted),
        "working_docx": str(output),
        "working_sha256": sha256(output),
        "source_formulas": manifest["source_formulas"],
        "unique_converted": manifest["unique_formulas"],
        "reused_duplicates": manifest["skipped_duplicates"],
        "copied_parts": cloned["copied_parts"],
        "raw_conversion_output": raw,
        "clone_seconds": clone_seconds,
        "restore_seconds": time.perf_counter() - started,
        "audit": {key: audit.get(key) for key in ("formula_count_state", "paragraph_count_equal", "nonmath_text_exact", "nonmath_text_contract_exact", "candidate_non_axmath_ole_count", "candidate_nonmain_omath_count")},
        "mandatory_next_step": "Run axmath.py diagnose on frozen source and working_docx, then the unchanged strict visual review; this is not final acceptance.",
    }
    output.with_name(output.name + ".cache-report.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return {key: value for key, value in result.items() if key != "audit"}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    analyze_parser = commands.add_parser("analyze", help="Count byte-identical OfficeMath repeats without changing files")
    analyze_parser.add_argument("--input", required=True)
    prepare_parser = commands.add_parser("prepare", help="Freeze source and create a reduced DOCX with bookmark placeholders")
    prepare_parser.add_argument("--input", required=True)
    prepare_parser.add_argument("--outdir", required=True)
    prepare_parser.add_argument("--min-duplicates", type=int, default=64)
    restore_parser = commands.add_parser("restore", help="Clone genuine AxMath OLE after the official reduced conversion")
    restore_parser.add_argument("--manifest", required=True)
    restore_parser.add_argument("--ready", required=True)
    restore_parser.add_argument("--output", required=True)
    restore_parser.add_argument("--raw", action="store_true", help="Use raw conversion output for a benchmark; normal acceptance uses normalized working_docx")
    restore_parser.add_argument("--numbering-scope", choices=("line-start", "anywhere"), default="line-start")
    args = parser.parse_args()
    if args.command == "analyze":
        source = Path(args.input).resolve(strict=True)
        result = {"source": str(source), "source_sha256": sha256(source), **repeated_formulas(source)}
        result.pop("duplicate_template_indices")
    elif args.command == "prepare":
        result = prepare(Path(args.input), Path(args.outdir), min_duplicates=args.min_duplicates)
    else:
        result = restore(Path(args.manifest), Path(args.ready), Path(args.output), raw=args.raw, numbering_scope=args.numbering_scope)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
