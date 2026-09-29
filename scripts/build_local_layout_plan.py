from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from build_source_tex_map import static_axmath_inventory
from source_math_structure import analyze_source_structure, classify_count_state

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _source_ordinal_locations(source: Path) -> dict[int, int]:
    with zipfile.ZipFile(source) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    paragraphs = root.findall(".//w:p", NS)
    result: dict[int, int] = {}
    ordinal = 0
    for pi, paragraph in enumerate(paragraphs, 1):
        for _ in paragraph.findall(".//m:oMath", NS):
            ordinal += 1
            result[ordinal] = pi
    return result


def _paragraph_texts(docx: Path) -> list[str]:
    with zipfile.ZipFile(docx) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    return [
        "".join(t.text or "" for t in p.findall(".//w:t", NS))
        for p in root.findall(".//w:p", NS)
    ]


def _find_unique(texts: list[str], anchor: str, label: str) -> int:
    hits = [i for i, text in enumerate(texts, 1) if anchor in text]
    if len(hits) != 1:
        raise ValueError(
            f"{label} anchor must be unique; found {len(hits)} matches for {anchor!r}: {hits[:12]}"
        )
    return hits[0]


def build_plan(
    source: Path,
    working: Path,
    ordinals: list[int],
    *,
    start_anchor: str | None = None,
) -> dict:
    source = source.resolve()
    working = working.resolve()
    source_sha = sha256_file(source)
    working_sha = sha256_file(working)
    structure = analyze_source_structure(source)
    inv = static_axmath_inventory(working)
    source_locations = _source_ordinal_locations(source)

    rows = []
    source_anchor = None
    working_anchor = None
    paragraph_offset = None

    if start_anchor:
        source_anchor = _find_unique(_paragraph_texts(source), start_anchor, "source")
        working_anchor = _find_unique(_paragraph_texts(working), start_anchor, "working")
        paragraph_offset = working_anchor - source_anchor

    count_state = classify_count_state(
        structure, inv["axmath_count"], inv["omath_count"]
    )
    working_by_ordinal = {
        int(row["ordinal"]): row for row in inv["locations"]
    }

    for ordinal in sorted({int(x) for x in ordinals}):
        source_paragraph = source_locations.get(ordinal)
        if source_paragraph is None:
            raise ValueError(f"source ordinal not found: {ordinal}")

        if paragraph_offset is not None:
            working_paragraph = source_paragraph + paragraph_offset
            if working_paragraph <= working_anchor:
                raise ValueError(
                    f"target source ordinal {ordinal} maps before/at frozen anchor"
                )
            candidates = [
                x for x in inv["locations"]
                if int(x["paragraph_index"]) == working_paragraph
            ]
            if len(candidates) != 1:
                raise ValueError(
                    f"mapped working paragraph {working_paragraph} must contain exactly "
                    f"one AxMath before visual-line split; found {len(candidates)}"
                )
        else:
            if count_state["state"] != "exact":
                raise ValueError(
                    "without a unique start anchor, local plan requires exact formula "
                    f"identity/count; got {count_state['state']}"
                )
            location = working_by_ordinal.get(ordinal)
            if location is None:
                raise ValueError(f"working ordinal not found: {ordinal}")
            working_paragraph = int(location["paragraph_index"])
            candidates = [location]

        rows.append(
            {
                "source_ordinal": ordinal,
                "source_paragraph_index": source_paragraph,
                "working_paragraph_index": working_paragraph,
                "working_current_axmath_count": len(candidates),
            }
        )

    if not rows:
        raise ValueError("local layout plan has no rows")

    return {
        "schema": "axmath-local-layout-plan/v1",
        "source": str(source),
        "source_sha256": source_sha,
        "working": str(working),
        "working_sha256": working_sha,
        "source_raw_omath_count": structure["raw_omath_count"],
        "working_axmath_count": inv["axmath_count"],
        "working_omath_count": inv["omath_count"],
        "working_paragraph_count": inv["paragraph_count"],
        "formula_count_state": count_state,
        "start_anchor": start_anchor,
        "source_anchor_paragraph": source_anchor,
        "working_anchor_paragraph": working_anchor,
        "paragraph_offset": paragraph_offset,
        "first_repair_paragraph": min(x["working_paragraph_index"] for x in rows),
        "rows": rows,
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--working", required=True)
    ap.add_argument("--ordinals", required=True)
    ap.add_argument("--start-anchor")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    ordinals = [
        int(x.strip()) for x in args.ordinals.split(",") if x.strip()
    ]
    payload = build_plan(
        Path(args.source),
        Path(args.working),
        ordinals,
        start_anchor=args.start_anchor,
    )
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "row_count": len(payload["rows"]),
                "first_repair_paragraph": payload["first_repair_paragraph"],
                "paragraph_offset": payload["paragraph_offset"],
                "out": str(out.resolve()),
            },
            ensure_ascii=True,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
