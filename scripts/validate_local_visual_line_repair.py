from __future__ import annotations

import argparse
import json
from pathlib import Path

from validate_local_layout_repair import validate


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--baseline", required=True)
    ap.add_argument("--candidate", required=True)
    ap.add_argument("--map", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    repair_map = Path(args.map).resolve()
    data = json.loads(repair_map.read_text(encoding="utf-8-sig"))
    source = Path(str(data.get("source") or "")).resolve()
    if not source.is_file():
        raise FileNotFoundError(
            "repair map does not point to an accessible frozen source"
        )

    report = validate(
        source,
        Path(args.baseline),
        Path(args.candidate),
        repair_map,
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
