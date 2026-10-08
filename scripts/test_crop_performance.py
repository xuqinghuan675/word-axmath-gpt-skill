"""Regression: full visual gate opens/decompresses a page once, not once per formula."""
from __future__ import annotations

import tempfile
from pathlib import Path
from unittest import mock

from PIL import Image

from snapshot_docx import _crop_formula_pages


def main():
    with tempfile.TemporaryDirectory() as folder:
        root = Path(folder)
        page = root / "page.png"
        Image.new("RGB", (480, 640), "white").save(page)
        pages = [{
            "page": 1, "path": str(page),
            "width_px": 480, "height_px": 640,
            "width_pt": 240, "height_pt": 320,
        }]
        formulas = [
            {"formula_id": f"omath-{i:04d}", "page": 1,
             "x_pt": float(10 + (i % 6) * 30), "y_pt": float(10 + i * 2),
             "visual_width_pt": 18.0, "font_size_pt": 12}
            for i in range(100)
        ]
        original = Image.open
        with mock.patch.object(Image, "open", wraps=original) as intercepted:
            _crop_formula_pages({"omath": formulas, "axmath": []}, pages, root / "crops")
            assert intercepted.call_count == 1, (
                f"Expected exactly one image open for 100 formulas, got {intercepted.call_count}"
            )
        assert all(Path(x["crop_png"]).is_file() for x in formulas)
        assert len(list((root / "crops").glob("*.png"))) == 100
    print("CROP_ONE_DECODE_PER_PAGE_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
