from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock
import json
import zipfile
from types import SimpleNamespace
import pymupdf

sys.path.insert(0, str(Path(__file__).resolve().parent))
import mineru_api


class MinerUApiTests(unittest.TestCase):
    def test_batch_payload_places_ocr_on_file(self):
        payload = mineru_api.build_payload(Path("math.pdf"), "a" * 32, "vlm")
        self.assertTrue(payload["files"][0]["is_ocr"])
        self.assertNotIn("is_ocr", payload)
        self.assertEqual(payload["extra_formats"], ["docx"])
        self.assertTrue(payload["enable_formula"])
        self.assertTrue(payload["enable_table"])

    def test_safe_extract_and_full_docx_selection(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            archive = root / "good.zip"
            output = root / "mineru_output"
            with zipfile.ZipFile(archive, "w") as zf:
                zf.writestr("full.docx", b"docx")
                zf.writestr("images/page1.png", b"image")
            mineru_api.safe_extract_zip(archive, output)
            self.assertEqual(mineru_api.choose_docx(output).name, "full.docx")


    def test_ambiguous_submit_receipt_never_resubmits(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            pdf_path = root / "sample.pdf"
            with pymupdf.open() as doc:
                doc.new_page()
                doc.save(pdf_path)
            workspace = root / "workspace"
            workspace.mkdir()
            source_hash = mineru_api.sha256_file(pdf_path)
            mineru_api.write_json(workspace / "mineru_job.json", {
                "source_sha256": source_hash,
                "state": "submission_unknown",
                "batch_id": None,
            })
            args = SimpleNamespace(
                input=str(pdf_path),
                outdir=str(workspace),
                model_version="vlm",
                poll_interval=1,
                max_wait_seconds=1,
            )
            with self.assertRaisesRegex(mineru_api.MinerUError, "请求结果不确定"):
                mineru_api.run_parse(args)

    def test_token_is_read_from_local_file(self):
        with tempfile.TemporaryDirectory() as temp:
            token_file = Path(temp) / "mineru_token.txt"
            token_file.write_text("test-token-only", encoding="utf-8")
            with mock.patch.object(mineru_api, "TOKEN_PATH", token_file):
                self.assertEqual(mineru_api.load_token(), "test-token-only")

    def test_token_value_is_not_in_batch_payload(self):
        payload = mineru_api.build_payload(Path("math.pdf"), "b" * 32, "vlm")
        self.assertNotIn("test-token-only", json.dumps(payload))

    def test_zip_slip_is_rejected(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            archive = root / "bad.zip"
            with zipfile.ZipFile(archive, "w") as zf:
                zf.writestr("../escape.txt", b"bad")
            with self.assertRaises(mineru_api.MinerUError):
                mineru_api.safe_extract_zip(archive, root / "bad_output")


if __name__ == "__main__":
    unittest.main(verbosity=2)
