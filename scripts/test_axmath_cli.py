"""Safe stdlib-only tests for the local AxMath launcher (no Word COM invoked)."""
from __future__ import annotations

import contextlib
import io
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import axmath


class LocalCliTests(unittest.TestCase):
    def check_launch(self, argv, script, expected_tokens=()):
        with patch("axmath.subprocess.run") as runner:
            runner.return_value.returncode = 0
            self.assertEqual(axmath.main(argv), 0)
            runner.assert_called_once()
            call = runner.call_args
            command = call.args[0]
            self.assertEqual(command[0], sys.executable)
            self.assertEqual(Path(command[1]).name, script)
            self.assertEqual(call.kwargs["cwd"], str(axmath.ROOT))
            self.assertEqual(call.kwargs["env"]["PYTHONIOENCODING"], "utf-8")
            for token in expected_tokens:
                self.assertIn(token, command)
            return command

    def test_read_only_inspect(self):
        command = self.check_launch(
            ["inspect", "--input", r"C:\文档\原稿.docx"],
            "one_click_convert.py",
            ["--inspect-only", "--input", "--numbering-scope"],
        )
        self.assertNotIn("--resume-run", command)

    def test_convert_is_not_automatically_resumed(self):
        command = self.check_launch(
            ["convert", "--input", r"C:\文档\原稿.docx"],
            "one_click_convert.py",
            ["--input", "line-start"],
        )
        self.assertNotIn("--inspect-only", command)
        self.assertNotIn("--resume-run", command)

    def test_resume_keeps_original_checkpoint(self):
        command = self.check_launch(
            ["resume", "--input", r"C:\原稿.docx", "--resume-run", r"C:\run-2026"],
            "one_click_convert.py",
            ["--resume-run", "--input"],
        )
        self.assertNotIn("--inspect-only", command)

    def test_diagnosis_dispatch_options(self):
        command = self.check_launch(
            ["diagnose", "--source", r"C:\frozen.docx", "--working", r"C:\working.docx",
             "--outdir", r"C:\review", "--deep-geometry", "--numbering-scope", "anywhere"],
            "diagnose_after_conversion.py",
            ["--deep-geometry", "anywhere", "--source", "--working"],
        )
        self.assertNotIn("--inspect-only", command)

    def test_final_review_not_automatically_approved(self):
        self.check_launch(
            ["compare", "--source", r"C:\source.docx", "--final", r"C:\candidate.docx",
             "--outdir", r"C:\strict", "--expected-source-sha256", "a" * 64],
            "strict_final_compare.py",
            ["--expected-source-sha256", "a" * 64],
        )
        self.check_launch(
            ["finalize", "--report", r"C:\strict\STRICT_FINAL_COMPARE.json",
             "--review", r"C:\strict\VISUAL_REVIEW.json"],
            "finalize_visual_review.py",
            ["--review", "--report"],
        )

    def test_status_does_not_fabricate_pass(self):
        with tempfile.TemporaryDirectory() as tmp:
            report = Path(tmp) / "STRICT_FINAL_COMPARE.json"
            report.write_text(
                json.dumps({"schema": "strict/v2", "acceptance_pass": False,
                            "visual_review_required": True}), encoding="utf-8"
            )
            stream = io.StringIO()
            with patch("axmath.subprocess.run") as runner, contextlib.redirect_stdout(stream):
                self.assertEqual(axmath.main(["status", "--path", tmp]), 0)
                runner.assert_not_called()
            summary = json.loads(stream.getvalue())
            self.assertTrue(summary["read_only"])
            self.assertFalse(summary["reports"][0]["acceptance_pass"])
            self.assertIn("NOT ACCEPTED", summary["reports"][0]["acceptance_note"])

    def test_status_missing_does_not_touch_word(self):
        with tempfile.TemporaryDirectory() as tmp:
            with patch("axmath.subprocess.run") as runner, contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(axmath.main(["status", "--path", tmp]), 2)
                runner.assert_not_called()

    def test_path_is_resolved_before_launch(self):
        with tempfile.TemporaryDirectory() as tmp:
            original = os.getcwd()
            try:
                os.chdir(tmp)
                command = self.check_launch(
                    ["inspect", "--input", "例子.docx"],
                    "one_click_convert.py",
                )
                self.assertEqual(command[command.index("--input") + 1],
                                 str(Path(tmp) / "例子.docx"))
            finally:
                os.chdir(original)


if __name__ == "__main__":
    unittest.main()
