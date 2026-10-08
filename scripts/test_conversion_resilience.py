"""Pure safety checks for per-batch continuation and conversion ownership."""
from __future__ import annotations

import argparse
import datetime
import json
import tempfile
from pathlib import Path

import run_skill
from axmath_batch_dialog_watcher import seconds_since_state, verified_owned_word
from execution_lock import conversion_lock


def main():
    scripts = Path(__file__).resolve().parent
    ps = (scripts / "convert_officemath_to_axmath.ps1").read_text(encoding="utf-8")
    assert "while($doc.OMaths.Count -gt 0){" in ps
    assert "$batch -lt 64" not in ps
    assert "Write-ConversionCheckpoint" in ps
    assert "-ResumeOutput" in (scripts / "run_skill.py").read_text(encoding="utf-8")

    now = datetime.datetime.now(datetime.timezone.utc)
    assert (seconds_since_state({"updated_at": now.isoformat()}) or 0) < 5
    assert verified_owned_word({"owned_word_pid": 0}) is None

    with tempfile.TemporaryDirectory() as folder:
        root = Path(folder)
        source = root / "source.docx"
        output = root / "working.docx"
        source.write_bytes(b"SOURCE")
        output.write_bytes(b"SAVED_PARTIAL")
        args = argparse.Namespace(
            input=str(source), output=str(output),
            overwrite_output=False, resume=True, watchdog_seconds=1800,
        )
        try:
            run_skill._execute(args)
        except FileNotFoundError as exc:
            assert "checkpoint" in str(exc)
        else:
            raise AssertionError("Resumed a working copy without a checkpoint.")
        checkpoint = Path(str(output) + ".conversion-checkpoint.json")
        checkpoint.write_text(json.dumps({
            "schema": "axmath-conversion-checkpoint/v1",
            "source_sha256": "invalid",
            "working_sha256": "invalid",
            "source_path": str(source),
            "working_path": str(output),
        }), encoding="utf-8")
        try:
            run_skill._execute(args)
        except ValueError as exc:
            assert "checkpoint" in str(exc)
        else:
            raise AssertionError("Resumed with untrusted checkpoint hashes.")
        with conversion_lock():
            try:
                with conversion_lock():
                    raise AssertionError("exclusive mutex double-acquired")
            except RuntimeError:
                pass
    print("CONVERSION_RESILIENCE_TEST_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
