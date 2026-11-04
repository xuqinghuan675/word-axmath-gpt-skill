# Contributing / 参与贡献

Thanks for helping improve **word-axmath-gpt-skill**. This repository contains an open-source automation workflow; **Microsoft Word and AxMath are separate proprietary applications**. You must obtain and license those applications yourself. They are not distributed here.

## Development and checks

Use Windows and Python 3.11 for the closest match to CI.

```powershell
python -m pip install -r requirements.txt
python scripts/repo_selfcheck.py
python scripts/test_prime_normalization.py
python scripts/test_numbering_punctuation.py
python scripts/test_internal_content_gate.py
python -m unittest discover -s scripts -p "test_axmath_cli.py" -v
```

See [repository-static-check.yml](.github/workflows/repository-static-check.yml) for the rest of the no-Office regression suite. It runs on GitHub Actions Windows runners. Python and document-structure checks **do not prove** visual correctness of a converted document.

## Reporting problems

Open an issue with OS, Python/Word/AxMath versions, the command used, expected and actual behavior, and the deterministic repair class from `NEXT_ACTION.json` if available. Use a small synthetic or redacted DOCX. **Never upload confidential customer documents, product keys, personal information, commercial templates, or unredacted logs.**

## Pull-request contract

1. Preserve the frozen source and default non-overwrite behavior. Never change a file in place merely to make a test pass.
2. Add a minimal reproducible regression test for logic or parsing changes.
3. Keep `SKILL.md`, `README.md` and any affected `docs/` files consistent.
4. Run the no-Office regression suite and PowerShell parsing checks, then report which manual Word/AxMath visual checks were *actually* performed.
5. Do not assert `acceptance_pass=true` without the documented internal-content and per-page visual review.

Changes should be small and independently verifiable. Ownership of commercial AxMath/Word binaries and third-party Python dependencies does not transfer under this repository's license.
