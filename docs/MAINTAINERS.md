# Maintenance and evidence guide / 维护与验收证据

This document helps users and external maintainers reproduce the claims of the **word-axmath-gpt-skill** repository. It is not evidence of external user counts or a claim that conversion is fully automatic in every document.

## What is maintained

The repository owner **xuqinghuan675** maintains a Windows-oriented local workflow that converts Microsoft Word native OfficeMath/OMML into editable AxMath OLE and then audits formula identity, structure, content and layout. The maintained code consists of the CLI `axmath.py`, scripts for conversion and fault classification, hash-bound frozen-source checkpoints, guarded local repairs, post-conversion diagnostics, and regression tests.

The central contract is deliberately conservative:

1. Freeze the source rather than editing it in place.
2. Perform the official conversion once; resume from a checkpoint instead of repeating it.
3. Diagnose via `NEXT_ACTION.json` and only apply the documented repair for the actual failure class.
4. Review internal formula content and all affected pages before final acceptance.

Read [SKILL.md](../SKILL.md) before working on an actual document. CLI commands and constraints are documented in [README.md](../README.md).

## Reproducible verification

Windows, Python 3.11 and `requirements.txt` dependencies are used for static checks. To inspect the test suite without launching Office:

```powershell
python -m pip install -r requirements.txt
python scripts/repo_selfcheck.py
python -m compileall -q scripts
python scripts/test_legacy_ole_gate.py
python scripts/test_internal_content_gate.py
python -m unittest discover -s scripts -p "test_axmath_cli.py" -v
```

The complete no-Office CI checks, including PowerShell script parsing, are in [repository-static-check.yml](../.github/workflows/repository-static-check.yml). **A passing CI run is not a successful real Word/AxMath conversion or visual acceptance.** Production use requires a licensed Word and AxMath installation in an interactive Windows session and a page-by-page review.

## Honest limitations

- Word and AxMath proprietary applications, their license keys, and third-party plug-ins are **not** included or licensed by this repository.
- Real conversion timing and visual fidelity depend on the input DOCX, Office/AxMath versions and Windows state. Synthetic checks do not validate customer files.
- No audited third-party user counts, download figures, broad community adoption or external endorsements are asserted.
- Distinguish runnable tests from manually inspected outputs; report negative cases and unresolved classes honestly.

## Maintainership workflow

Issues should contain a reproducible synthetic example (never private source documents), environment information, and a diagnostic report with secret values removed. For PRs, preserve frozen-source hashing, non-overwrite guarantees, and explicit human/agent-reviewed acceptance gates. See [CONTRIBUTING.md](../CONTRIBUTING.md) and the repository's Git history and CI run records for verifiable activity.
