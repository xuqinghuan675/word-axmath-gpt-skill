# Changelog

## Unreleased (2026-11-12 AxMath formula structure gate)

- Add a source-first formula-layout contract: scripted `lim` operators must use explicit `\\limits`, derivative prime order is preserved, and source structure (fractions, radicals, matrices/equation arrays, delimiters, scripts and relation operators) is checked against the complete AxMath internal-TeX export.
- Wire the preflight into `one_click_convert.py`, the mandatory final gate into `strict_final_compare.py`, and add a frozen-source-only `limit-layout` donor-map path. Repairs from a damaged AxMath object's own export are rejected.
- Add regression coverage for lower/upper/combined `lim` scripts, `limsup`/`liminf`, prime order, structural omissions and frozen-source donor provenance.

## Unreleased (2026-10-21 maintenance)

- Publish an explicit MIT license for repository-owned code. Microsoft Word, AxMath and third-party Python packages retain their own licenses.
- Add a contribution guide, sanitized bug-report template and a reproducible maintainer/evidence guide.
- Document that static CI checks do not substitute for licensed Word/AxMath conversion and per-page visual acceptance.
- Reject nonpositive duplicate thresholds and nonfinite or negative center-tolerance overrides at the command-line boundary, with regression tests.

This entry describes maintenance work, not a new user-facing conversion release or a claim of external adoption.
