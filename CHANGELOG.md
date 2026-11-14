# Changelog

## Unreleased (2026-11-14 numbering and Skill routing)

- Route numbered math exercises and non-collapsible headings to one canonical, first-pass contract with Word/WPS acceptance.
- Detect Heading style identity even with outline level 9, plus basedOn/next/link referrals; add focused regression fixtures.
- Shorten the main SKILL and README; move the original conversion state machine intact to docs/AXMATH_CONVERSION_WORKFLOW.md, and preserve previous README as historical reference.
- Keep the existing MinerU GPT/API and Codex browser-first/fallback behavior.


## Unreleased (2026-11-13 MinerU GPT direct / Codex browser fallback)

- Add a single MinerU Precision API entry for local PDFs, including resumable batch polling, raw ZIP preservation, safe extraction and DOCX handoff into the existing AxMath pipeline.
- Keep the MinerU API token in the ignored local secrets directory; GPT can call the API directly on the local machine, while Codex uses its own logged-in conversation sidebar first and switches to the local API only for long web queues or a stalled front end.
- Add a problem-to-file routing table and document the official API request/status flow.

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
