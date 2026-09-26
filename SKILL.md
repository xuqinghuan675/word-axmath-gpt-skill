---
name: word-axmath-conversion
description: Convert Microsoft Word OfficeMath to real editable AxMath, then review and repair layout/inline-display defects against the source document.
---

# Word → AxMath GPT Skill

## Use this first

When handling Word/AxMath tasks, read this file before acting.

Goal: keep the original DOCX untouched, create an adjacent work area, convert OfficeMath to genuine editable `Equation.AxMath` objects, then review/repair only the working copy.

## Fast path

1. **Preflight** the source DOCX: valid file, OfficeMath/AxMath counts, lock state, existing Word process.
2. **Create an adjacent workspace** next to the source.
3. **Convert once** with `python scripts/one_click_convert.py --input "<file-or-folder>"`.
4. Wait for `READY_FOR_GPT_REVIEW.json`.
5. **GPT review/repair** the working DOCX against the frozen source.
6. Final gate: run `strict_final_compare.py`, fresh-open/export both source and final, generate source-vs-final images for **every page**, verify centered formulas remain centered, then visually inspect every page.

Never overwrite the source. Never close/kill a Word process that this task did not create.

## Repair routing

Classify first; do not try every repair in sequence.

- **Source is inline, AxMath became giant/display or broke a same-line group**
  → `repair_axmath_inline_roundtrip.ps1`
  → AxMath `AMSAM2TeX` → force single `$...$` → `AMSTeX2AM`.

- **Remaining non-inline internal metric problem**
  → `rebuild_axmath_baselines.ps1` / `ConvertAMERebuild`.

- **Only the external Word OLE box is wrong**
  → `calibrate_axmath_boxes.py`, moving width/height + `dxaOrig/dyaOrig` + `w:position` coherently.

- **Preview-only problem**
  → repair/split the preview relationship; do not rewrite the OLE.

- **Semantic formula content is wrong**
  → rebuild from the source formula; never choose a donor only because its size looks similar.

For a source formula proven inline, **inline roundtrip is the first repair**. Do not waste time first on repeated rebuilds, arbitrary width thresholds, shell-only resizing, or preview swapping.

## Review rules

- `snapshot_docx.py` is the source-reference evidence collector.
- `formula_geometry_audit.py` compares source OfficeMath and working AxMath by ordinal and geometry.
- Repair in grouped passes, preferably descending ordinals.
- Re-resolve current ordinal/text anchors after edits; Word Range coordinates can drift.
- Formula crop pixel differences are triage only, not proof.
- Rendered source + surrounding text beat raw OMML tag names for inline/display intent.
- `Exactly` line spacing can mimic formula clipping.
- Shared preview targets are risky; do not overwrite shared WMF media unless all formulas are semantically identical.
- Full-document rendering is reserved for the final gate.

## Strict final source-vs-final gate

Run:

```powershell
python scripts\strict_final_compare.py --source "<source.docx>" --final "<final.docx>" --outdir "<compare-output>"
```

The tool fresh-renders both documents, creates a side-by-side image for **every page**, and checks source-centered formulas structurally.

GPT must inspect **every** generated source-vs-final page. This is not optional and must not be replaced by crop sampling or structural checks alone.

## Done means

- source SHA-256 unchanged;
- working DOCX opens normally;
- residual OfficeMath = 0;
- genuine `Equation.AxMath` count matches the source formula count;
- unrelated text and paragraph structure preserved;
- no unresolved source-reference geometry/same-line blockers;
- paragraph alignment/spacing preserved;
- **every source-centered formula remains centered**;
- `center_alignment_breaks = 0` and `center_position_breaks = 0`;
- no unresolved OLE/preview relationship conflict;
- no task-created Word/AxMath process remains;
- fresh reopen + fresh PDF for both source and final;
- GPT visually inspects every side-by-side page and finds no unexplained change in formula size, baseline, wrapping, line/page breaks, spacing, indentation, or centering.

`structural_pass=true` is required but does **not** replace the visual review.

Ambiguous exceptions are `needs_human_review`; never guess.
