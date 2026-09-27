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

## Authority and evidence order

For every formal Word/AxMath task, use this authority order:

1. **The user's current explicit instruction.**
2. **The current authoritative Skill file and its named scripts.**
3. **Current, freshly collected evidence from the real source/working files and tool output.**
4. Earlier generated reports only after their inputs, detector version, and provenance are revalidated.
5. Previous chat messages, handoff summaries, assistant explanations, remembered conclusions, and speculative diagnoses are **not evidence**.

Never promote a previous assistant's statement such as "this bug was verified", "this formula is inline", "the detector is wrong", or "this repair worked" into a code/Skill change without independently checking the underlying evidence in the current task.

When the machine or source document is unavailable, do **not** modify executable repair/detector code based only on old chat claims. Documentation may be hardened from the user's explicit process requirements, but implementation changes must wait for reproducible evidence or independent public/official corroboration.

If a prior assistant's narrative conflicts with the current Skill, follow the current Skill. If current tool evidence conflicts with an old report, refresh the report. Never let a handoff summary silently override the user's written procedure.

## Mandatory repair state machine

The repair routing above is a **hard execution order**, not a menu of experiments.

1. Freeze source evidence and build one repair plan.
2. Repair confirmed **Class A** source-inline defects with `repair_axmath_inline_roundtrip.ps1`.
3. Re-audit. Only defects still proven to be **Class B** may use rebuild/baseline repair.
4. Re-audit. Only defects proven to be **Class C** external Word OLE-box mismatches may use `calibrate_axmath_boxes.py`.
5. Class D preview-only and Class E semantic repairs remain isolated to their own evidence.
6. When blockers reach zero, run the strict final gate and inspect every page.

Do **not** leave this state machine merely because the document page count differs from the source. Page count is an acceptance result, not an optimization target.

Unless the current class's prescribed repair has failed with reproducible evidence, the following are prohibited on a formal document:

- global width/height percentage sweeps;
- global `w:position` percentage sweeps or forcing `w:position=0`;
- arbitrary per-formula width thresholds or repeated one-by-one size probes;
- shrinking OLE shells merely to force the source page count;
- preview crop/swap experiments used to solve internal AxMath metrics;
- choosing a repair because it makes the page count look right.

If the prescribed route fails, stop mutation first. Re-read this Skill and the current evidence, then research the generic mechanism (official Word/OLE documentation and public implementation experience) before adding a new repair. Validate the generic fix on the formal document, then update the Skill; never turn the formal document into a parameter-search sandbox.

A detector failure is not permission to start geometry experiments. Fix the detector, regenerate source evidence, rebuild the repair plan, then resume at the correct class.


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
