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
5. **GPT review/repair** the working DOCX against the frozen source. For intermediate geometry/source evidence, prefer `snapshot_docx.py ... --profile geometry`; this keeps the SHA-bound COM formula inventory but skips full PDF/page rendering, formula crops, and preview extraction that the geometry audit does not consume.
6. Final gate: run `strict_final_compare.py`, fresh-open/export both source and final, generate source-vs-final images for **every page**, verify centered formulas remain centered, then visually inspect every page. Final strict compare intentionally uses the default **full** snapshot profile; never substitute geometry-only evidence for final visual acceptance.

Never overwrite the source. Never close/kill a Word process that this task did not create.

## GPT execution ownership

Once the user has specified the task goal and authorized the operation, GPT owns the end-to-end working-copy workflow. Do not hand foreground/background Word/AxMath work, Save/Save As handling, task-owned dialogs, retries, detection, repair, page-by-page review, or task-process cleanup back to the user merely because UI interaction is required.

A pre-existing user Word process is not a reason to abort. Record the pre-existing Word PIDs, create and verify a distinct task-owned Word session, and manipulate/close/kill only the task-owned process. If isolated ownership cannot be proven, fail that automation attempt without touching the pre-existing Word session and continue by a safe task-owned route.

GPT may handle task-owned dialogs and temporary foreground interaction when the workflow genuinely requires them. Do not steal focus from or operate an unrelated user Word window. Working copies must be writable; if a frozen source is read-only, clear the read-only attribute only on the copied working file, never on the source.

A Word lock file is advisory, not an automatic reason to hand the task back to the user. If the on-disk source remains readable, freeze the source copy and proceed under the existing source-hash checks; if the source changes during the run, fail acceptance rather than overwriting or closing the user's document.

## Phase A performance invariants

The OfficeMath → AxMath Phase A loop deliberately calls AxMath's official `AMSMML2AM` macro repeatedly. **AxMath owns the per-call batch size.** On AxMath 2.7.0.58, both the historical 768-formula test and the 1335-formula production run observed about **64–66 formulas per completed macro call**; AxMath itself displays the reason: batching is used to avoid Word becoming unresponsive.

`WaitingConvert=1` is a confirmed AxMath batch-completion signal used by the dialog watcher. It is **not a performance knob**. Do not force a larger hidden batch, bypass/reset the completion semantics, close unconfirmed AxMath dialogs, parallelize conversion of the same DOCX, or remove the per-batch `doc.Save()` recovery boundary merely to improve elapsed time.

The measured 1335-formula production run spent about **6718.6 / 6785.0 seconds (99.02%) inside `AMSMML2AM` itself**, about **5.03 macro-seconds per formula**. Python orchestration, watcher polling, persistence, integrity checks, and process cleanup are therefore not the dominant Phase A bottleneck. Conversion reports must retain per-batch macro/save/crash telemetry so future optimization is evidence-based.

Full-document snapshot/render work is expensive too. During repair iterations, prefer geometry audit plus targeted formula/page evidence. Reserve the complete fresh source/final snapshot + every-page strict compare for the final gate, unless a current diagnostic genuinely requires a new full render.

`snapshot_docx.py --profile geometry` is the intermediate fast path: it still opens the real DOCX read-only, computes the Word COM formula inventory/paragraph geometry and verifies the DOCX hash before/after, but it does not export/render every page or create per-formula crop/preview media. The default profile remains `full`; strict final acceptance must keep that default.

## Repair routing

Classify first; do not try every repair in sequence.

- **Source is inline, AxMath became giant/display or broke a same-line group**
  → `repair_axmath_inline_roundtrip.ps1`
  → AxMath `AMSAM2TeX` → force single `$...$` → `AMSTeX2AM`.
  → If `AMSAM2TeX` returns empty/control-character output for an ordinal, do not blindly retry the same roundtrip. Reclassify from fresh evidence and prefer source-semantic Class E when the source formula is available.
  → A broken same-line group is evidence that the line is wrong, **not** permission to roundtrip every member. GPT must identify the visually/semantically suspicious culprit ordinal(s); one invocation is limited to at most 12 reviewed targets, followed by a fresh audit.
  → Source formulas containing prime/derivative markers (`'`, `′`, `″`, `‴`, `⁗` and equivalent prime-like glyphs) are a known semantic-risk class for `AMSAM2TeX`: a syntactically valid result may silently collapse `f′(x)` to `f`. This risk is **independent of layout**. `formula_geometry_audit.py` must report every such source formula in `prime_semantic_candidates`, even when same-line geometry is perfect; route those ordinals through frozen-source Class E rather than Class A.

- **Remaining non-inline internal metric problem**
  → `rebuild_axmath_baselines.ps1` / `ConvertAMERebuild` is **probe-only**.
  → Use at most **3 explicit ordinals** in one probe, then re-audit before doing anything else.
  → If AxMath count changes, the rebuilt object collapses (for example to a tiny `7.5×16.5` shell/content), the OLE hangs, or semantics/geometry regress, stop this route for the document. Never turn a failed probe into a larger batch.

- **Only the external Word OLE box is wrong**
  → `calibrate_axmath_boxes.py`, moving width/height + `dxaOrig/dyaOrig` + `w:position` coherently.
  → Geometry audit rows are recommendations only. Visually confirm the affected formula(s), then pass their ordinals explicitly; do not auto-apply a stale/statistical plan.

- **Preview-only problem**
  → repair/split the preview relationship; do not rewrite the OLE.

- **Semantic formula content is wrong**
  → rebuild from the **frozen source formula at the same ordinal**; never choose a donor only because its size looks similar.
  → Preferred source-semantic route when direct OMML→AxMath produces damaged content: export the frozen OfficeMath formula through Word's LaTeX representation, normalize Word-specific LaTeX constructs, then create the donor with AxMath `AMSTeX2AM`.
  → For every `prime_semantic_candidates` ordinal, run `scripts/normalize_axmath_tex.py` on the exported Word LaTeX before approval. The verified AxMath 2.7.0.58 parser contract is **first prime = `\prime`, second prime = `''`, third prime = `'''`**. Normalize by derivative order, for example `F′(x) → F\prime(x)`, `F″(x) → F''(x)`, `F‴(ξ) → F'''(ξ)`, and `y′² → {y\prime}^{2}`.
  → This contract comes from two independent local checks: AxMath.exe's own parser-token table contains `''' `, `'' `, and `\prime `, and real Word+AxMath `AMSTeX2AM → AMSAM2TeX` round trips preserve prime orders 1/2/3. Raw Unicode `‴` is explicitly unsafe as donor TeX on this version because it round-trips as `?`.
  → `repair_axmath_from_approved_tex.ps1` is a hard gate: it rejects Unicode/typographic prime literals, a lone ASCII apostrophe for first derivative, prime orders above the verified 1..3 range, Unicode superscript literals, and ungrouped prime+exponent forms. For prime-bearing donors it also re-exports a copy through AxMath and verifies that the prime-order sequence survives before writing into the working document.
  → If the rendered source formula is genuinely multi-line, preserve source-derived line breaks with an `aligned`-style TeX donor instead of forcing the formula onto one line or shrinking its OLE box.
  → `formula_geometry_audit.py` reports `semantic_rebuild_candidates` for non-trivial source expressions trapped in tiny AxMath shells (for example `7.5×16.5`). This is a review queue, never an auto-apply list; legitimate single-glyph narrow formulas must remain untouched.
  → Production execution is a GPT-owned source-semantic path: `export_source_word_latex.ps1` exports reviewed source ordinals from the frozen DOCX without modifying it; `normalize_axmath_tex.py` creates deterministic AxMath-safe TeX candidates; GPT checks the source meaning and writes the final GPT-approved `{ordinal, tex}` map; `repair_axmath_from_approved_tex.ps1` applies only that approved map to a **new** working copy and verifies AxMath/OfficeMath/paragraph counts.
  → Do not let either script guess formula semantics. Word LaTeX export is evidence; the GPT-approved map is the semantic contract.

For a source formula proven inline, **inline roundtrip is the first repair only for a visually confirmed low-semantic-risk culprit**. A broken same-line group alone is not enough; derivative/prime or suspiciously collapsed formulas go directly to frozen-source Class E. A visually identical prime formula is still not permission to skip this route: `prime_semantic_candidates` are semantic routing requirements, not geometry warnings. Do not waste time on whole-group roundtrips, repeated rebuilds, arbitrary width thresholds, shell-only resizing, or preview swapping.

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
2. For confirmed **Class A** source-inline defects, inspect the broken same-line group and repair only the culprit ordinal(s), never the entire group by membership alone. Use `repair_axmath_inline_roundtrip.ps1` only for reviewed low-semantic-risk targets (max 12 per invocation); route derivative/prime or suspicious roundtrip cases directly to source-semantic Class E.
3. Re-audit. Only defects still proven to be **Class B** may use a **local probe of at most 3 ordinals** with rebuild/baseline repair. If that probe changes AxMath count, collapses content/geometry, hangs OLE, or otherwise regresses, abandon Class B for that document and route the affected formula(s) through source-semantic Class E instead.
4. Re-audit. Only defects proven to be **Class C** external Word OLE-box mismatches may use `calibrate_axmath_boxes.py`.
5. Class D preview-only and Class E semantic repairs remain isolated to their own evidence.
6. When no **visually observed** repair defect remains, run the strict final gate and inspect every page. Diagnostic geometry warnings alone do not justify another mutation.

Do **not** leave this state machine merely because the document page count differs from the source. Page count is an acceptance result, not an optimization target.

Unless the current class's prescribed repair has failed with reproducible evidence, the following are prohibited on a formal document:

- global width/height percentage sweeps;
- global `w:position` percentage sweeps or forcing `w:position=0`;
- arbitrary per-formula width thresholds or repeated one-by-one size probes;
- shrinking OLE shells merely to force the source page count;
- preview crop/swap experiments used to solve internal AxMath metrics;
- choosing a repair because it makes the page count look right.
- forcing/bypassing AxMath's observed 64–66-formula batch behavior, treating `WaitingConvert` as a speed-control flag, or removing the per-batch save/watcher safety boundary without independent reproducible evidence that AxMath no longer needs it.

If the prescribed route fails, stop mutation first. Re-read this Skill and the current evidence, then research the generic mechanism (official Word/OLE documentation and public implementation experience) before adding a new repair. Validate the generic fix on the formal document, then update the Skill; never turn the formal document into a parameter-search sandbox.

A detector failure is not permission to start geometry experiments. Fix the detector, regenerate source evidence, rebuild the repair plan, then resume at the correct class.


## Review rules

- `snapshot_docx.py` is the source-reference evidence collector.
- `formula_geometry_audit.py` compares source OfficeMath and working AxMath by ordinal and geometry. Its output is **diagnostic/triage evidence**, not an automatic reason to keep modifying a document that already matches visually.
- Repair in grouped passes, preferably descending ordinals.
- Re-resolve current ordinal/text anchors after edits; Word Range coordinates can drift.
- Formula crop pixel differences are triage only, not proof.
- Rendered source + surrounding text beat raw OMML tag names for inline/display intent.
- `Exactly` line spacing can mimic formula clipping.
- If page flow drifts while source/final paragraph formatting is otherwise identical, compare rendered ink bounds and source baseline spacing against the AxMath OLE shell height. A transparent/tall OLE shell can inflate Word auto line spacing even when the formula glyphs are already the right size; do not globally scale formula content to solve that mechanism.
- Shared preview targets are risky; do not overwrite shared WMF media unless all formulas are semantically identical.
- Full-document rendering is reserved for the final gate.

## Strict final source-vs-final gate

Run the evidence pass first, using the frozen pre-conversion source SHA-256:

```powershell
python scripts\strict_final_compare.py --source "<source.docx>" --final "<final.docx>" --outdir "<compare-output>" --expected-source-sha256 "<frozen-source-sha256>"
```

This command intentionally does **not** return success yet. It fresh-renders both documents, verifies hard content invariants, creates a side-by-side image for **every page**, and writes `VISUAL_REVIEW_TEMPLATE.json`.

After GPT has directly inspected **every** side-by-side page, fill a copy of that template from the actual visual review and finalize it:

```powershell
python scripts\finalize_visual_review.py --report "<compare-output>\STRICT_FINAL_COMPARE.json" --review "<compare-output>\VISUAL_REVIEW.json"
```

Only `acceptance_pass=true` from the finalizer is a completed final gate. The review is hash-bound to the source, final DOCX, and every reviewed page image, so a stale visual pass cannot survive later file changes.

**Visual layout is the final acceptance authority.** Structural geometry, same-line, centering, crop, shell, and baseline diagnostics exist to direct attention to suspicious areas; they must not trigger further mutation by themselves when the corresponding source-vs-final pages are visually correct.

GPT must inspect **every** generated source-vs-final page. This is not optional and must not be replaced by crop sampling or structural checks alone. If a structural warning remains but the relevant page is visually indistinguishable in layout and formula placement, record the warning as non-blocking rather than repeatedly resizing/rebuilding the formula.

## Done means

- source SHA-256 unchanged;
- working DOCX opens normally;
- residual OfficeMath = 0;
- genuine `Equation.AxMath` count matches the source formula count;
- hard content gate passes: frozen source SHA matches, source/final stay unchanged during compare, paragraph count matches, and non-math text is exact;
- unrelated text and paragraph structure preserved;
- no **visually observable** unresolved source-reference layout defect; geometry/same-line warnings may remain when page-level visual inspection confirms no layout difference;
- paragraph alignment/spacing is visually preserved;
- **every source-centered formula remains visually centered**;
- `center_alignment_breaks` / `center_position_breaks` are diagnostic flags, not standalone blockers when the corresponding pages pass visual inspection;
- no OLE/preview relationship conflict that causes a visible or semantic defect;
- no task-created Word/AxMath process remains;
- fresh reopen + fresh PDF for both source and final;
- GPT visually inspects every side-by-side page and finds no unexplained change in formula size, baseline, wrapping, line/page breaks, spacing, indentation, or centering.

`structural_pass=true` is preferred, but it is **not required when the only remaining failures are diagnostic geometry/layout warnings and every affected page passes direct visual comparison**. Final completion is instead `hard_content_pass=true` plus hash-bound per-page visual review with `acceptance_pass=true`. Visual correctness must never be sacrificed merely to make a metric reach zero.

Ambiguous exceptions are `needs_human_review`; never guess.
