---
name: word-axmath-conversion
description: Convert Microsoft Word OfficeMath to real editable AxMath with frozen-source evidence, deterministic repair routing, and strict visual acceptance.
---

# Word → AxMath GPT Skill

## 0. Read this before doing anything

For every Word / OfficeMath / OMML / AxMath task, this file is the execution authority after the user's current instruction.

**A new chat must read this file before it runs Word, AxMath, repair scripts, geometry experiments, or edits a formal DOCX.** Do not reconstruct the procedure from memory, old chats, handoff summaries, or a previous assistant's explanation.

The invariant is:

> frozen source → one official conversion → deterministic diagnosis → only the prescribed repair class → re-diagnose → one final full visual gate

Never overwrite the source. Never mutate the user's formal finished copy while investigating a generic bug. Never close or kill a Word process that this task did not prove it created.

## 1. Mandatory entry protocol

1. Read this `SKILL.md`.
2. Run preflight / inspect-only:

```powershell
python scripts\one_click_convert.py --input "<source.docx>" --inspect-only
```

3. Fix missing environment dependencies **before** a long conversion. Preflight checks Python modules, PowerShell, and the AxMath template.
4. Run the official conversion once:

```powershell
python scripts\one_click_convert.py --input "<source.docx>"
```

5. Open `READY_FOR_GPT_REVIEW.json`. Use its **frozen source** and **working DOCX**, not the live original.
6. Before any repair, run:

```powershell
python scripts\diagnose_after_conversion.py --source "<frozen-source.docx>" --working "<working.docx>" --outdir "<review-dir>"
```

7. Read `NEXT_ACTION.json` and follow its repair class. **Do not skip the current class and do not invent another repair route.**
   - Default diagnosis is **fast static + targeted Word proof**, not an all-formula COM sweep.
   - It uses DOCX/XML for counts, M1 structure, AxMath VML shell size, and conservative over-wide candidates; only suspicious source ordinals are opened in Word to prove multi-line intent.
   - Use `--deep-geometry` only when final/source-vs-working visual evidence exposes an ambiguous defect that fast triage cannot classify.
8. Every mutation writes a **new** working copy. After each grouped repair, run the diagnosis again.
9. Only when diagnosis says `ready_for_strict_final_compare`, run the full strict final gate once.

If a new chat has not performed steps 1 and 6, it is not authorized to improvise a repair.

## 2. Preflight contract

`one_click_convert.py --inspect-only` records:

- source readability, path and SHA-256;
- OfficeMath / AxMath / paragraph counts;
- source lock-file state and existing Word PIDs;
- static OMML structure, including raw `m:oMath` count, `m:oMathPara` groups, direct multi-sibling display groups, and the exact extra-node count that can explain a known AxMath batch-collapse signature;
- required environment: `psutil`, `pywin32`, `lxml`, `numpy`, `Pillow`, `PyMuPDF`, `omml2latex`, PowerShell, and `AxMath.dotm`.

Chinese paths must not fail because of a legacy console encoding. CLI entrypoints use UTF-8 output; a `cp1252` or console-print failure is an environment/output defect, not proof that the DOCX is corrupt.

A live source lock is advisory. If the original is readable, freeze it and bind the run to its SHA. All later source-semantic operations use the frozen source copy.

## 3. Phase A: official OfficeMath → AxMath conversion

The conversion loop uses AxMath's official `AMSMML2AM` macro.

**AxMath owns its per-call batch size.** On AxMath 2.7.0.58, production documents repeatedly showed about **64–66 formulas per completed call**. This is AxMath's stability behavior, not a tuning knob.

`WaitingConvert=1` is a completion signal, **not** a speed-control flag.

Never:

- force the hidden batch above AxMath's own size;
- bypass/reset completion semantics to make it faster;
- parallelize conversion of the same DOCX;
- remove per-batch `doc.Save()` recovery boundaries;
- close an unconfirmed AxMath dialog;
- infer completion from a control JSON placeholder such as `after_omath=0`.

Judge completion from the final conversion report, residual OfficeMath count, owned Word process state, and the saved working file.

Performance telemetry must retain batch durations, macro/save time, crash dumps, and owned Word PID. Optimize from measurements, not intuition.

## 4. Formula count is a state machine, not a yes/no check

`source_math_structure.py` inspects the frozen source before repair.

### State: exact

`candidate AxMath == source raw OfficeMath` and residual OfficeMath is zero. Proceed to geometry/semantic diagnosis.

### State M1: known multi-sibling collapse

A source `m:oMathPara` may contain several **direct sibling `m:oMath` children**. Word exposes those children individually through `OMaths`, while AxMath batch conversion can collapse the whole display group into one OLE.

If:

> source raw count − candidate AxMath count == frozen source multi-sibling extra-node count

then the gap is a known `M1_MULTISIBLING_OMATHPARA` signature, not unexplained formula loss.

Do **not** declare the missing objects randomly lost, tune width/height, chase page count, or run geometry calibration before M1 is fixed.

Build the frozen-source map:

```powershell
python scripts\build_source_tex_map.py multisibling --source "<frozen-source.docx>" --working "<working.docx>" --out "<M1-map.json>"
```

Repair to a new copy:

```powershell
powershell -File scripts\repair_multisibling_groups.ps1 -InputDocx "<working.docx>" -OutputDocx "<M1-fixed.docx>" -MapPath "<M1-map.json>"
```

M1 is the **only** structural class in which one collapsed AxMath is intentionally replaced by several AxMath objects, because the frozen source itself contains several raw OfficeMath children. The script verifies the exact expected count increase, OfficeMath=0, paragraph stability, and map/input SHA binding.

Then run `diagnose_after_conversion.py` again.

### State: unexplained loss / excess / residual OfficeMath

Stop mutation. Do not call it M1 unless the frozen structure proves the exact signature. Formula identity must be resolved before geometry/layout repair.

## 5. Mandatory repair classes after counts are exact

`diagnose_after_conversion.py` is the dispatcher. `formula_geometry_audit.py` is diagnostic evidence, not an auto-mutator.

### Class E — semantic/internal AxMath content is wrong

Examples:

- complex source formula became a tiny AxMath shell;
- double-clicking AxMath shows a short or unrelated formula;
- AxMath→TeX export is empty or semantically collapsed;
- source contains prime/derivative markers and roundtrip is risky.

The source of truth is always the **frozen OfficeMath at the same ordinal**.

Production route:

1. Export frozen-source Word LaTeX when useful with `export_source_word_latex.ps1`.
2. Treat COM success as insufficient. Word can emit non-empty malformed linear math containing control characters, placeholder glyphs such as `▒`, `〖`, `〗`, or broken constructs.
3. `build_source_tex_map.py` validates Word output and can fall back to **exact source OMML → LaTeX** via `omml2latex`.
4. GPT reviews the source-derived TeX contract.
5. Apply only that approved map with `repair_axmath_from_approved_tex.ps1`.

Do not use a damaged AxMath object's own exported TeX as the semantic donor.

Prime/derivative semantics are a **mandatory semantic route, independent of layout**. `formula_geometry_audit.py` reports every frozen-source formula containing `'`, `′`, `″`, `‴`, `⁗` or equivalent prime-like markers in `prime_semantic_candidates`; do not wait for a geometry warning. Before any source-semantic donor is approved, run `normalize_axmath_tex.py`. The AxMath 2.7.0.58 contract has been verified both from its parser-token table and real Word+AxMath round trips: **first derivative = `\prime`, second = `''`, third = `'''`**. Examples: `F′(x) → F\prime(x)`, `F″(x) → F''(x)`, `F‴(ξ) → F'''(ξ)`, `y′² → {y\prime}^{2}`. Raw Unicode `‴` is source evidence only: direct `AMSTeX2AM` input round-trips as `?` on this version. `repair_axmath_from_approved_tex.ps1`, M1 repair, and M2-B visual-line repair all hard-reject noncanonical prime literals and re-export prime-bearing donors through real AxMath before writing them into the working DOCX.

### M2 — source visual wrap is lost and AxMath overflows the right boundary

This is the long-formula bug from production:

- one source OfficeMath object is rendered by Word across multiple visual lines;
- AxMath conversion produces one fixed-width OLE;
- that OLE crosses the measured paragraph/page text-right boundary.

Detection is evidence-based:

- source snapshot proves the formula spans visual lines;
- working snapshot measures a reliable single-column text boundary;
- the AxMath right edge actually crosses that measured boundary.

**No arbitrary `width > N pt` threshold may trigger M2.** Table cells and multi-column sections stay review-only unless their real bounds are measured.

Export the frozen source's actual visual lines:

```powershell
powershell -File scripts\export_source_visual_lines.ps1 -SourceDocx "<frozen-source.docx>" -Ordinals "<M2 ordinals>" -OutputJson "<visual-lines.json>"
```

The extractor:

- does not load AxMath while source `OMath.Range` objects are alive;
- uses a separate task-owned Word process per ordinal;
- resolves visual-line boundaries from Word's real page/Y layout using bounded binary search instead of one COM call per character;
- validates resolved intervals so internal fraction/superscript vertical coordinates are not mistaken for line breaks;
- saves each line as a source OfficeMath fragment;
- tries Word LaTeX but keeps the fragment for exact OMML fallback;
- verifies frozen-source SHA before/after.

Build the aligned one-object repair map:

```powershell
python scripts\build_source_tex_map.py visual-wrap --visual-report "<visual-lines.json>" --working "<working.docx>" --out "<M2-map.json>"
```

Apply it:

```powershell
powershell -File scripts\repair_axmath_from_approved_tex.ps1 -InputDocx "<working.docx>" -OutputDocx "<M2-fixed.docx>" -MapPath "<M2-map.json>"
```

The map combines source-derived lines into one `aligned` TeX donor. `AMSTeX2AM` then creates **one real `Equation.AxMath` object**.

**M2-A invariant:** when the repair contract is an internal `aligned` multi-line AxMath, one source formula stays one AxMath object. **Do not generalize this to all M2 cases.**

This M2-A route is empirically verified on the production machine: a synthetic three-line `aligned` donor and a real four-line long-formula donor each converted into exactly one AxMath object.

#### M2-B — preserve frozen-source visual rows as multiple AxMath objects

Use M2-B when the source/formal reference clearly treats the derivation as separate visual mathematics rows, or when production evidence proves that exact row restoration is required. This is the route that fixed 24 long-formula paragraphs in the real `待排版5` document.

If the task starts from a known textual boundary, first run:

```powershell
python scripts\diagnose_local_layout.py --source "<frozen-source.docx>" --working "<working.docx>" --start-anchor "<unique anchor text>" --outdir "<local-review>"
```

It proves a unique source/working anchor, a stable paragraph offset, and only selects formulas whose frozen source is genuinely multi-line. Then follow the generated plan:

```powershell
python scripts\build_source_tex_map.py visual-line-split --visual-report "<visual-lines.json>" --working "<working.docx>" --paragraph-plan "<LOCAL_LAYOUT_PLAN.json>" --out "<M2-split-map.json>"
powershell -File scripts\repair_visual_line_split.ps1 -InputDocx "<working.docx>" -OutputDocx "<M2-split-fixed.docx>" -MapPath "<M2-split-map.json>"
python scripts\validate_local_visual_line_repair.py --baseline "<working.docx>" --candidate "<M2-split-fixed.docx>" --map "<M2-split-map.json>" --out "<M2-split-ledger.json>"
```

For M2-B, **one proven source visual row → one AxMath object**, separated by Word soft line breaks in the same paragraph. The split points come only from frozen-source visual lines; never from character count, guessed `=` positions, screen width, or arbitrary pixels.

M2-B intentionally increases AxMath object count. This is allowed only when the hash-bound `axmath-local-layout-repair-ledger/v1` proves the exact expected count increase, unchanged frozen prefix, unchanged non-math text, unchanged source SHA, and every target-row contract. Final comparison must receive that ledger via `--repair-ledger`.

The real M2 invariant is therefore: **preserve frozen-source semantics and visual structure**. Do not force `1 source OfficeMath = 1 AxMath` when a validated visual-row split is required; equally, do not split merely because an object is wide.

### Class A — source inline formula became giant/display or broke a same-line group

Use `repair_axmath_inline_roundtrip.ps1` only for visually confirmed low-semantic-risk culprit(s), maximum 12 reviewed ordinals per invocation.

Route: `AxMath → AMSAM2TeX → exactly one $...$ → AMSTeX2AM`.

A broken group is evidence that something is wrong, not permission to roundtrip every member. If `AMSAM2TeX` is empty, contains control characters, or the source has prime/derivative semantics, abandon Class A for that ordinal and use Class E.

### Class B — remaining non-inline internal metric problem

`rebuild_axmath_baselines.ps1` / `ConvertAMERebuild` is **probe-only**:

- maximum 3 explicit ordinals;
- re-audit immediately;
- if AxMath count changes, content collapses to a tiny shell, OLE hangs, or semantics/geometry regress, abandon Class B for that document and route affected formulas through Class E.

Never turn a failed probe into a batch.

### Class C — external Word OLE box only

Use `calibrate_axmath_boxes.py` only when evidence proves the internal formula is correct and only the external Word OLE shell is wrong.

It must update width/height, `dxaOrig/dyaOrig`, and `w:position` coherently. Plans are recommendations only; explicit reviewed ordinals are required. No global percentage sweeps.

### Class D — preview/cache differs from real OLE content

Symptom: Word displays formula A, but double-clicking AxMath opens formula B, or the preview appears stale while the internal OLE is correct.

Diagnose internal OLE content without mutating the working file:

```powershell
powershell -File scripts\inspect_axmath_tex.ps1 -InputDocx "<working.docx>" -Ordinals "<targets>" -OutputJson "<internal-tex.json>"
```

The script copies each selected AxMath to a temporary document and runs `AMSAM2TeX` **only on the copy**.

Decision:

- internal AxMath matches frozen source, Word preview is wrong → Class D preview/cache repair;
- internal AxMath is short/unrelated/wrong → Class E source-semantic rebuild;
- derivative/prime content → never trust the AxMath roundtrip as a semantic donor.

Do not overwrite shared WMF/preview relationships unless every object sharing that media is proven semantically identical.

## 6. Word/COM stability rules learned from production failures

These are mandatory:

- Use `$wordPid`, never `$PID`; PowerShell variable names are case-insensitive and `$PID` is reserved.
- Resolve Program Files x86 with `[Environment]::GetEnvironmentVariable('ProgramFiles(x86)')` or valid braced PowerShell environment-variable syntax; do not use the invalid unbraced form with parentheses.
- PowerShell COM collections such as Word `InlineShapes` are **not** trusted through `foreach`. Production testing on a 1521-shape document returned 1520 AxMath objects with explicit `.Count + .Item(i)`, but 0 through `foreach ($s in $doc.InlineShapes)`. All production AxMath enumerators therefore use indexed COM access; helper functions return the element stream normally and callers wrap with `@(...)` when they need an array. Never reintroduce `return ,@($arr)` + outer `@(...)`, which nests the whole collection as one item.
- PowerShell `Get-FileHash` can fail on a DOCX that is still readable through a shared handle. Repair/source-evidence scripts use a SHA-256 helper that opens the file read-only with `FileShare.ReadWrite|Delete`; integrity remains strict, but transient Office/OLE sharing does not become a false corruption signal.
- **Separate source-reading from AxMath mutation.** Do not hold a live source `OMath.Range` while loading/operating the AxMath add-in in another document. Production testing showed `PROPERTYGET` / RPC failures and invalidated ranges.
- On `0x800706BE` / RPC failure, discard the task-owned Word instance; do not keep using poisoned COM objects. Re-open the frozen source in a fresh owned Word process and retry only the failed target.
- `export_source_visual_lines.ps1` isolates by ordinal for this reason.
- Never kill all `WINWORD.EXE`; only the PID proven to belong to this run.
- Pre-existing user Word sessions are allowed and must remain untouched.
- `OwnedWord` must prove a **new/distinct WINWORD PID** before setting `Visible`, `DisplayAlerts`, opening a document, or calling `Quit()`. `DispatchEx` is not trusted by itself: Word may reuse a pre-existing headless `/Automation -Embedding` server. If distinct ownership cannot be proved, release the COM proxy and fail that attempt without mutating or quitting the pre-existing process.
- After a COM timeout, inspect owned process state and partial outputs before retrying. Do not repeat a deterministic failure with identical evidence.
- Use the frozen source for hashing/source-semantic operations. A live original may be temporarily locked in a way that blocks `Get-FileHash`; that is not permission to mutate or close the user's document.

## 7. Performance and tool-use rules

### Phase A

Do not spend engineering effort bypassing the observed 64–66 batch behavior unless independent reproducible evidence shows a newer AxMath version changed it.

### Repair iterations

Default diagnosis is deliberately **fast static + targeted** and must be used first:

```powershell
python scripts\diagnose_after_conversion.py --source "<frozen-source.docx>" --working "<working.docx>" --outdir "<review-dir>"
```

It does not scan every AxMath object through Word COM. It first uses static DOCX/XML evidence, then asks Word only about suspicious source ordinals. If no high-confidence repair defect is found, proceed to the strict final page-by-page visual gate.

For a genuinely ambiguous visual defect, explicitly request the expensive comprehensive pass:

```powershell
python scripts\diagnose_after_conversion.py --source "<frozen-source.docx>" --working "<working.docx>" --outdir "<review-dir>" --deep-geometry
```

`--deep-geometry` is the only route that should run the whole-document geometry audit. The lower-level geometry snapshot remains available as supporting evidence:

```powershell
python scripts\snapshot_docx.py "<frozen-source.docx>" "<geometry-dir>" --profile geometry
```

Geometry profile keeps COM formula inventory and SHA checks but skips full PDF/page rendering. Run complete source/final PDF + every-page images **once at final acceptance**, not after every repair.

### Named production tools

- static OMML structure → `source_math_structure.py`;
- content/count check → `audit_docx.py`;
- mandatory dispatcher → `diagnose_after_conversion.py`;
- source/working geometry → `snapshot_docx.py --profile geometry`;
- formula comparison → `formula_geometry_audit.py`;
- M1 map → `build_source_tex_map.py multisibling`;
- M1 repair → `repair_multisibling_groups.ps1`;
- source visual lines → `export_source_visual_lines.ps1`;
- M2 map → `build_source_tex_map.py visual-wrap`;
- M2/Class E one-object repair → `repair_axmath_from_approved_tex.ps1`;
- internal AxMath diagnostic → `inspect_axmath_tex.ps1`;
- final gate → `strict_final_compare.py` + `finalize_visual_review.py`.

Do not substitute ad-hoc scripts when a named production tool already covers the class.

## 8. Prohibited repair behavior

Unless the prescribed class has failed with reproducible evidence:

- no global width/height sweeps;
- no global `w:position` sweeps or forced zero;
- no arbitrary fixed-width thresholds;
- no one-by-one try-a-size-and-see loops;
- no shrinking OLE shells to force page count;
- no choosing a repair merely because page count matches;
- no preview swapping to fix internal semantics;
- no whole-group Class A roundtrip;
- no current-bad-AxMath-as-semantic-donor;
- no **unvalidated** one-source-long-formula → multiple-AxMath split; M2-B is allowed only from frozen-source visual lines with a hash-bound local-layout ledger;
- no experiments on the formal finished file;
- no repeated full-document render during repair iterations.

Page count is an acceptance observation, not an optimization target.

## 9. Evidence order

Use this authority order:

1. user's current explicit instruction;
2. current `SKILL.md` and named scripts;
3. current frozen-source/working-file evidence;
4. earlier reports only after validating input hashes/provenance;
5. old chat summaries or remembered explanations only as leads, never proof.

If detector evidence conflicts with an old report, regenerate it. If a detector fails, fix the detector before mutating the document.

## 10. Strict final source-vs-final gate

Only after the diagnosis queue is empty:

```powershell
python scripts\strict_final_compare.py --source "<frozen-source.docx>" --final "<final.docx>" --outdir "<compare-output>" --expected-source-sha256 "<frozen-source-sha256>"
```

This intentionally does not return final success. It fresh-opens both documents, verifies hashes/stability/content/formula counts, fresh-renders both, creates a source-vs-final image for every page, and writes `VISUAL_REVIEW_TEMPLATE.json`.

GPT must inspect **every** side-by-side page. Then fill the review manifest and run:

```powershell
python scripts\finalize_visual_review.py --report "<compare-output>\STRICT_FINAL_COMPARE.json" --review "<compare-output>\VISUAL_REVIEW.json"
```

Only `acceptance_pass=true` is formal completion.

Visual layout is the final acceptance authority. Geometry/centering/same-line diagnostics guide review; if page-level comparison is genuinely correct, a non-visible metric warning does not justify destructive repeated repair.

## 11. Done means

- original source SHA unchanged;
- all work happened on frozen source / new working copies;
- final DOCX opens normally;
- residual OfficeMath = 0;
- genuine `Equation.AxMath` count matches raw frozen-source OfficeMath count after structural repair;
- M1 structural merges are resolved or explicitly absent;
- no visually observable M2 right-overflow/wrap-loss remains;
- no semantic/internal mismatch remains;
- preview-only defects are isolated from semantic defects;
- paragraph count and non-math text are exact;
- unrelated text/style structure is preserved;
- centered formulas remain visually centered;
- fresh source/final render exists;
- every source-vs-final page was reviewed;
- `finalize_visual_review.py` reports `acceptance_pass=true`;
- no task-created Word/AxMath process remains.

Ambiguous cases are `needs_human_review`. Never guess, and never turn the formal document into a parameter-search sandbox.
