# Advanced AxMath repair reference

This is exception-only operational detail extracted verbatim from the 2026-10-08 Skill baseline. Current `SKILL.md` has priority; consult this file only if the prescribed diagnostic class or Word/COM failure requires it. Never follow an old handoff instead of current frozen-source evidence.

## Legacy OLE handling

For `Equation.3`, `Equation.DSMT4`, MathType variants, or unknown OLE, use `embedded_object_inventory.py` to locate the story, paragraph and ProgID. Preserve the original source and hash. Review each object and rebuild verified mathematical content as editable OfficeMath on a separate DOCX, then inspect the new copy before starting AxMath conversion. Do not automatically confirm the Word conversion dialog for all formulas. Verify prime orders and indices manually; unknown OLE is not necessarily a formula. Keep the mapping and both source versions.

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

## 12. Word 页面级视觉属性恢复

重新生成 DOCX 或 AxMath 转换链不得默认丢弃原稿页面级视觉属性。

若发现：
- 页面背景色丢失；
- 页面边框丢失；
- 页面装饰属性与原稿不一致；

不得手工逐页修改，应从冻结源文档提取并恢复 Word XML 页面属性。

工具：

```powershell
python scripts\restore_word_page_layout.py --reference "<source.docx>" --target "<working.docx>" --output "<fixed.docx>"
```

原则：
- 只恢复页面级属性；
- 不修改正文 XML；
- 不触碰 AxMath OLE、OMML、公式内容；
- 不改变公式数量；
- 不重新分页。

默认恢复：
- `w:background`
- `w:pgBorders`

完成后仍需运行最终视觉检查。

## 13. 2026-10-08 conversion/resume and speed hardening

The AxMath plugin still owns its per-macro batch size (roughly 64–66 formulas on
2.7.0.58). **The Skill no longer caps total macro invocations at 64**. A new
invocation continues while OfficeMath remains, but each call must strictly
reduce the number of OfficeMath objects; unchanged or increased counts stop
the task. Never patch AxMath internals or bypass WaitingConvert.

Every completed batch saves the working DOCX and then atomically records
`<working.docx>.conversion-checkpoint.json`, including the SHA-256 of both
frozen source and saved working DOCX, remaining OfficeMath count, and last
batch ordinal. Before resuming, both hashes, paths and residual count must
match. A mismatching checkpoint must **fail closed**; never automatically
ignore it or overwrite saved progress.

Resume a previous incomplete, single-document run (use the original input
path and the existing run directory):

```powershell
python scripts\one_click_convert.py --input "<original-source.docx>" --resume-run "<existing-run-YYYYMMDD-HHMMSS>" --numbering-scope line-start
```

Or continue only the saved conversion working copy:

```powershell
python scripts\run_skill.py --input "<frozen-source.docx>" --output "<partial-working.docx>" --resume
```

`run_skill.py` serializes this Skill's conversions with an exclusive
per-user mutex, because AxMath shares HKCU completion state. Its dialog
watcher closes **only confirmed** completed AxMath dialogs. The default
30-minute per-macro/save-phase watchdog checks both owned WINWORD PID and
process creation time before terminating a stalled **task-owned** process,
never a pre-existing user Word. For unusual environments it can be changed
via `--watchdog-seconds` (minimum 60). The watcher never treats a timeout
or a missing PID proof as successful conversion.

### Fast diagnosis and semantic checks

`diagnose_after_conversion.py` uses lightweight static OMML inspection for
prime/derivative and collection-operator risks. Prime-containing source
formulas are not silently declared good because width/page geometry matches.
The diagnostic writes `SOURCE_SEMANTIC_RISKS.json` and
`PRIME_REVIEW_TEMPLATE.json`. When prime risks exist, evidence must come
from actual source-to-AxMath semantic verification; an unverified template
cannot be marked passed. Give the hash-bound approved review via:

```powershell
python scripts\diagnose_after_conversion.py --source "<frozen.docx>" --working "<working.docx>" --outdir "<review>" --prime-review "<verified-prime-review.json>"
```

Set-operator candidates are **visual review priorities, not permission to
guess replacement glyphs**. Geometry anomalies alone do not justify formula
rewriting. Any page background or border mismatch is diagnosed as a separate
page-style repair, and the final strict acceptance gate now rejects such
mismatches even if the formula count and non-math text match.

### Optional punctuation scope

The safe default remains `--numbering-scope line-start`. If explicitly
requested, `--numbering-scope anywhere` changes every literal Arabic digit
or parenthesized Arabic-number label immediately followed by `、`, even
mid-paragraph; this intentionally also changes `第1、2项` to `第1.2项`.
It is **not suitable for normal book proofreading by default**. Use the
same explicit numbering scope in `one_click_convert.py`,
`diagnose_after_conversion.py`, and `strict_final_compare.py`.

### Performance rules and validation

- `snapshot_docx.py` opens and decodes each rendered page once when
  cropping all formula images on that page; it never reopens the page
  for every formula.
- Final page-by-page source/final visual review stays mandatory. Do not
  cache a previous Word render as "fresh final evidence" merely because
  the DOCX hash matches; Word layout can depend on fonts/printer/runtime.
- `restore_word_page_layout.py` writes a new output only, inserts
  `w:pgBorders` in schema order, verifies the ZIP/page style and preserves
  OLE/embedding bytes. It fails closed if section counts differ or a
  page decoration references media relationships not copied.
- Pure Python regression checks: `test_conversion_resilience.py`,
  `test_page_style_and_semantics.py`, `test_crop_performance.py`,
  `test_numbering_punctuation.py`, `test_prime_normalization.py`.
  CI does not replace a real Word+AxMath smoke conversion.

### 新增：独立显示公式居中偏移修复（必须视觉证实）

在十三月实机合成测试中，`m:oMath` 独占段落时，Word 的原生数学布局可能自动居中，但转换后的 `Equation.AxMath` 会按段落的左对齐位置显示；两者即使 Word COM `x_pt` 相同，PDF 实际像素也可能不同。现在静态检测会列出 `STANDALONE_DISPLAY_ALIGNMENT_VISUAL` 风险供逐页重点核查；**仅在源稿/完成稿视觉对比证实且明确指定公式序号后**，可执行最小修复：

```powershell
python scripts\repair_standalone_display_alignment.py --source "<冻结源稿>" --working "<待修复工作稿>" --output "<新完成稿.docx>" --ordinals "1,2" --expected-source-sha256 "<源SHA256>" --expected-working-sha256 "<工作稿SHA256>"
```

此修复只修改确认的公式段落 `w:pPr/w:jc=center`，不碰 AxMath/OLE 嵌入内容，也不强制给所有公式居中。修复后**必须重新逐页视觉验收**。十三月实机已验证 3 个合成独立公式从错误左对齐恢复与原稿一致，并通过 `acceptance_pass=true`。实机验证只覆盖小样本，不代表 4000 公式全书已跑完。

另外，编号无替换时使用字节级复制，避免将包含数千个 OLE 的 DOCX 整体重新压缩；上层会复用已计算的 frozen source 和 working 静态分析结果，减少重复 ZIP/XML 读取。
