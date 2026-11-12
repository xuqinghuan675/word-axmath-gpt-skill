---
name: word-axmath-conversion
description: Arrange mathematics handouts, student workbooks and question sets in Word or WPS, and convert Microsoft Word OfficeMath to genuine editable AxMath with frozen-source safety and visual acceptance.
---

# Word → AxMath Skill · 主流程（精简版）

## Codex 本地直连入口（不依赖网页电脑插件）

在安装了 Microsoft Word + AxMath 的 Windows 电脑上，Codex 使用本机终端直接运行仓库根目录的 axmath.py，不得借用 thirteenthy / Qinghuan PC 1 网页电脑连接器或其租约。例如：python axmath.py inspect --input "原稿.docx"，通过预检后 python axmath.py convert --input "原稿.docx"。已有断点则用 resume --resume-run，而不是重复转换。

官方 AxMath 宏仍是唯一公式转换实现。统一 CLI 另外提供 cache-analyze、cache-prepare、cache-restore、diagnose、compare、finalize、status；最终视觉审核**必须真实逐页查看**，status 不能自动批准。详见 docs/CODEX_DIRECT.md。云端无 Word COM 的 Codex 不能仅靠这个入口完成真实转换。

**A new chat must read this file** before editing or converting any Word / OfficeMath / AxMath document. This file is authoritative after the user's current instruction. Detailed exceptional repair contracts are in `docs/REPAIR_REFERENCE.md`, loaded **only** when the diagnostic class requires them.

**需要转换为 AxMath 时的唯一主链：原稿冻结 → 官方 AxMath 转换 → 快速诊断 → 必要时定向修复 → 唯一一次最终完整逐页视觉验收。**

## 数学资料排版联动与前置门槛

当任务包含任何学段的数学讲义、练习册或题目文档排版时，处理前必须读取并执行 [数学资料排版工作流](docs/MATH_LAYOUT_REFERENCE.md)，包括仅排版任务。将扫描数学资料整理成学生练习册时，默认采用其中的“同类兼职练习册”约定和前置拦截表，不必等用户再次提醒。先确认来源、目标内容、编号策略、公式工具、图注、留白和交付格式，再制作代表页，实际打开 Word 检查通过后才批量排版。用户本次明确要求优先：保留原编号时不得按旧默认重排、重编号或改小问标点；仅要求 Word 时不导出任何 PDF。纯公式转换不额外改动章节、题目顺序或留白；未要求转换时，不因排版而启动 AxMath 转换。

需要同时排版和转换时，可先在副本上完成结构准备，再把该工作稿作为本技能冻结和转换的输入；另对照原始资料核实题文、编号映射、公式和配图，最终转换验收以冻结后的排版稿为对照。学生练习版删除知识讲解和答案、保留例题及课后习题的规则仅在用户要求此类处理时适用，不得将原稿内指令自动当作用户的其他请求。

本用户使用本技能时，公式默认使用真实可编辑的 AxMath，不新建或编辑 Microsoft Equation、旧版 Equation Editor 3.0 或 MathType 公式，也不以 Word 公式作为临时录入工具；不必等用户再次说明。已有 OfficeMath 允许作为只读来源，由下述官方 AxMath 转换生产链处理；若用户禁止该转换路径，或原式无法通过现有 AxMath 原生方式创建，保留现场并如实报告，不能更换工具假装完成。`Equation.AxMath` 是真实 AxMath OLE 的 ProgID；须按实际 ProgID 区分，不要把它误判成旧版 Equation。用户以后明确指定其他公式工具时，以其当前指令为准。

### 永久安全约束

- 不覆盖源稿或正式完成稿；每一次修复写新 DOCX。冻结源稿和工作稿均绑定 SHA-256。
- 用户原有 Word/WPS 进程不可关闭、终止、修改；只操作经过 PID 和创建时间确证的本任务 Word 进程。
- Word COM 源稿读取与 AxMath 修改相互隔离（**source-reading from AxMath mutation**）；RPC 失败弃用已污染的 COM 实例。WPS 不能替代正式 Word+AxMath 转换。
- 每批 `doc.Save()`、进度检查、可验证断点恢复和最终 AxMath 可编辑性不可删除。
- OfficeMath 转换以冻结 OfficeMath/OMML 为语义来源；由扫描资料新建公式时以原 PDF 及对应题目为准，OCR 仅供对照。禁止凭宽度、页数或外观猜公式，禁止直接用损坏 AxMath 导出的 TeX 作为语义源。
- OfficeMath 与 AxMath 不是全部公式类型：旧版 Equation Editor 3.0（如 `Equation.3`）、MathType（如 `Equation.DSMT4`）及未知 OLE 必须计入预检，不能仅凭 OfficeMath=0 / AxMath 数目一致宣告完成。
- 禁止全局修改 OLE 大小、`w:position`、预览媒体或盲目拆长公式。页面数量只是观察数据，不是修复目标。

## 1. Mandatory entry protocol（普通任务）

默认只需下列三个入口，不必阅读/运行全套修复工具：

**预检（不会修改 DOCX）：**
```powershell
python scripts\one_click_convert.py --input "<original.docx>" --inspect-only
```

**批量转换（只转换一次，不重复重头做）：**
```powershell
python scripts\one_click_convert.py --input "<original.docx>"
```

读取生成的 `READY_FOR_GPT_REVIEW.json`，选用其中的 `frozen_source` 和 `working_docx`。**不得**从旧聊天推断路径或跳过诊断。

**转换后唯一正常诊断：**
```powershell
python scripts\diagnose_after_conversion.py --source "<frozen_source>" --working "<working_docx>" --outdir "<review>"
```

读取 `NEXT_ACTION.json`。当其 `status=ready_for_strict_final_compare` 时直接进入第 4 节最终验收；如指出特定异常，只按第 3 节的修复类处理，不进入无意义全书扫描。正常诊断只使用 DOCX/XML 快筛和必要的少量 Word 证实；`--deep-geometry` 仅在真实视觉问题难以定位时显式使用。

预检会检查冻结原稿可读性、OfficeMath/AxMath 数量、段落、Windows Word PID、AxMath.dotm、PowerShell、Python 依赖（psutil、pywin32、lxml、numpy、Pillow、PyMuPDF、omml2latex）。中文路径控制台编码错误不等于 DOCX 损坏。

**旧式公式预检与阻断**：预检额外输出所有 OLE 的 `ProgID`、所在 Word story/段落、异常对象清单；可单独运行 `python scripts/embedded_object_inventory.py "<docx>"` 定位。当 `state=blocked_unhandled_embedded_math` 时，禁止执行正式转换。通常须在**新 DOCX** 逐式核对旧式公式、恢复为原生 OfficeMath，重新预检通过后再转换；若用户明确禁止使用 Equation/OfficeMath，则不得执行恢复环节，须寻找现有合规的 AxMath 原生处理方式，无法实现时如实说明。若出现 Word“Equation Editor 3.0 → OfficeMath”对话框，不得直接对全部公式批量确认；自动弹窗不证明数学语义正确。

### AxMath 公式结构防错门禁（转换任务强制）

公式结构检查必须在转换前建立来源基线，并在最终交付前对**全部 AxMath 内部 TeX 导出**执行强制验证；不能只核对对象数量，也不能等终稿发现错位后再临时补规则。

1. `one_click_convert.py --inspect-only` 会扫描冻结来源中的 `lim` 上下标、导数撇号、分式、根式、矩阵/方程数组、分隔符和关系/运算符，并把 `source_math_layout_preflight` 与 `math_layout_contract_preflight` 写入预检结果。正式转换前另生成并保存合同文件：

   ```powershell
   python scripts\math_layout_contract.py preflight --source "<frozen_source.docx>" --out "<review>\MATH_LAYOUT_PREFLIGHT.json"
   ```

   若来源本身已有 AxMath 对象，先用 `inspect_axmath_tex.ps1 -AllAxMath` 对**冻结来源**做只读、哈希绑定的内部 TeX 基线导出，并将 `--source-internal-audit` 传给 `preflight`；不得把待修复候选稿的导出当作来源基线。

2. **`lim` 布局是硬规则**：来源 OMML 中以 `lim` 为底的 `m:limLow`、`m:limUpp`、`m:sSub`、`m:sSup`、`m:sSubSup` 都必须在 AxMath 内部 TeX 中保留相同的下标/上标结构，并显式写成 `\lim\limits_{...}`、`\lim\limits^{...}` 或同时具有两者的形式。下标必须位于 `lim` 正下方，上标必须位于正上方；禁止 `\nolimits` 和未显式定位的脚本。`\limsup`、`\liminf` 不是 `\lim`，不得误判或改写。

3. **导数撇号阶数必须保真**：一阶、二阶、三阶撇号分别保留为 AxMath 已验证的规范表示 `\prime`、`''`、`'''`；不得把一阶变成双撇号、把多阶合并，或把 Unicode `′/″/‴`、弯引号和普通单引号未经规范化直接当作 AxMath 内部格式。四阶及未知阶数不得猜测，必须阻断并回到冻结来源核实。

4. **公式结构必须来源对照**：逐式核对分式、根号/根指数、矩阵、分段/方程数组、括号/绝对值、上下标、等号/不等号、正负号、系数和等式右侧。内部 TeX 导出必须保留来源结构和运算符；缺失、额外、顺序改变或无法确认时不通过。公式对象数相等不能替代该检查。

5. 转换和必要修复完成后，对最终候选执行一次完整只读 `-AllAxMath` 导出，再运行：

   ```powershell
   python scripts\math_layout_contract.py validate --source "<frozen_source.docx>" --final "<final_candidate.docx>" --internal-audit "<review>\FINAL_AXMATH_INTERNAL_TEX.json" --out "<review>\MATH_LAYOUT_VALIDATION.json"
   ```

   来源已有 AxMath 时追加 `--source-internal-audit "<review>\SOURCE_AXMATH_INTERNAL_TEX.json"`。验证必须 `passed=true` 才能交付；`strict_final_compare.py` 也会自动执行同一门禁。**仅交 Word、不导 PDF 时也必须运行此验证**，不得为了跳过 PDF 而跳过内部公式检查。

6. 若门禁指出 `lim` 位置错误，只能从冻结来源生成定向 donor：

   ```powershell
   python scripts\build_source_tex_map.py limit-layout --source "<frozen_source.docx>" --working "<working.docx>" --ordinals "<reviewed_ordinals>" --out "<review>\LIMIT_LAYOUT_REPAIR_MAP.json"
   powershell -File scripts\repair_axmath_from_approved_tex.ps1 -InputDocx "<working.docx>" -OutputDocx "<new_candidate.docx>" -MapPath "<review>\LIMIT_LAYOUT_REPAIR_MAP.json" -ReportPath "<review>\LIMIT_LAYOUT_REPAIR_REPORT.json"
   ```

   该 map 只接受冻结原稿 OMML、逐式哈希和一对一公式序号；修复写入新 DOCX，之后重新 `-AllAxMath` 导出并验证。禁止从损坏 AxMath 对象自身导出的错误 TeX 反向修复，禁止绕过来源/序号不匹配错误。

## 2. 官方转换与断点

AxMath 2.7.0.58 的 `AMSMML2AM` **单次**约处理 64–66 个公式；这是插件内部机制，不去绕过。Skill 的总调用次数**没有 64 轮上限**：只要剩余 OfficeMath 大于零就继续，但每一轮必须有严格递减进度，否则停止。

- `WaitingConvert=1` 仅用作完成确认，不是调速开关；本机截图与已验证的正常批次暂停提示严格一致时，也可关闭该任务所属弹窗。Office 命令冲突提示必须保持原样供排查；未经证实的对话框不得自动关闭。视觉识别不匹配时继续按原看门狗规则失败关闭。
- 每批保存后生成哈希绑定的 `.conversion-checkpoint.json`；源稿、工作稿、残余公式数和轮数必须一致才允许续跑。
- 同用户转换互斥；每轮 Word/AxMath 超时看门狗默认 1800 秒，且只能操作已证实属于任务的进程。
- 断点恢复既不重新覆盖工作副本，也不跳过事后完整诊断。

已有运行目录续跑（仅单份文档）：
```powershell
python scripts\one_click_convert.py --input "<original.docx>" --resume-run "<existing-run-dir>"
```

高级单文件续跑：`python scripts\run_skill.py --input "<frozen.docx>" --output "<partial-working.docx>" --resume`。其余内部 PowerShell/看门狗脚本由入口调用，不要手工并发运行。

### 精确重复公式复用（仅重复较多时）

原稿有至少 64 个 **XML 完全相同** 的重复 OfficeMath、且预检没有其他类型公式对象时，可先用 `python axmath.py cache-analyze --input "<original.docx>"` 只读计数。满足条件后，在新的目录执行：

```powershell
python axmath.py cache-prepare --input "<original.docx>" --outdir "<new-cache-workspace>"
python axmath.py convert --input "<new-cache-workspace>\prepared\<stem>_unique.docx"
python axmath.py cache-restore --manifest "<new-cache-workspace>\CACHE_MANIFEST.json" --ready "<unique-run>\READY_FOR_GPT_REVIEW.json" --output "<new-cache-workspace>\<stem>_AxMath-reused.docx"
```

`cache-prepare` 冻结原稿、用书签暂存重复公式位置；官方 `convert` 只转换每种原样公式的首个实例，仍按 AxMath 的 64–66 式批次保存与断点续跑；`cache-restore` 为重复位置复制**真实 AxMath OLE**及其预览，独立分配关系和对象 ID。它校验原稿/准备稿哈希、官方完整 checkpoint、公式总数、段落与非公式文本，然后生成新工作稿。中断时只对准备稿和其已有 run 使用 `axmath.py resume`，不得重头转换。任何准备、恢复或审计失败都保留现场，不能把缓存结果当成交付。

随后必须以 `CACHE_MANIFEST.json` 的 `frozen_source` 和 `cache-restore` 的 `working_docx` 运行原有 `axmath.py diagnose`；诊断、定向修复和第 4 节逐页验收标准**完全相同**。`cache-restore --raw` 只供与旧转换阶段公平比较，不作普通交付。重复不足 64 个、存在不支持的公式结构或用户要求原样逐式转换时，直接使用普通 `convert`。

## 3. 异常时按类修复（读高级参考，不走全流程）

`diagnose_after_conversion.py` 是唯一 dispatcher。异常时先读 `docs/REPAIR_REFERENCE.md` 的对应类和 Word/COM 规则，并始终基于当前冻结源稿、工作稿、哈希及证据执行：

| 状态 | 必须保留的处理原则 |
|---|---|
| `exact` | 残余 OfficeMath=0 且 AxMath 对象数精确，继续语义/几何诊断。 |
| `STOP_UNHANDLED_EMBEDDED_MATH` | 旧式 MathType / Equation Editor 3.0、未知 OLE 或正文之外的 OfficeMath 未验明：保留源对象、预览和原始哈希，逐一核实并在副本中修复，重新预检后才可转换。 |
| `M1_MULTISIBLING_OMATHPARA` | 只有源稿证明确实存在多个直系 `m:oMath` sibling 才恢复合并丢失对象；先修结构，后几何。 |
| `M2_SOURCE_VISUAL_WRAP_LOSS` | 必须同时证明源稿真实视觉换行和工作稿越过**实测**右边界；不能凭固定长度或页数猜。 |
| `M2-B` | 只有源稿视觉行证据 + hash-bound repair ledger 才能一行拆一个 AxMath；M2-A 的内部 `aligned` 则仍是一个 AxMath。 |
| Class E（内部内容/语义） | 黑块、乱码、单字符公式异常、手机截图或内部公式不符时，先用正文/公式锚点定位**最终稿 AxMath 对象序号**；`inspect_axmath_tex.ps1` 只在临时副本导出该对象。冻结 OMML → 经验证的 TeX → 官方宏重建真实 AxMath 到**新 DOCX**；Word 输出损坏则用 OMML→LaTeX 备用转换。禁止以坏对象的导出或共享 WMF/预览为语义源。 |
| Class A/B/C/D | 仅对明确序号、已证实的目标执行 inline、baseline probe、OLE 壳校准或预览/缓存诊断；不可全局处理。 |

即使几何正常，撇号/导数也是独立语义风险。AxMath 2.7.0.58 经实机验证：一阶 `\prime`、二阶 `''`、三阶 `'''`；高风险源式必须用真实 AxMath 和冻结源稿审定，`PRIME_REVIEW_TEMPLATE.json` 不能自动填写通过。

集合符号 `∪∩⊂⊆` 等仅标记视觉优先复核，不凭源符号直接替换。独立数学段落可能发生原生 Word 自动居中 → AxMath OLE 左对齐的偏移；只有真实逐页对照证明且指定序号、输入 SHA 才可调用 `repair_standalone_display_alignment.py`，完成后重新验收。

页背景/边框不一致时只使用 `restore_word_page_layout.py` 从冻结源稿恢复 `w:background` 和 `w:pgBorders` 到**新的**工作稿；不触碰正文、公式、OLE。页面格式一致性是最终硬门禁。

AxMath 内部公式与预览不一致时用 `inspect_axmath_tex.ps1` **只在临时副本**导出检查；不要直接替换共享 WMF/预览。无证据不要运行几何试错脚本。

公式与 AxMath 对象一一对应时，Class E 的选定对象应生成哈希绑定审核记录：`inspect_axmath_tex.ps1 -Ordinals "<序号>"` → `python scripts/class_e_review.py prepare --source "<frozen>" --working "<working>" --inspection "<internal.json>" --ordinals "<序号>" --out "<review.json>"`。核对冻结 OMML、内部 TeX 后填写 `decision=verified_match` 或 `needs_rebuild` 与具体证据，运行 `class_e_review.py validate`，再用 `diagnose_after_conversion.py --class-e-review "<review.json>"` 复核。同一源稿/工作稿/导出哈希下，已证实正确的短公式不再反复进入 Class E 修复队列；文档一变，记录自动失效。内部为空或含控制字符不能标为正确。M2-B 拆行时先按已验 ledger 建立源式与结果对象的映射，不得套用一一对应序号。

## 4. 正式逐页验收（必做）

**先确认交付限制。** 本节的 `strict_final_compare.py` 会生成临时 PDF；当用户只要 Word 或禁止导出 PDF 时，不运行该命令或 `axmath.py compare`，先完成下面的内部导出，再按本节末尾的“Word 页面视图”分支验收。不得为了套用默认验收流程违背用户要求。

只有诊断队列不含待修复/未批准的关键语义风险时，**先对最终候选稿全部 AxMath 做一次只读内部导出**。对象数量正确、页面看着正常，都不能替代这一门禁；空内容或控制字符必须回到 Class E。`-Ordinals` 仍限于局部诊断，最终完整审核使用 `-AllAxMath`：

```powershell
powershell -File scripts\inspect_axmath_tex.ps1 -InputDocx "<final.docx>" -AllAxMath -OutputJson "<final-review>\FINAL_AXMATH_INTERNAL_TEX.json"
```

**允许生成 PDF 时**，导出完整且无空内容/控制字符后：

```powershell
python scripts\strict_final_compare.py --source "<frozen.docx>" --final "<final.docx>" --outdir "<final-review>" --expected-source-sha256 "<frozen_sha>" --internal-audit "<final-review>\FINAL_AXMATH_INTERNAL_TEX.json"
```

如果合法 `M2-B` 拆行，追加 `--repair-ledger "<verified-ledger.json>"`。转换、续跑、缓存恢复、诊断及最终比较必须使用同一 `--numbering-scope`；显式选择非默认值时每个阶段都须传入。

最终比较必须重新打开 Word、逐页渲染源稿和结果、生成全部并排截图。**相同页码不等于相同内容**：公式重排可改变页数和页边界，必须按题号、正文短语、公式序号等内容锚点跨页追踪。GPT **实际查看源稿和结果稿每一页**，对小型行内公式放大或裁出足够清晰的局部视图；检查 AxMath 正文语义、撇号/集合符号、跨行、居中、边界、背景/边框、正文与分页差异，并在 `VISUAL_REVIEW.json` 写明跨页锚点和局部视图证据后：

```powershell
python scripts\finalize_visual_review.py --report "<final-review>\STRICT_FINAL_COMPARE.json" --review "<final-review>\VISUAL_REVIEW.json"
```

**允许导出 PDF 时**，仅完整内部导出通过、`acceptance_pass=true`、最终非 AxMath 的 OLE 对象为零、非正文 OfficeMath 残留为零、全部公式可编辑、源稿未变、无遗留本任务 Word 进程，才是正式完成。不能用静态检查通过、对象数相同、页数相同、截图文件存在或完成宏返回码代替真实逐页检查。内部导出证明非空与可读取；非空但语义错误仍靠来源对照和视觉审查识别。

**用户只要求 Word 或明确不导出 PDF 时**：本仓库当前 `strict_final_compare.py`、`axmath.py compare` 和 `snapshot_docx.py` 的 `full` 模式会生成临时 PDF，禁止调用；不把其 `acceptance_pass` 写成已通过。仍执行 `-AllAxMath`、嵌入对象清点、源/结果哈希和内容核对，可按需使用 `snapshot_docx.py ... --profile geometry` 与 `audit_docx.py` 做不导出 PDF 的结构辅助检查；在真实 Word 页面视图逐页查看候选稿，并按题号及原始资料逐题核实公式、配图、图注和留白，记录实际看过的页码和发现的问题。报告分别写明已完成的 Word 视觉核查与未运行的 PDF 型严格比较；某项无法完成时不得声称全部验收通过。普通排版任务按排版参考文件检查，不凭空启动公式转换。

## 5. 输出、选项与性能纪律

正常成功仅需 `SOURCE_STATE.json`（原稿预检和断点来源）、`READY_FOR_GPT_REVIEW.json`（**唯一转换交接状态**）、冻结源稿、可恢复转换工作稿、`.conversion-checkpoint.json`、`.skill-report.json` 和必要的 AxMath watcher 日志。

- 没有可替换编号时，`working_docx` **直接指向转换工作稿**，不复制整份含 OLE 的 DOCX。
- 转换及核查默认 `--numbering-scope preserve`：不自动改写题号标点，并按原样审计；也可显式传入以保持跨版本清晰。同类兼职练习册在排版稿上默认将单一数字主问题号显示为“数字 + 英文句点”，多段题号、原题数字和小问标点不变；用户要求原标点时按原样保留。再以排版稿冻结转换，仍选 `preserve`。显式选择 `line-start` 会将 `1、` **和 `（1）、`** 改为点号，只有符合任务要求时才能使用；`anywhere` 还会改变 `第1、2项`，必须明确选择。旧版运行目录续跑前检查其原编号策略，新版运行目录会拒绝与原策略不一致的续跑。
- `normalize_numbering_punctuation.py` 只替换文本标点，不会把自动编号变为普通文字。题号应为普通文字时，先根据 Word 实际列表标签在工作副本逐题转成普通文本，再用 `python scripts/inspect_question_numbering.py "<docx>"` 排查直接或样式继承的列表、编号域及大纲标题，并在 Word 确认候选位置；脚本结果只供定位，不能代替逐题核实。
- 正常成功不写多份内容重复的运行结果、控制状态和日志 JSON；失败时保留诊断与恢复证据。旧运行目录中的历史文件不主动删除。
- 只有异常时才生成详细修复文件；无超宽公式不启动 Word 多行探针，无撇号风险不生成撇号审核表。
- 转换期间不做全量几何 COM 扫描；最终逐页查看需基于当前候选文件。仅在任务允许 PDF 时使用 PDF 渲染；旧 PDF 缓存不能冒充当前文件的验收证据。

高级故障指引和 AxMath COM 经验：**`docs/REPAIR_REFERENCE.md`**；脚本静态回归与完整测试都位于 `scripts/test_*.py`，CI 是辅助而非 Word/AxMath 实机验收。
