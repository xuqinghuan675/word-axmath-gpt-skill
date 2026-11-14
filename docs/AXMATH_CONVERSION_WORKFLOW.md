# AxMath 转换与验收详细工作流

此文件保留原主 Skill 中的全部公式转换安全约束、预检、断点、异常分派及最终验收细节。
入口和执行优先级以 ../SKILL.md 为准。数学资料编号请先读 NUMBERING_AND_OUTLINE.md，内容排版见 MATH_LAYOUT_REFERENCE.md。
本文件只在需要转换 OfficeMath/OMML → AxMath 时加载；纯文字排版不用加载。

## 永久安全约束

- 不覆盖源稿或正式完成稿；每一次修复写新 DOCX。冻结源稿和工作稿均绑定 SHA-256。
- 用户原有 Word/WPS 进程不可关闭、终止、修改；只操作经过 PID 和创建时间确证的本任务 Word 进程。
- Word COM 源稿读取与 AxMath 修改相互隔离（**source-reading from AxMath mutation**）；RPC 失败弃用已污染的 COM 实例。WPS 不能替代正式 Word+AxMath 转换。
- 每批 `doc.Save()`、进度检查、可验证断点恢复和最终 AxMath 可编辑性不可删除。
- OfficeMath 转换以冻结 OfficeMath/OMML 为语义来源；由扫描资料新建公式时以原 PDF 及对应题目为准，OCR 仅供对照。禁止凭宽度、页数或外观猜公式，禁止直接用损坏 AxMath 导出的 TeX 作为语义源。
- OfficeMath 与 AxMath 不是全部公式类型：旧版 Equation Editor 3.0（如 `Equation.3`）、MathType（如 `Equation.DSMT4`）及未知 OLE 必须计入预检，不能仅凭 OfficeMath=0 / AxMath 数目一致宣告完成。
- 禁止全局修改 OLE 大小、`w:position`、预览媒体或盲目拆长公式。页面数量只是观察数据，不是修复目标。

### 首轮生成硬规则（强制）

**所有规则在第一次写入新 DOCX 时执行。不要先按旧方式生成，再等终稿审计发现错误后补救。**

#### 0. 冻结源稿并处理正文标点

- 原 PDF/DOCX 永不覆盖。MinerU 输出先保存为 mineru_original.docx；已有 DOCX 也先保留原件。
- 用下面的专用脚本生成新的 frozen_source.docx。它只处理主文档正文普通 w:t 文本节点中的中文句号，不进入 AxMath OLE、OfficeMath、图片/图表、字段、文本框或页眉页脚；并验证 DOCX 中除 word/document.xml 外的其他包部件全部保持原样。
- 报告必须记录替换数量，且普通正文残留数为 0。禁止用 Word 全局查找替换或直接修改 OLE 内部。

    python scripts\normalize_body_punctuation.py --source "mineru_original.docx" --output "frozen_source.docx" --report "review\BODY_PUNCTUATION.json"

frozen_source.docx 是后续所有预检、公式目标清单和转换的唯一冻结来源；记录其 SHA-256。原件与 MinerU 原始结果继续保留。

#### 1. 题号首轮写成普通文字

- 生成前运行只读编号检查，并读取 Word 实际显示的题号；检查段落/样式继承链中的 w:numPr、w:numId、w:ilvl、numbering.xml、LISTNUM/SEQ/AUTONUM 字段、w:pStyle 和有效 w:outlineLvl。不得只凭缩进判断。
- 题号直接写入正文文字，例如 1.、(1)、（1）；保留原序号、标点、跳号和层级。禁止自动编号、多级列表、编号域和 Heading 折叠身份。
- 同级编号使用统一左缩进、悬挂缩进，并在编号后设置制表位；正文与续行从同一制表位对齐。保留不同层级原有缩进差异，不能只统一段落缩进。

    python scripts\inspect_question_numbering.py "frozen_source.docx"

#### 2. 先建公式目标清单，再生成公式

生成前必须创建哈希绑定的 SOURCE_FINAL_FORMULA_MAP.json 模板：

    python scripts\math_layout_contract.py map-template --source "<READY_FOR_GPT_REVIEW.json 中的 frozen_source 绝对路径>" --out "<final-review>\SOURCE_FINAL_FORMULA_MAP.json"

必须使用 READY_FOR_GPT_REVIEW.json 记录的实际 frozen_source 路径和哈希；后续转换、内部导出和最终比较始终引用同一冻结来源。模板会为来源中的每个 OfficeMath/AxMath 对象建立 source_formula 目标；每个目标都必须在生成前从冻结来源填写 source_anchor 和 expected_tex，禁止从终稿导出反推。随后必须补齐两类目标：

- source_formula_visual：PDF/扫描图中的源公式未以 OMML/AxMath 对象存在时，按 PDF 页码和区域建立目标，并填写源式 expected_tex。
- added_body_math：正文中独立变量、表达式等虽不是源公式对象，但按本任务要求应生成 AxMath OLE 时建立目标，并填写其来源文字锚点和 expected_tex。

目标清单必须在首次生成公式前冻结。不得等到终稿发现多出对象或少了公式，再临时猜测来源。

#### 3. 公式生成规则

- 源 OfficeMath/OMML 使用官方 AxMath 转换宏 AMSMML2AM；从冻结来源核定的 TeX 以及正文数学目标使用官方 AMSTeX2AM。必须生成真实、可编辑的 AxMath OLE，禁止用斜体普通文字、Word 公式或图片代替。
- 调用宏前先划清文字边界：中文说明、题号、题目标记等非数学内容先分离为普通 Word 文字，数学部分单独生成 AxMath。例如“求”和 *(2) 留在普通正文，不要塞进 OLE 再等终稿导出后补猜。
- 源式为大积分时使用 \int，禁止生成 \smallint。normalize_axmath_tex.py 会把来源 TeX 中的 smallint 规范为 int；最终 AxMath 原始导出若仍出现 smallint，验收直接失败。
- 下置极限统一使用 \underset{...}{\lim}；同时有上下限时使用 \underset{下标}{\lim}^{上标}。不要使用 \lim\limits 或 \nolimits：实机往返已确认 AxMath 导出可能规范化掉 limits，导致下置位置丢失。仅上置脚本可写成 \lim^{...}。limsup/liminf 不得误识别为 lim。
- 导数撇号、分式、根式、矩阵、方程数组、分隔符、上下标、关系符和运算符必须按冻结来源保真；禁止凭损坏 AxMath 对象自己的 TeX 导出反向猜原式。

#### 4. 完成源到终稿映射

生成后填写 SOURCE_FINAL_FORMULA_MAP.json：

- source_formula：每个源对象记录一个或多个 final_ordinals；一源拆成多个终稿 OLE 合法。
- source_formula_visual / added_body_math：每个目标必须有 source_anchor、expected_tex 和 final_ordinals。
- final.path、final.sha256、final.axmath_count 必须对应最终候选 DOCX。
- 每个源公式都必须有归属；每个终稿 AxMath 对象必须且只能归属一个目标。新增对象、拆分对象不得漏记。
- 验收不再要求源公式总数与终稿 OLE 总数相等；由映射清单验证覆盖率，并按映射汇总核对结构、撇号、极限位置、运算符和对象数量。

#### 5. 一轮最终验收

所有生成和定向修复完成后，对最终候选只做一次完整只读 -AllAxMath 导出。随后按交付限制二选一：

- **只交付 Word**：运行 math_layout_contract.py validate 一次，命令必须带 --formula-map；来源本身已有 AxMath 时追加 --source-internal-audit。
- **允许生成 PDF**：直接运行第 4 节 strict_final_compare.py 并带 --formula-map；它会调用同一映射门禁。不要再单独运行 math_layout_contract.py validate，避免重复验收。

Word-only 命令：

    python scripts\math_layout_contract.py validate --source "<READY 中的 frozen_source 绝对路径>" --final "final_candidate.docx" --internal-audit "<review>\FINAL_AXMATH_INTERNAL_TEX.json" --formula-map "<final-review>\SOURCE_FINAL_FORMULA_MAP.json" --out "<final-review>\MATH_LAYOUT_VALIDATION.json"

验证必须 passed=true。最终同时确认：

- 所有 AxMath OLE 均成功导出；无残留 OfficeMath、旧式公式编辑器或未知 OLE。
- 题号是普通正文文字；目标段落无活动 w:numPr、编号域或 Heading 折叠身份。
- 映射清单覆盖全部源公式和终稿对象；拆分、新增对象均有归属。
- 正文标点报告的替换数量可追溯，普通正文中文句号残留数为 0；公式/OLE 内容未被全局替换。
- 最后只做一次真实 Word 页面渲染和逐页验收。

若验收发现问题，按对象序号或段落锚点定向修复；修复时只复核受影响对象，不重新做无证据的全书扫描。修复完成后对最终候选完成一次最终验收，不循环重复审计。strict_final_compare.py 也会调用同一公式映射门禁；需要该命令时传入 --formula-map，不要先单独 validate、再让 strict_final_compare 重复跑同一门禁。



## 1. Mandatory entry protocol（普通任务）

默认只需下列三个入口，不必阅读/运行全套修复工具：

**预检（不会修改 DOCX）：**
```powershell
python scripts\one_click_convert.py --input "<frozen_source.docx>" --inspect-only
```

**批量转换（只转换一次，不重复重头做）：**
```powershell
python scripts\one_click_convert.py --input "<frozen_source.docx>"
```

读取生成的 `READY_FOR_GPT_REVIEW.json`，选用其中的 `frozen_source` 和 `working_docx`。**不得**从旧聊天推断路径或跳过诊断。

**转换后唯一正常诊断：**
```powershell
python scripts\diagnose_after_conversion.py --source "<frozen_source>" --working "<working_docx>" --outdir "<review>"
```

读取 `NEXT_ACTION.json`。当其 `status=ready_for_strict_final_compare` 时直接进入第 4 节最终验收；如指出特定异常，只按第 3 节的修复类处理，不进入无意义全书扫描。正常诊断只使用 DOCX/XML 快筛和必要的少量 Word 证实；`--deep-geometry` 仅在真实视觉问题难以定位时显式使用。

预检会检查冻结原稿可读性、OfficeMath/AxMath 数量、段落、Windows Word PID、AxMath.dotm、PowerShell、Python 依赖（psutil、pywin32、lxml、numpy、Pillow、PyMuPDF、omml2latex）。中文路径控制台编码错误不等于 DOCX 损坏。

**旧式公式预检与阻断**：预检额外输出所有 OLE 的 `ProgID`、所在 Word story/段落、异常对象清单；可单独运行 `python scripts/embedded_object_inventory.py "<docx>"` 定位。当 `state=blocked_unhandled_embedded_math` 时，禁止执行正式转换。通常须在**新 DOCX** 逐式核对旧式公式、恢复为原生 OfficeMath，重新预检通过后再转换；若用户明确禁止使用 Equation/OfficeMath，则不得执行恢复环节，须寻找现有合规的 AxMath 原生处理方式，无法实现时如实说明。若出现 Word“Equation Editor 3.0 → OfficeMath”对话框，不得直接对全部公式批量确认；自动弹窗不证明数学语义正确。


## 2. 官方转换与断点

AxMath 2.7.0.58 的 `AMSMML2AM` **单次**约处理 64–66 个公式；这是插件内部机制，不去绕过。Skill 的总调用次数**没有 64 轮上限**：只要剩余 OfficeMath 大于零就继续，但每一轮必须有严格递减进度，否则停止。

- `WaitingConvert=1` 仅用作完成确认，不是调速开关；本机截图与已验证的正常批次暂停提示严格一致时，也可关闭该任务所属弹窗。Office 命令冲突提示必须保持原样供排查；未经证实的对话框不得自动关闭。视觉识别不匹配时继续按原看门狗规则失败关闭。
- 每批保存后生成哈希绑定的 `.conversion-checkpoint.json`；源稿、工作稿、残余公式数和轮数必须一致才允许续跑。
- 同用户转换互斥；每轮 Word/AxMath 超时看门狗默认 1800 秒，且只能操作已证实属于任务的进程。
- 断点恢复既不重新覆盖工作副本，也不跳过事后完整诊断。

已有运行目录续跑（仅单份文档）：
```powershell
python scripts\one_click_convert.py --input "<frozen_source.docx>" --resume-run "<existing-run-dir>"
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
python scripts\strict_final_compare.py --source "<READY 中的 frozen_source 绝对路径>" --final "<final.docx>" --outdir "<final-review>" --expected-source-sha256 "<frozen_sha>" --internal-audit "<final-review>\FINAL_AXMATH_INTERNAL_TEX.json" --formula-map "<final-review>\SOURCE_FINAL_FORMULA_MAP.json"
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
