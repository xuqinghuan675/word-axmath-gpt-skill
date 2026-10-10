---
name: word-axmath-conversion
description: Convert Microsoft Word OfficeMath to genuine editable AxMath, with frozen-source safety, targeted repairs and strict visual acceptance.
---

# Word → AxMath Skill · 主流程（精简版）

**A new chat must read this file** before editing or converting any Word / OfficeMath / AxMath document. This file is authoritative after the user's current instruction. Detailed exceptional repair contracts are in `docs/REPAIR_REFERENCE.md`, loaded **only** when the diagnostic class requires them.

**唯一主链：原稿冻结 → 官方 AxMath 转换 → 快速诊断 → 必要时定向修复 → 唯一一次最终完整逐页视觉验收。**

### 永久安全约束

- 不覆盖源稿或正式完成稿；每一次修复写新 DOCX。冻结源稿和工作稿均绑定 SHA-256。
- 用户原有 Word/WPS 进程不可关闭、终止、修改；只操作经过 PID 和创建时间确证的本任务 Word 进程。
- Word COM 源稿读取与 AxMath 修改相互隔离（**source-reading from AxMath mutation**）；RPC 失败弃用已污染的 COM 实例。WPS 不能替代正式 Word+AxMath 转换。
- 每批 `doc.Save()`、进度检查、可验证断点恢复和最终 AxMath 可编辑性不可删除。
- 公式内容以冻结 OfficeMath/OMML 为真源；禁止凭宽度、页数、外观猜公式，禁止直接用损坏 AxMath 导出的 TeX 作为语义源。
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

**旧式公式预检与阻断**：预检额外输出所有 OLE 的 `ProgID`、所在 Word story/段落、异常对象清单；可单独运行 `python scripts/embedded_object_inventory.py "<docx>"` 定位。当 `state=blocked_unhandled_embedded_math` 时，禁止执行正式转换。必须先在**新 DOCX** 逐式核对旧式公式、恢复为原生 OfficeMath，重新预检通过后再转换。若出现 Word“Equation Editor 3.0 → OfficeMath”对话框，不得直接对全部公式批量确认；自动弹窗不证明数学语义正确。

## 2. 官方转换与断点

AxMath 2.7.0.58 的 `AMSMML2AM` **单次**约处理 64–66 个公式；这是插件内部机制，不去绕过。Skill 的总调用次数**没有 64 轮上限**：只要剩余 OfficeMath 大于零就继续，但每一轮必须有严格递减进度，否则停止。

- `WaitingConvert=1` 仅用作完成确认，不是调速开关；未经证实的对话框不得自动关闭。
- 每批保存后生成哈希绑定的 `.conversion-checkpoint.json`；源稿、工作稿、残余公式数和轮数必须一致才允许续跑。
- 同用户转换互斥；每轮 Word/AxMath 超时看门狗默认 1800 秒，且只能操作已证实属于任务的进程。
- 断点恢复既不重新覆盖工作副本，也不跳过事后完整诊断。

已有运行目录续跑（仅单份文档）：
```powershell
python scripts\one_click_convert.py --input "<original.docx>" --resume-run "<existing-run-dir>"
```

高级单文件续跑：`python scripts\run_skill.py --input "<frozen.docx>" --output "<partial-working.docx>" --resume`。其余内部 PowerShell/看门狗脚本由入口调用，不要手工并发运行。

## 3. 异常时按类修复（读高级参考，不走全流程）

`diagnose_after_conversion.py` 是唯一 dispatcher。异常时先读 `docs/REPAIR_REFERENCE.md` 的对应类和 Word/COM 规则，并始终基于当前冻结源稿、工作稿、哈希及证据执行：

| 状态 | 必须保留的处理原则 |
|---|---|
| `exact` | 残余 OfficeMath=0 且 AxMath 对象数精确，继续语义/几何诊断。 |
| `STOP_UNHANDLED_EMBEDDED_MATH` | 旧式 MathType / Equation Editor 3.0、未知 OLE 或正文之外的 OfficeMath 未验明：保留源对象、预览和原始哈希，逐一核实并在副本中修复，重新预检后才可转换。 |
| `M1_MULTISIBLING_OMATHPARA` | 只有源稿证明确实存在多个直系 `m:oMath` sibling 才恢复合并丢失对象；先修结构，后几何。 |
| `M2_SOURCE_VISUAL_WRAP_LOSS` | 必须同时证明源稿真实视觉换行和工作稿越过**实测**右边界；不能凭固定长度或页数猜。 |
| `M2-B` | 只有源稿视觉行证据 + hash-bound repair ledger 才能一行拆一个 AxMath；M2-A 的内部 `aligned` 则仍是一个 AxMath。 |
| Class E（语义） | 冻结 OMML → 经验证的 TeX → 真实 AxMath；如果 Word 输出损坏则使用 OMML→LaTeX 备用转换，写入前核对语义。 |
| Class A/B/C/D | 仅对明确序号、已证实的目标执行 inline、baseline probe、OLE 壳校准或预览/缓存诊断；不可全局处理。 |

即使几何正常，撇号/导数也是独立语义风险。AxMath 2.7.0.58 经实机验证：一阶 `\prime`、二阶 `''`、三阶 `'''`；高风险源式必须用真实 AxMath 和冻结源稿审定，`PRIME_REVIEW_TEMPLATE.json` 不能自动填写通过。

集合符号 `∪∩⊂⊆` 等仅标记视觉优先复核，不凭源符号直接替换。独立数学段落可能发生原生 Word 自动居中 → AxMath OLE 左对齐的偏移；只有真实逐页对照证明且指定序号、输入 SHA 才可调用 `repair_standalone_display_alignment.py`，完成后重新验收。

页背景/边框不一致时只使用 `restore_word_page_layout.py` 从冻结源稿恢复 `w:background` 和 `w:pgBorders` 到**新的**工作稿；不触碰正文、公式、OLE。页面格式一致性是最终硬门禁。

AxMath 内部公式与预览不一致时用 `inspect_axmath_tex.ps1` **只在临时副本**导出检查；不要直接替换共享 WMF/预览。无证据不要运行几何试错脚本。

## 4. 正式逐页验收（必做）

只有诊断队列不含待修复/未批准的关键语义风险时：

```powershell
python scripts\strict_final_compare.py --source "<frozen.docx>" --final "<final.docx>" --outdir "<final-review>" --expected-source-sha256 "<frozen_sha>"
```

如果合法 `M2-B` 拆行，追加 `--repair-ledger "<verified-ledger.json>"`。若显式使用了 `--numbering-scope anywhere`，诊断和最终比较也必须传同一选项。

最终比较必须重新打开 Word、逐页渲染源稿和结果、生成全部并排截图。GPT **实际查看每一页**，检查 AxMath 正文语义、撇号/集合符号、跨行、居中、边界、背景/边框、正文与分页差异；填写真实 `VISUAL_REVIEW.json` 后：

```powershell
python scripts\finalize_visual_review.py --report "<final-review>\STRICT_FINAL_COMPARE.json" --review "<final-review>\VISUAL_REVIEW.json"
```

**仅 `acceptance_pass=true`、最终非 AxMath 的 OLE 对象为零、非正文 OfficeMath 残留为零、全部公式可编辑、源稿未变、无遗留本任务 Word 进程，才是正式完成。** 不能用静态检查通过、页数相同、截图文件存在或完成宏返回码代替真实逐页检查。

## 5. 输出、选项与性能纪律

正常成功仅需 `SOURCE_STATE.json`（原稿预检和断点来源）、`READY_FOR_GPT_REVIEW.json`（**唯一转换交接状态**）、冻结源稿、可恢复转换工作稿、`.conversion-checkpoint.json`、`.skill-report.json` 和必要的 AxMath watcher 日志。

- 没有可替换编号时，`working_docx` **直接指向转换工作稿**，不复制整份含 OLE 的 DOCX。
- 确实有编号时才额外生成规范化工作稿。默认仅转换段首/Word 换行后的 `1、`、`（1）、` 为点号；`--numbering-scope anywhere` 还会改变 `第1、2项`，必须明确选择，不能作为常规默认。
- 正常成功不写多份内容重复的运行结果、控制状态和日志 JSON；失败时保留诊断与恢复证据。旧运行目录中的历史文件不主动删除。
- 只有异常时才生成详细修复文件；无超宽公式不启动 Word 多行探针，无撇号风险不生成撇号审核表。
- 转换期间不做全量几何 COM 扫描；最终所有页面实际渲染仍只做一次。来源和输出 Word 的 PDF 缓存不能冒充“新鲜验收”。

高级故障指引和 AxMath COM 经验：**`docs/REPAIR_REFERENCE.md`**；脚本静态回归与完整测试都位于 `scripts/test_*.py`，CI 是辅助而非 Word/AxMath 实机验收。
