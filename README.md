# word-axmath-gpt-skill

把 Microsoft Word 原生 OfficeMath / OMML 批量转换为真实可编辑的 `Equation.AxMath`，并使用冻结源稿、确定性故障分类和逐页视觉验收完成生产级收口。

## 最重要的一句话

**不要转换完就开始猜。必须先读 `SKILL.md`，然后运行 `diagnose_after_conversion.py`，按 `NEXT_ACTION.json` 的 repair class 走。**

正式流程只有一条：

> 冻结源稿 → 官方转换一次 → 确定性诊断 → 当前类唯一修复路线 → 重新诊断 → 最终逐页验收

新对话不得从旧聊天记忆、页数差、公式宽度印象或历史报告直接发明修复。

## 快速开始

### 1. 安装依赖

```powershell
python -m pip install -r requirements.txt
```

依赖包括 `pywin32`、`PyMuPDF`、`omml2latex` 等。不要跑完大文档后才发现渲染/OMML fallback 缺包。

### 2. 预检

```powershell
python scripts\one_click_convert.py --input "<源.docx>" --inspect-only
```

预检会同时检查：Word/AxMath 数量、段落、锁文件、现有 Word PID、OMML 多 sibling 结构、Python 依赖、PowerShell、`AxMath.dotm`。

### 3. 转换一次

```powershell
python scripts\one_click_convert.py --input "<源.docx>"
```

源文件不会被覆盖。工作区会保存 frozen source、working copy、转换报告和恢复边界。

### 4. 强制诊断入口

读取 `READY_FOR_GPT_REVIEW.json` 中的 frozen source / working 路径，然后：

```powershell
python scripts\diagnose_after_conversion.py --source "<frozen-source.docx>" --working "<working.docx>" --outdir "<review>"
```

只按 `NEXT_ACTION.json` 执行。每次修复必须输出到新的 DOCX，再重新诊断。

默认诊断是 **静态 XML 快筛 + 只对疑点做 Word 证实**，不会每轮用 COM 扫 1500+ 个公式。只有最终/source-vs-working 视觉证据发现难分类异常时，才显式加 `--deep-geometry` 做全量几何诊断。

## 已覆盖的生产故障

### M1：`m:oMathPara` 多 sibling 被 AxMath 合并

典型症状：源稿 raw OfficeMath 数比 AxMath 多，但差值恰好等于 frozen source 中多 sibling display group 的 extra nodes。

这不是“随机丢公式”。`source_math_structure.py` 会在转换前就记录这个结构，`audit_docx.py` 会把它分类成 `known_multisibling_collapse`。

```powershell
python scripts\build_source_tex_map.py multisibling --source "<source>" --working "<working>" --out "<m1.json>"
powershell -File scripts\repair_multisibling_groups.ps1 -InputDocx "<working>" -OutputDocx "<m1-fixed>" -MapPath "<m1.json>"
```

M1 是唯一允许“1 个塌缩 AxMath → 多个 AxMath”的类，因为 frozen source 自身就包含多个 raw OfficeMath 子对象。

### M2：超长公式失去 Word 自动换行，右侧冲出页面

检测不使用固定宽度阈值。只有同时满足：

- source OfficeMath 在 Word 中真实跨视觉行；
- working AxMath OLE 越过实测的单栏正文右边界；

才进入 M2。

```powershell
powershell -File scripts\export_source_visual_lines.ps1 -SourceDocx "<frozen-source>" -Ordinals "<ordinals>" -OutputJson "<lines.json>"
python scripts\build_source_tex_map.py visual-wrap --visual-report "<lines.json>" --working "<working>" --out "<m2.json>"
powershell -File scripts\repair_axmath_from_approved_tex.ps1 -InputDocx "<working>" -OutputDocx "<m2-fixed>" -MapPath "<m2.json>"
```

修复器把 source-derived visual lines 合成一个 `aligned` TeX donor。

**M2-A 约束：** 当 contract 是内部 `aligned` 多行 AxMath 时，1 个 source 公式保持 1 个 AxMath；不要把这条规则套到所有长公式。

十三月实机已验证 M2-A：三行 synthetic `aligned` 和真实四行长公式都经 `AMSTeX2AM` 得到 exactly one `Equation.AxMath`。

**M2-B：按原稿视觉数学行恢复多个 AxMath。** `待排版5` 的 24 个长公式段落实修证明，有些位置需要“每个原稿视觉行一个 AxMath + Word soft break”才能稳定还原原版。此时 AxMath 对象数增加是预期行为，但必须有 hash-bound repair ledger。

```powershell
python scripts\diagnose_local_layout.py --source "<frozen-source>" --working "<working>" --start-anchor "<唯一文字锚点>" --outdir "<review>"
python scripts\build_source_tex_map.py visual-line-split --visual-report "<visual-lines.json>" --working "<working>" --paragraph-plan "<LOCAL_LAYOUT_PLAN.json>" --out "<split-map.json>"
powershell -File scripts\repair_visual_line_split.ps1 -InputDocx "<working>" -OutputDocx "<split-fixed>" -MapPath "<split-map.json>"
python scripts\validate_local_visual_line_repair.py --baseline "<working>" --candidate "<split-fixed>" --map "<split-map.json>" --out "<split-ledger.json>"
```

M2-B 的断行点只能来自 frozen source 的真实视觉行，禁止按字符数、屏幕宽度、像素位置或猜测的等号机械切割。ledger 必须证明：锚点前前缀未变、非数学正文未变、源稿 SHA 未变、每个目标段的 AxMath/soft-break 数符合 source visual line、总 AxMath 增量精确匹配 map。最终 `strict_final_compare.py` 必须通过 `--repair-ledger` 接受这类 intentional split。

### Class D / E：页面预览正常，但双击 AxMath 后变短、变成别的公式

先判断是 preview/cache 错还是 OLE 内部语义真坏：

```powershell
powershell -File scripts\inspect_axmath_tex.ps1 -InputDocx "<working>" -Ordinals "<targets>" -OutputJson "<internal.json>"
```

它只在临时复制对象上运行 `AMSAM2TeX`，不会改 working。

- internal 正确、preview 错 → Class D;
- internal 短/错/不相关 → Class E，从 frozen source 同 ordinal 重建;
- prime/derivative 是独立语义风险，不要求先出现版式异常。所有 `' / ′ / ″ / ‴ / ⁗` 源公式都进入 `prime_semantic_candidates`，走 frozen-source Class E。AxMath 2.7.0.58 已实机确认：一阶 `\prime`、二阶 `''`、三阶 `'''`；raw Unicode `‴` 直接喂给 `AMSTeX2AM` 会回转成 `?`。先用 `normalize_axmath_tex.py` 归一化，写回前还要由真实 AxMath 再 round-trip 验证 prime 阶数。

Word LaTeX 导出也不能因为“非空”就相信。实测出现过 control chars、`▒`、`〖〗`、错误线性表示。`build_source_tex_map.py` 会校验；坏输出回退到 frozen OMML → LaTeX。

## 昨天真实踩坑已经写入设计

- `after_omath=0` 可能只是控制文件占位，不能当转换完成信号。
- Chinese path + legacy console encoding 不能把输出错误误判成文档错误；CLI 强制 UTF-8。
- `$PID` 是 PowerShell 保留变量，脚本统一用 `$wordPid`。
- Word COM 的 `InlineShapes` 在 PowerShell 下不能可靠用 `foreach`：同一份 1521-shape 文档实测，显式 `.Count/.Item(i)` 得到 1520 个 AxMath，而 `foreach` 得到 0；生产脚本统一改为索引枚举，并禁止 `return ,@($arr)` 再被外层 `@(...)` 包成单元素嵌套数组。
- `DispatchEx("Word.Application")` 也不能单独当“新进程证明”。实测 Word 可以复用已有 headless `/Automation -Embedding` server；`OwnedWord` 现在必须先证明出现新的独立 WINWORD PID，证明不了就只释放代理并失败，绝不 `Quit()`、绝不改已有 Word。
- `Get-FileHash` 对仍有 Office/OLE 共享句柄的 DOCX 可能报“being used by another process”，即使 ZIP/Python 仍能读取；修复脚本统一用允许 `ReadWrite|Delete` 共享的只读 SHA-256 流，避免把共享锁误判成文件损坏。
- Program Files x86 不能用错误的 unbraced env 语法。
- source `OMath.Range` 与 AxMath add-in 操作混在一个长 COM 生命周期里会出现 `PROPERTYGET` / `0x800706BE`；source reading 和 AxMath mutation 分阶段，困难 source extraction 按 ordinal 隔离 Word 实例。
- live original 可能被 Word 独占导致 hash 失败；语义修复统一基于 frozen source。
- Phase A 的 64–66 公式批量是 AxMath 稳定性机制，不是性能 knob。
- page count 只能当验收观察，不能用来指导 width/height/w:position 猜参数。
- 中间循环默认使用静态 XML + 疑点 Word 定向证明；全量 geometry audit 只在显式 `--deep-geometry` 升级时运行，完整 PDF + 每页对照只留最终一次。

## 工具地图

| 任务 | 工具 |
|---|---|
| Source OMML 结构 / M1 signature | `source_math_structure.py` |
| 内容、段落、公式计数 | `audit_docx.py` |
| 强制 repair dispatcher | `diagnose_after_conversion.py` |
| 中间 source/working 几何证据 | `snapshot_docx.py --profile geometry` |
| same-line / tiny shell / M2 overflow / box 诊断 | `formula_geometry_audit.py` |
| M1 frozen-source map | `build_source_tex_map.py multisibling` |
| M1 repair | `repair_multisibling_groups.ps1` |
| M2 source visual lines | `export_source_visual_lines.ps1` |
| M2-A aligned map | `build_source_tex_map.py visual-wrap` |
| M2-B 局部锚点/段落映射 | `diagnose_local_layout.py` / `build_local_layout_plan.py` |
| M2-B visual-line map | `build_source_tex_map.py visual-line-split` |
| M2-B repair / ledger | `repair_visual_line_split.ps1` / `validate_local_visual_line_repair.py` |
| Class E source Word LaTeX | `export_source_word_latex.ps1` |
| Prime donor 规范化 | `normalize_axmath_tex.py` |
| Prime 真实 AxMath contract probe | `probe_axmath_prime_contract.ps1` |
| Prime 写回硬门 | `axmath_prime_contract.ps1` |
| M2 / Class E single-object repair | `repair_axmath_from_approved_tex.ps1` |
| AxMath internal content diagnostic | `inspect_axmath_tex.ps1` |
| Low-risk inline Class A | `repair_axmath_inline_roundtrip.ps1` |
| Class B probe only | `rebuild_axmath_baselines.ps1` |
| Class C OLE shell | `calibrate_axmath_boxes.py` |
| Final full source-vs-final | `strict_final_compare.py` |
| Hash-bound final acceptance | `finalize_visual_review.py` |

## 性能

Phase A 真正耗时主要在 AxMath 宏内部。历史 1335 公式生产跑约 99.02% 时间在 `AMSMML2AM`；另一份 1574 公式文档实际转换约 32 分 40 秒、24 批。

因此优化重点不是绕过 batch safety，而是把后处理从“人工猜 bug”变成自动分类和唯一 repair route：

- 预先记录 M1 结构；
- 只在 count exact 后做 geometry；
- M2 用 source visual line + measured overflow；
- Class D/E 先辨 internal vs preview;
- repair iteration 不全页渲染；
- final 只 full-render 一次。

## 禁止事项

- 不覆盖 source / finished file。
- 不全局调 width/height/w:position。
- 不为了页数相等缩公式。
- 不允许**未经 frozen-source 视觉行证据与 repair ledger 验证**就把一个 M2 source 公式拆成多个 AxMath；M2-B 是有证据的正式路线，不属于碰运气拆分。
- 不拿坏 AxMath 自己导出的 TeX 作为 Class E 语义真源。
- 不因 same-line group 报警就整组 roundtrip。
- 不在正式文档上做参数搜索。
- 不杀用户已有 Word。

## 最终验收

```powershell
python scripts\strict_final_compare.py --source "<frozen-source>" --final "<final>" --outdir "<compare>" --expected-source-sha256 "<sha>"
```

然后必须逐页检查所有 source-vs-final 图片，再：

```powershell
python scripts\finalize_visual_review.py --report "<compare>\STRICT_FINAL_COMPARE.json" --review "<compare>\VISUAL_REVIEW.json"
```

只有 `acceptance_pass=true` 才是正式完成。
