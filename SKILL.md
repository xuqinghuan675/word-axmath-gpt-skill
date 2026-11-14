---
name: word-axmath-conversion
description: Handle math PDF/DOCX formatting, plain-text numbered exercises, non-collapsible headings, and genuine editable AxMath conversion through MinerU and Word.
---

# AxMath 数学资料排版 Skill（唯一入口）

**每次新任务先读本文件，再按下表选择对应文档；不要凭旧聊天猜工作流。**

| 任务 | 必读文件 |
|---|---|
| 数学讲义、练习册、题目排版及分页留白 | `docs/MATH_LAYOUT_REFERENCE.md` + `docs/NUMBERING_AND_OUTLINE.md` |
| 编号错位、Word 自动编号、WPS/Word 标题折叠 | **`docs/NUMBERING_AND_OUTLINE.md`（唯一权威）** |
| PDF → 可编辑 DOCX | `docs/MINERU_API.md` |
| DOCX OfficeMath/OMML → 原生 AxMath | `docs/AXMATH_CONVERSION_WORKFLOW.md` |
| Codex 在 Windows 本地直接操作 Word/AxMath | `docs/CODEX_DIRECT.md` |
| `NEXT_ACTION.json` 指定异常类别 | `docs/REPAIR_REFERENCE.md` 对应类别（只在出现时读） |

## 1. 编号和层级：首次制作时落实

- **所有题号为直接输入的普通文字**，保留原题号、标点、跳号；禁止 Word 自动编号、多级列表和 `LISTNUM/SEQ/AUTONUM` 等题号字段。
- **同级编号左缘对齐**，编号后正文和续行用统一制表位与悬挂缩进对齐；不同层级保持原稿本来的视觉缩进差异，不强行压成同一级。
- **消除可折叠标题身份**：题目段落使用普通正文或自定义普通样式；不能只将 `w:outlineLvl` 设为 `9`。若原段落使用 `Heading 1/2` 等内置样式，必须克隆视觉格式到不含 Heading 身份的自定义样式，同步修改 `w:pStyle`、`basedOn`、`next`、`link`，并移除有效大纲级别。
- **先样页，后全文**：样页在 Word/WPS 查看编号起点、续行和折叠箭头，合格后使用相同结构批量制作；不要终稿发现错误再逐题返工。
- 编辑已有 DOCX 时，先读取 Word 实际显示的编号并物化为普通文字，然后解除自动列表。只在工作副本中修改 `word/document.xml`、`word/styles.xml` 等必要部分，不碰 AxMath OLE 和媒体。

**`--numbering-scope preserve` 只保留编号文字的标点，不能将自动编号转为普通文字，也不能消除折叠箭头。** 具体方法及验收只看 `docs/NUMBERING_AND_OUTLINE.md`。

## 2. 排版任务

先按 `docs/MATH_LAYOUT_REFERENCE.md` 确定内容、题目归属、公式、图注、留白、分页与交付格式。用户要求学生练习册时使用其中的同类兼职约定。所有公式默认使用真实可编辑的 AxMath；已有 OfficeMath 只读作为转换来源，不以图片、MathType 或新建 Word 公式冒充。仅排版时不无故启动 AxMath 转换。

用户只要 Word 或禁止导出 PDF 时，**不产生交付或临时 PDF**，在真实 Word 页面视图逐页验收。保留原 PDF / DOCX 与工作稿，不覆盖原件。

## 3. PDF → DOCX：MinerU

- **GPT 本机**：可直接调用本机 MinerU Precision API。
- **Codex**：优先使用当前 Codex 对话侧边栏已登录的 MinerU 网页，无需 Token；预计排队超过 120 秒、连续 60 秒无进展或前端卡住时，立即改用本机 API，不继续傻等。
- 两者功能权限一致，区别是使用优先级。API 只负责 PDF → DOCX，不影响后续 AxMath 转换方式。
- 本机 API 命令：`python scripts\mineru_api.py parse --input "<原稿.pdf>" --outdir "<独立工作区>"`。
- 两条路线统一冻结 `mineru_original.docx`，保留原始 ZIP、识别输出、任务状态和恢复记录；网络中断从已有任务恢复，不重复提交。
- Token 存在本机忽略的 `secrets\mineru_token.txt`；不得输出、上传、写入 Git 或聊天内容。当前任务明确禁止使用 Token 时遵守。

具体 API、状态码、断点与故障见 `docs/MINERU_API.md`。

## 4. AxMath 公式转换主链（只有需要转换时执行）

依次运行：

    python scripts\one_click_convert.py --input "<原稿.docx>" --inspect-only
    python scripts\one_click_convert.py --input "<原稿.docx>"
    python scripts\diagnose_after_conversion.py --source "<frozen_source>" --working "<working_docx>" --outdir "<review>"

通过 `READY_FOR_GPT_REVIEW.json` 读取真实路径；通过 `NEXT_ACTION.json` 选择后续修复或最终验收。已有运行目录用 `--resume-run` 恢复，禁止无故重头转换。官方 AxMath 宏与真实 OLE 可编辑性、每批保存/断点、公式结构及内部 TeX 门禁等全部严格遵照 `docs/AXMATH_CONVERSION_WORKFLOW.md`。

Codex 使用本机终端和 Word/AxMath，不依赖网页电脑插件；详情见 `docs/CODEX_DIRECT.md`。

## 5. 全局收口

- 不覆盖原稿或已有完成稿；冻结来源 SHA-256；定向修改输出新 DOCX。
- 不关闭或终止用户原有 Word/WPS；真实转换只用本任务确证的 Microsoft Word/AxMath。
- 结构检查、内部公式验证和实际页面检查按任务做；**只做一次有价值的最终验收**，避免反复全面审计。
- 不仅凭 `outlineLvl=9`、静态检查、页数/对象数相同宣称排版通过；用户要求 Word/WPS 表现时必须真实打开相应应用检查。
- 每次完成后说明实际验证范围与未完成项。用户当前明确要求优先于历史默认值。
