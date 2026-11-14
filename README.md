# word-axmath-gpt-skill

将数学 PDF / DOCX 整理为 Word 练习资料，保留可编辑 AxMath 公式。Microsoft Word 和 AxMath 需单独安装，仓库自有代码采用 [MIT](LICENSE)。

## 从这里开始

**[SKILL.md](SKILL.md) 是唯一任务入口。** Codex / GPT 每次处理新文件先读它，并按任务加载下表所需文件。

| 场景 | 使用哪个文件 |
|---|---|
| 题目编号、缩进对齐、自动列表、Heading/WPS 折叠箭头 | [唯一编号规范](docs/NUMBERING_AND_OUTLINE.md) |
| 数学内容、题图、留白、分页、只交 Word | [数学资料排版工作流](docs/MATH_LAYOUT_REFERENCE.md) |
| MinerU PDF 转 DOCX | [MinerU API 说明](docs/MINERU_API.md) |
| OfficeMath 转原生 AxMath | [AxMath 转换详细工作流](docs/AXMATH_CONVERSION_WORKFLOW.md) |
| Codex 本机直连 | [Codex 入口](docs/CODEX_DIRECT.md) |
| NEXT_ACTION 指定的少数异常 | [定向修复参考](docs/REPAIR_REFERENCE.md) |

**编号硬规则：** 普通文字题号、同级对齐、跨层级保留视觉缩进；标题不要内置 Heading 身份，取消折叠不能只设 `outlineLvl=9`。样页先过 Word/WPS，再批量做全文。

**MinerU 方法差异（请按优先级执行）：**
- GPT：在本机直接使用 API。
- Codex：先用当前对话侧边栏已登录 MinerU 网页；排队超 120 秒、60 秒无进展或卡顿则切本机 API。两者均有操作权限。
- 结果都交付 `mineru_original.docx` 进入同一个 AxMath 流程，Token 留在本机忽略的 `secrets\mineru_token.txt`，不输出。

## 主要命令

    python scripts\one_click_convert.py --input "<原稿.docx>" --inspect-only
    python scripts\one_click_convert.py --input "<原稿.docx>"
    python scripts\diagnose_after_conversion.py --source "<冻结源.docx>" --working "<转换稿.docx>" --outdir "<review>"

使用 `READY_FOR_GPT_REVIEW.json` 和 `NEXT_ACTION.json` 继续任务。仅 Word 交付时禁止 PDF 渲染验收命令；已有断点应恢复，不重复转换或审计。

开发测试位于 `scripts/test_*.py`，CI 不代替 Windows Word/AxMath 真实转换与 Word/WPS 界面验收。

[变更记录](CHANGELOG.md) · [开发贡献](CONTRIBUTING.md) · [安全说明](SECURITY.md) · [历史技术记录](docs/TECHNICAL_BACKGROUND.md)（仅背景，不覆盖当前规则）
