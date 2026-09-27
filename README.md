# Word → AxMath GPT Skill

给 **Windows + Microsoft Word + 已激活 AxMath** 用的 GPT Skill。

用途很简单：让 GPT 把 Word 里的 OfficeMath 批量转换成真正可编辑的 AxMath，然后按照原稿检查并修复行内/行间、大小、换行、外框和预览异常。

## 宝宝最简单的用法

### 1. 下载/克隆仓库

```powershell
git clone https://github.com/xuqinghuan675/word-axmath-gpt-skill.git
cd word-axmath-gpt-skill
python -m pip install -r requirements.txt
```

### 2. 直接告诉 GPT

把 Word 文件路径给 GPT，然后说：

> 先读取这个仓库的 SKILL.md。把这个 Word 的 OfficeMath 转成 AxMath，原文件不要改；先按 Skill 在旁边建立工作区并一键转换，转换完成后继续按 Skill 对照原稿排查和修复。最后必须 fresh Word/PDF 严格逐页对照原稿做视觉比对，不能只抽查；原稿居中的公式最终也必须保持居中。

就可以了。

不需要自己研究脚本，也不需要自动打开新的 GPT 对话。

## 手动一键转换（可选）

```powershell
python scripts\one_click_convert.py --input "C:\path\to\document.docx"
```

程序会先检查文件状态，然后在原文件旁边创建：

```text
document_AxMath-workspace\
  run-YYYYMMDD-HHMMSS\
    SOURCE_STATE.json
    READY_FOR_GPT_REVIEW.json
    document\
      source\
      working\
      logs\
```

转换完成后，让 GPT 读取 `READY_FOR_GPT_REVIEW.json` 和 `SKILL.md` 继续检查即可。

## 环境要求

- Windows
- Microsoft Word
- AxMath 已安装并激活
- Python 3.10+
- GPT/Agent 能访问本机文件和运行命令

Python 依赖：

```powershell
python -m pip install -r requirements.txt
```

## 这套流程解决什么

- OfficeMath → 真正的 `Equation.AxMath` 可编辑对象
- 批量转换
- 源文件只读保护
- 行内公式被错误转成巨大 display 公式
- 转换后公式掉到下一行
- AxMath 内部类型/尺寸异常
- Word OLE 外框 width/height/baseline 异常
- preview/OLE 缓存问题
- 原稿居中的公式转换后偏左/偏右
- 最后 fresh Word/PDF 与原稿 **逐页严格视觉比对**，不是抽查
- 最终视觉通过会绑定 source/final DOCX 哈希和每页并排图哈希；文件之后变化，旧视觉通过自动失效

其中已经验证过的一条关键修复是：

> AxMath → `AMSAM2TeX` → 强制单 `$...$` 行内 LaTeX → `AMSTeX2AM`

用于修复“原稿明明是行内公式，批量转换后却变成巨大行间公式”的情况。

## 安全规则

- 不覆盖原始 DOCX；所有写入脚本都拒绝 input=output，已有输出默认也不覆盖
- 如果已经有 Word 在运行，后台自动转换直接停止，不碰现有 Word
- 只清理由本任务自己创建的 Word 进程
- 不靠固定宽度阈值判断公式是否正确
- 不自动打开/新建 GPT 网页对话
- 不包含任何用户文档、日志、账号、凭据或个人路径

## 主要文件

- `SKILL.md`：GPT 必读规则
- `scripts/one_click_convert.py`：预检 + 邻接工作区 + 一键转换
- `scripts/run_skill.py`：受保护的转换执行器
- `scripts/snapshot_docx.py`：原稿/改稿取证
- `scripts/formula_geometry_audit.py`：公式几何与 inline 意图检查
- `scripts/repair_axmath_inline_roundtrip.ps1`：行内/display 类型修复
- `scripts/rebuild_axmath_baselines.ps1`：剩余 AxMath 内部指标重建
- `scripts/calibrate_axmath_boxes.py`：Word 外部 OLE 框校准
- `scripts/strict_final_compare.py`：最终 source-vs-final 每页并排图 + 内容/结构诊断 + 视觉复核模板
- `scripts/finalize_visual_review.py`：校验逐页视觉复核清单与文件/图片哈希，只有它输出 `acceptance_pass=true` 才算最终通过
- `scripts/repo_selfcheck.py`：跨平台静态自检，检查 Python 语法和关键安全/验收契约

AxMath 的 `AxMath.dotm` 会从常见 Program Files 位置自动寻找；如果安装在其他位置，GPT 可以给 PowerShell 脚本显式传 `-TemplatePath`。

---

这是简化后的生产版，不包含旧的 Web-GPT/Watchdog 自动唤醒流程。
