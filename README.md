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

## 为什么大文档转换会慢

这里的主要瓶颈不是 Python，也不是 GPT 在“等”。AxMath 2.7.0.58 的 Word 插件会自己把 `AMSMML2AM` 转换拆成约 **64~66 个公式/批**，前端提示其目的就是避免 Word 无响应；脚本只是识别每批完成、保存进度，再继续下一批。

已经有两组实测：

- 早期 768 公式文档：官方单次全选只转换 66 个；自动循环 12 批才完成。
- 1335 公式正式文档：21 批完成，`AMSMML2AM` 宏本身累计约 **6718.6 秒 / 总 6785.0 秒 = 99.02%**，约 **5.03 秒/公式**。

因此不能靠“把 66 改成 200”“绕过 `WaitingConvert`”“并行点同一份 Word”“去掉每批保存”来提速。历史大文档转换还出现过 AxMath.exe crash dump，而批次保存让任务可以把已经完成的批次落盘；这些是稳定性设计，不是多余等待。

当前转换报告会记录每批宏耗时、保存耗时、新 crash dump，以及宏耗时占比。真正要继续提速，应先用这些数据定位 AxMath 插件内部的慢批次；在没有新的可复现证据前，不改官方批处理语义。

修复阶段也避免重复做完整文档渲染：中间轮次优先 geometry audit + 异常公式/异常页定向检查；完整 source/final fresh render + 每页并排图保留给最终验收。

中间 source 几何证据可直接使用：

```powershell
python scripts\snapshot_docx.py "<frozen-source.docx>" "<review\source_geometry>" --profile geometry
```

`geometry` profile 仍读取真实 Word COM 公式位置/段落信息并做 DOCX 前后 SHA 校验，但跳过 PDF/逐页 PNG、1335 个公式 crop 和 preview 媒体提取。`snapshot_docx.py` 默认仍是 `full`；`strict_final_compare.py` 不传 profile，因此最终全页验收行为不变。

## 安全规则

- 不覆盖原始 DOCX；所有写入脚本都拒绝 input=output，已有输出默认也不覆盖
- 已有用户 Word 不是自动停止条件：记录现有 PID，使用独立的任务自有 Word 会话继续；只能关闭/杀掉任务自有 Word，绝不能碰用户已有 Word
- Word 锁文件只作为状态记录，不再单独阻断；只要磁盘上的 source 可读，就冻结副本并用前后 SHA 校验，若运行期间 source 发生变化则验收失败
- 前台/后台切换、保存/另存为、任务自身弹窗、重试、检测/修复、逐页验收和任务进程清理由 GPT 自己完成，不因需要 UI 操作把步骤甩回用户
- frozen source 即使只读也不改原稿；复制出的 working copy 会清除继承的只读属性，避免首个 AxMath batch 后触发“另存为”
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
- `scripts/repair_axmath_inline_roundtrip.ps1`：行内/display 局部修复；只处理 GPT 已确认的 culprit，单次最多 12 个，禁止把 broken same-line group 整组无脑 roundtrip
- `scripts/rebuild_axmath_baselines.ps1`：Class B 内部指标**局部 probe**（每次最多 3 个明确 ordinal；禁止批量重建）
- `scripts/export_source_word_latex.ps1`：从 frozen source 只读导出指定 ordinal 的 Word LaTeX，带 source SHA 前后校验
- `scripts/normalize_axmath_tex.py`：把 Word/Unicode 的 prime 形式按 AxMath 2.7.0.58 已实机验证的内置语法规范化：一阶 `\prime`、二阶 `''`、三阶 `'''`；同时处理 `y′² → {y\prime}^{2}` 这类 prime+上标结合
- `scripts/probe_axmath_prime_contract.ps1`：在真实 Word + AxMath 上复测一/二/三阶 prime 回转，并确认 raw Unicode `‴` 仍是不安全 donor；用于安装版本变化后的再校验
- `scripts/repair_axmath_from_approved_tex.ps1`：把 GPT 已批准的 `{ordinal, tex}` map 写回新的 working 副本；不猜公式语义，逐个保存并保持 AxMath/OfficeMath/paragraph count
- `scripts/calibrate_axmath_boxes.py`：Word 外部 OLE 框校准
- `scripts/strict_final_compare.py`：最终 source-vs-final 每页并排图 + 内容/结构诊断 + 视觉复核模板
- `scripts/finalize_visual_review.py`：校验逐页视觉复核清单与文件/图片哈希，只有它输出 `acceptance_pass=true` 才算最终通过
- `scripts/repo_selfcheck.py`：跨平台静态自检，检查 Python 语法和关键安全/验收契约

语义损坏或直接 OMML→AxMath donor 已经塌缩时，优先从 frozen source 的同 ordinal 公式重建：Word 导出 LaTeX → `normalize_axmath_tex.py` 归一化 → GPT 复核 → AxMath `AMSTeX2AM`。其中 prime/导数符号是独立语义风险：`formula_geometry_audit.py` 会对所有含 prime 的源公式生成 `prime_semantic_candidates`，不再要求先出现 same-line/尺寸异常。徐清欢电脑上的 AxMath 2.7.0.58 已确认：`\prime`、`''`、`'''` 分别作为一/二/三阶 prime 的内置解析语法，并能在 `AMSTeX2AM → AMSAM2TeX` 回转中保留阶数；直接把 Unicode `‴` 喂给 `AMSTeX2AM` 会回转为 `?`，因此 raw Unicode prime 只作源证据，不作 donor TeX。源公式真实跨多行时保留 source-derived line breaks，用 `aligned` 类结构重建，不靠缩小 OLE 外框硬塞回一行。

`formula_geometry_audit.py` 会额外列出三种 GPT review/routing queue：`semantic_rebuild_candidates` 用于“非平凡 source 表达式却落进 tiny AxMath shell”的疑似语义塌缩；`prime_semantic_candidates` 对所有 prime/导数源公式独立生效；`roundtrip_semantic_risk_candidates` 则标记 broken same-line group 中不应走 AxMath→TeX roundtrip 的 prime 公式。历史正式文档已证明 `AMSAM2TeX` 可能把 `f′(x)` 静默变成 `f`，所以不能再用“宏调用成功”代替语义验收。

AxMath 的 `AxMath.dotm` 会从常见 Program Files 位置自动寻找；如果安装在其他位置，GPT 可以给 PowerShell 脚本显式传 `-TemplatePath`。

---

这是简化后的生产版，不包含旧的 Web-GPT/Watchdog 自动唤醒流程。
