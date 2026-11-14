# Codex 本地直连 AxMath（不经过网页电脑插件）

## MinerU PDF → DOCX：网页优先，长队列切本机 API

1. Codex 先使用**当前 Codex 对话自己的侧边栏浏览器**打开已经登录的 MinerU 网页，直接上传 PDF 并导出 DOCX；这一步不需要 API 或 Token。不要另开未登录的浏览器会话。
2. 网页预计排队超过 **120 秒**、状态连续 **60 秒没有进展**，或前端卡住时，不继续等网页。用户已明确授权 Codex 在本机运行 scripts\mineru_api.py，使用本机 secrets\mineru_token.txt 调用同一个 MinerU Precision API 作为 PDF → DOCX 备用通道。
3. GPT 在本机执行时可直接调用该 API；Codex 的备用授权仅适用于 MinerU PDF → DOCX 阶段，不延伸到 AxMath 转换阶段或其他任务。若当前任务提示词明确禁止 Token，必须遵守。
4. 两条路线都输出 mineru_original.docx。随后 Codex 只处理该 DOCX，从 AxMath 的 inspect / convert / resume / diagnose 主链开始，不再使用 MinerU Token。

Token 不得写入提示词、命令行参数、日志、Git、共享文件或输出。API 入口和故障处理见 docs\MINERU_API.md。

仓库根目录的 axmath.py 是对现有权威 Python 脚本的轻量入口。**MinerU 网页阶段使用当前 Codex 对话自己的侧边栏浏览器；DOCX 交付后，AxMath 转换阶段使用 Codex 自己运行所在 Windows 电脑的终端、Python、PowerShell、Microsoft Word COM 和 AxMath，完全不调用 thirteenthy / Qinghuan PC 1，也不需要网页插件租约。**

- 十三月电脑：D:\CHATGPT\skills\word-axmath-gpt-skill
- 清欢电脑：D:\work\axmath\word-axmath-gpt-skill（已部署本地入口）

先读本仓库 SKILL.md 和 AXMATH_CONVERSION_WORKFLOW.md 的“首轮生成硬规则”；每次任务必须先冻结源稿、完成正文标点专用处理、检查编号并创建 SOURCE_FINAL_FORMULA_MAP.json，再开始 AxMath 生成。遇到语义或几何异常时才读 REPAIR_REFERENCE.md。不能绕过 Microsoft Word / AxMath 的 Windows 本机环境；Codex 云端 Linux 环境不能执行 Word COM；64–66 个公式的 AxMath 宏内部批次限制没有改变.

## 直连命令

在相应电脑的 Skill 仓库根目录运行，也可从其他目录以 axmath.py 的绝对路径运行。以下尖括号仅为占位符，真实路径必须取自现场 JSON。

先按 SKILL.md 和 AXMATH_CONVERSION_WORKFLOW.md 的首轮生成硬规则处理冻结来源：正文句号只用专用脚本转换，题号保持普通文字和原层级，公式目标映射必须在生成前建立。转换及后续核查默认 --numbering-scope preserve；不要用它代替题号普通文字检查。下列第 5、6 步只适用于允许生成 PDF 的任务：compare 会导出临时 PDF。仅交 Word 或禁止 PDF 时跳过这两条，使用 Word-only 公式映射验收和页面视图分支完成核查。

    # 0. 原件保留；从 MinerU 输出或原始 DOCX 生成冻结副本，只替换正文普通文字中的中文句号
    python .\scripts\normalize_body_punctuation.py --source "D:\任务\mineru_original.docx" --output "D:\任务\frozen_source.docx" --report "D:\任务\review\BODY_PUNCTUATION.json"
    python .\scripts\inspect_question_numbering.py "D:\任务\frozen_source.docx"

    # 1. 冻结来源只读预检；读取 READY_FOR_GPT_REVIEW.json 中的 frozen_source 绝对路径
    python .\axmath.py inspect --input "D:\任务\frozen_source.docx"
    # 对 READY 文件记录的同一 frozen_source 建立映射；填好 expected_tex 和视觉/正文目标后再生成
    python .\scripts\math_layout_contract.py map-template --source "<READY.frozen_source>" --out "D:\任务\final-review-dir\SOURCE_FINAL_FORMULA_MAP.json"
    # 为每个 source_formula 从冻结来源填写 source_anchor 和 expected_tex；再补齐 source_formula_visual（PDF/图片源公式）和 added_body_math（正文数学对象），然后继续。

    # 2. 使用同一 READY run / frozen_source 继续正式转换；不得重新冻结成另一个路径
    python .\axmath.py convert --input "<READY.frozen_source>"

    # 已有运行目录时必须恢复，不重新开始转换
    python .\axmath.py resume --input "D:\任务\原稿.docx" --resume-run "D:\任务\run-YYYYMMDD-HHMMSS"

    # 查看转换阶段真实生成的 READY_FOR_GPT_REVIEW.json
    python .\axmath.py status --path "D:\任务\run-YYYYMMDD-HHMMSS"

    # 3. frozen_source 和 working_docx 均由 READY_FOR_GPT_REVIEW.json 提供
    python .\axmath.py diagnose --source "<frozen_source>" --working "<working_docx>" --outdir "<review-dir>"

    # 只按 NEXT_ACTION.json 指出的类定向修复
    python .\axmath.py status --path "<review-dir>"

    # 4. 对最终候选稿的全部 AxMath 做一次只读内部内容导出；空内容会阻断验收
    powershell -File .\scripts\inspect_axmath_tex.ps1 -InputDocx "<final.docx>" -AllAxMath -OutputJson "<final-review-dir>\FINAL_AXMATH_INTERNAL_TEX.json"

    # 5. 仅允许生成 PDF 时：用上一步内部审核和映射清单做最终逐页源稿比较
    python .\axmath.py compare --source "<READY.frozen_source>" --final "<final.docx>" --outdir "<final-review-dir>" --expected-source-sha256 "<frozen_sha256>" --internal-audit "<final-review-dir>\FINAL_AXMATH_INTERNAL_TEX.json" --formula-map "<final-review-dir>\SOURCE_FINAL_FORMULA_MAP.json"

    # Word-only：改用 math_layout_contract.py validate，传入同一 formula-map；不要再同时运行 compare。
    # 6. 按正文/公式锚点跨页核对，放大小型行内公式，填写 VISUAL_REVIEW.json 后确认
    python .\axmath.py finalize --report "<final-review-dir>\STRICT_FINAL_COMPARE.json" --review "<final-review-dir>\VISUAL_REVIEW.json"

支持 --help 和子命令 --help；路径支持空格、中文与相对路径（路径在子进程切换到仓库工作目录前转换为绝对路径）。十三月电脑的 Python 优先使用 D:\我的python\python3.12\python.exe。缺少依赖时请按 requirements.txt 安装。

精确重复的 OfficeMath 达到 64 个时，可选用缓存路径节省宏调用。先执行 `python .\axmath.py cache-analyze --input "<original.docx>"`；再用 `cache-prepare --input "<original.docx>" --outdir "<new-dir>"` 获得哈希绑定的 `CACHE_MANIFEST.json` 和 `prepared\<stem>_unique.docx`。对该准备稿执行上面的官方 `convert`；已有断点仍对准备稿使用 `resume`。转换 READY 生成后，执行 `python .\axmath.py cache-restore --manifest "<CACHE_MANIFEST.json>" --ready "<READY_FOR_GPT_REVIEW.json>" --output "<new-final.docx>"`。随后以缓存 manifest 的冻结原稿和恢复后的工作稿执行 `diagnose` 与符合用户交付限制的最终逐页检查；允许 PDF 时执行 `compare` 和 `finalize`。准备或恢复失败时保留现场，不把缓存结果当成验收通过。

## Codex 自动发现（可选）

普通情况下，在 Codex 任务指令中写出 SKILL.md 的真实绝对路径，它即可使用本地 CLI。需要 Codex 自动发现时，在相应本机用户目录的 .agents\skills 下创建一个指向仓库的目录联接，避免复制出两个渐渐分叉的脚本版本。

以下仅为**十三月电脑**的 PowerShell 示例，目标存在就不操作，不要覆盖旧技能目录：

    $source = "D:\CHATGPT\skills\word-axmath-gpt-skill"
    $dest = Join-Path $HOME ".agents\skills\word-axmath-conversion"
    New-Item -ItemType Directory -Force -Path (Split-Path $dest) | Out-Null
    if (-not (Test-Path $dest)) {
        New-Item -ItemType Junction -Path $dest -Target $source | Out-Null
    }

Codex 可能需要新会话才会发现新加入的 Skill。该目录联接是本机文件系统功能，不是 ChatGPT 网页连接器，也不需要 MCP/本地 HTTP 服务/API Key。

## 保留的不可绕过的约束

1. 只能在安装并可执行 Microsoft Word + AxMath 的 Windows 用户会话运行真实转换；官方转换批次限制不变。
2. inspect/convert/resume、diagnose、compare、finalize **分别**转发到既有的 one_click_convert.py、diagnose_after_conversion.py、strict_final_compare.py、finalize_visual_review.py；缓存命令只在官方转换前后处理完全相同的公式对象，不另起公式转换或验收引擎。
3. 不覆盖原始 DOCX，检查点必须可验证；禁止关闭用户原有 Word；出现错误要按冻结源稿逐式修复。
4. status 只读并摘要已有 JSON；即使静态检查通过也不能自动判定视觉验收完成，必须确实查看所有页面和公式可编辑性。
5. 新 CLI 的无 Word 单元测试：python -m unittest discover -s scripts -p "test_axmath_cli.py" -v；静态通过不代表真实 Word/AxMath 排版测试通过。
