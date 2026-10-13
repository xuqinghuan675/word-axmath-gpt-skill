# Codex 本地直连 AxMath（不经过网页电脑插件）

仓库根目录的 axmath.py 是对现有权威 Python 脚本的轻量入口。**Codex 使用自己运行所在 Windows 电脑的终端、Python、PowerShell、Microsoft Word COM 和 AxMath，完全不调用 thirteenthy / Qinghuan PC 1**，也不需要网页 ChatGPT 连接器的租约。

- 十三月电脑：D:\CHATGPT\skills\word-axmath-gpt-skill
- 清欢电脑：D:\work\axmath\word-axmath-gpt-skill（已部署本地入口）

请先读取本仓库 SKILL.md，遇到语义或几何异常时才读 docs/REPAIR_REFERENCE.md。不能绕过 Microsoft Word / AxMath 的 Windows 本机环境；Codex 云端 Linux 环境不能执行 Word COM；64–66 个公式的 AxMath 宏内部批次限制没有改变。

## 直连命令

在相应电脑的 Skill 仓库根目录运行，也可从其他目录以 axmath.py 的绝对路径运行。以下尖括号仅为占位符，真实路径必须取自现场 JSON。

    # 1. 原稿只读预检
    python .\axmath.py inspect --input "D:\任务\原稿.docx"

    # 2. 正式转换：只转换一次
    python .\axmath.py convert --input "D:\任务\原稿.docx"

    # 已有运行目录时必须恢复，不重新开始转换
    python .\axmath.py resume --input "D:\任务\原稿.docx" --resume-run "D:\任务\run-YYYYMMDD-HHMMSS"

    # 查看转换阶段真实生成的 READY_FOR_GPT_REVIEW.json
    python .\axmath.py status --path "D:\任务\run-YYYYMMDD-HHMMSS"

    # 3. frozen_source 和 working_docx 均由 READY_FOR_GPT_REVIEW.json 提供
    python .\axmath.py diagnose --source "<frozen_source>" --working "<working_docx>" --outdir "<review-dir>"

    # 只按 NEXT_ACTION.json 指出的类定向修复
    python .\axmath.py status --path "<review-dir>"

    # 4. 诊断所有阻塞均清零后，一次性做最终完整逐页源稿比较
    python .\axmath.py compare --source "<frozen_source>" --final "<final.docx>" --outdir "<final-review-dir>" --expected-source-sha256 "<frozen_sha256>"

    # 5. 逐页实际查看、填写 VISUAL_REVIEW.json 后方可确认
    python .\axmath.py finalize --report "<final-review-dir>\STRICT_FINAL_COMPARE.json" --review "<final-review-dir>\VISUAL_REVIEW.json"

支持 --help 和子命令 --help；路径支持空格、中文与相对路径（路径在子进程切换到仓库工作目录前转换为绝对路径）。十三月电脑的 Python 优先使用 D:\我的python\python3.12\python.exe。缺少依赖时请按 requirements.txt 安装。

精确重复的 OfficeMath 达到 64 个时，可选用缓存路径节省宏调用。先执行 `python .\axmath.py cache-analyze --input "<original.docx>"`；再用 `cache-prepare --input "<original.docx>" --outdir "<new-dir>"` 获得哈希绑定的 `CACHE_MANIFEST.json` 和 `prepared\<stem>_unique.docx`。对该准备稿执行上面的官方 `convert`；已有断点仍对准备稿使用 `resume`。转换 READY 生成后，执行 `python .\axmath.py cache-restore --manifest "<CACHE_MANIFEST.json>" --ready "<READY_FOR_GPT_REVIEW.json>" --output "<new-final.docx>"`。随后以缓存 manifest 的冻结原稿和恢复后的工作稿执行同一套 `diagnose`、`compare`、逐页查看与 `finalize`。准备或恢复失败时保留现场，不把缓存结果当成验收通过。

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
