# MinerU API 入口（GPT 专用）

## 唯一命令

从 AxMath Skill 仓库根目录运行：

    python scripts\mineru_api.py parse --input "D:\任务\原稿.pdf" --outdir "D:\任务\mineru_workspace"

使用 Precision API：本任务需要扫描数学 PDF 的公式识别和 DOCX 导出；Agent 轻量 API 只有 10 MB / 20 页限制，且只输出 Markdown，不适用。默认使用 vlm 模型，中文、OCR、公式和表格识别开启，额外导出 DOCX。每次只提交一个 PDF。任务完成后产生：

- mineru_original.docx：未经 AxMath 修改的 MinerU 原始 Word，后续转换只以它的副本为工作稿。
- mineru_result.zip：MinerU 原始完整结果。
- mineru_output\：ZIP 完整解压结果，含 Markdown、JSON、图片和布局文件。
- mineru_job.json：提交回执、batch_id、状态、源文件 SHA-256 和恢复信息。
- mineru_ready.json：给 AxMath 主流程的交接文件。

查看本地任务记录（不请求网络）：

    python scripts\mineru_api.py status --workspace "D:\任务\mineru_workspace"

如果网络中断或任务排队，重新运行同一条 parse 命令。脚本会读取 mineru_job.json，继续同一个 batch，不会重新创建任务。若上次提交请求结果不确定，会停止并要求先到 MinerU 任务列表核实，不会盲目重提。

## Token 与权限

- GPT 专属 Token 只保存在本机 secrets\mineru_token.txt，由脚本读取；不放在 SKILL.md、README、命令行、环境变量或任务 JSON。
- .gitignore 已忽略整个 secrets\ 目录。绝不提交、push、复制到共享目录或输出 Token。
- Codex 默认禁止读取 secrets\mineru_token.txt、执行 scripts\mineru_api.py 或使用这个 Token。只有用户在当前任务提示词里明确写明允许 Codex 使用本机 MinerU Token，才可在该任务范围内使用；授权不自动延续到其他任务。
- 未获授权的 Codex 应停在 MinerU 步骤，等待 GPT 提供 mineru_original.docx 和 mineru_ready.json 后再继续 AxMath 转换。

## API 固定流程

脚本复用 MinerU 官方 Precision API：

1. POST https://mineru.net/api/v4/file-urls/batch，Authorization: Bearer Token。
2. files[0] 设置 name、data_id、is_ocr=true；顶层设置 enable_formula、enable_table、language、extra_formats=[docx]。使用返回的签名 URL 对本地 PDF 执行 PUT；PUT 不附带 Authorization，也不手工设置 Content-Type。
3. GET https://mineru.net/api/v4/extract-results/batch/{batch_id} 轮询 waiting-file / pending / running / converting / done / failed。
4. done 后从 full_zip_url 下载 ZIP。下载签名 URL 时不附带 API Token。
5. 检查 ZIP 完整性和解压路径，提取 DOCX 为 mineru_original.docx。

默认限制按当前 MinerU API 文档执行：单文件不超过 200 MB、600 页；本脚本每个任务只上传一个 PDF。默认模型 vlm；如需改用 pipeline，显式加 --model-version pipeline。MinerU 文档说明 VLM 的 enable_formula 参数主要影响行内公式；最终数学语义仍必须按源 PDF 核对。

官方文档：https://mineru.net/apiManage/docs

## 按问题定位

| 现象 | 先看哪里 | 处理 |
|---|---|---|
| 找不到 Token | secrets\mineru_token.txt | 由 GPT 检查本地文件是否存在；不在聊天或日志里打印 Token。 |
| HTTP 401/403、A0202、A0211 | mineru_job.json、Token 文件 | A0202 是 Token 错误，A0211 是 Token 过期；停止重试，核对或更新 Token。 |
| submission_unknown | mineru_job.json、MinerU 任务列表 | 提交请求可能已到达服务器但回执丢失；先查任务列表，绝不直接重发。 |
| waiting-file | mineru_job.json 的 upload_attempts / last_state | 检查签名 PUT 是否成功；同一 URL 最多有限重试，不创建新 batch。 |
| pending / running / converting | mineru_job.json | 正常排队或解析；重跑同一命令继续轮询，不重新提交。 |
| -60002 | 输入文件名和扩展名 | 保留正确 .pdf 后缀，确认文件确为 PDF。 |
| -60003 / -60004 | 原始 PDF | 文件损坏或为空；先核对原件。 |
| -60005 / -60006 | PDF 文件大小、页数 | 拆成页段并使用独立工作区；每段单独提交。 |
| -500 / -10002 | scripts\mineru_api.py 的 build_payload | 请求参数或 Content-Type 不正确；按官方字段检查 payload，不猜参数。 |
| -10001 / -60001 | mineru_job.json | 服务或上传 URL 创建异常；稍后查询同一任务状态，不直接重提。 |
| -60007 / -60009 / -60010 | mineru_job.json | 模型暂不可用、队列满或解析失败；保留回执，按状态恢复或等待，不盲目重复提交。 |
| -60008 | mineru_job.json、源 PDF | 文件读取超时；核对上传是否完成和源文件可读性。 |
| -60011 | mineru_job.json 的 batch_id / upload_attempts | 未获取到有效上传文件；检查签名 PUT 是否成功。 |
| -60012 / -60013 | mineru_job.json 的 batch_id | 任务不存在或无权限；确认 batch_id 和创建任务的账号/Token。 |
| -60015 / -60016 | mineru_result.zip、mineru_job.json | 文件转换失败；保留原始 ZIP 和回执，先核对 extra_formats，不要重做 OCR。 |
| -60018 | MinerU 账号任务额度 | 当日任务数量达到上限；等额度恢复后再用新工作区提交。 |
| ZIP 无 DOCX | mineru_result.zip、mineru_output\、mineru_job.json | 检查 extra_formats 是否包含 docx；保留 ZIP，不把 Markdown 改名成 DOCX。 |
| AxMath 转换/公式错误 | SKILL.md、docs\REPAIR_REFERENCE.md、NEXT_ACTION.json | MinerU 阶段已经结束；按 AxMath 诊断类别定向处理，不回头重复 OCR。 |

## 与 AxMath 主链的交接

MinerU 完成后，将 mineru_ready.json 中的 mineru_docx 作为输入，先执行 AxMath Skill 的 DOCX 预检：

    python scripts\one_click_convert.py --input "D:\任务\mineru_workspace\mineru_original.docx" --inspect-only

预检通过后只运行一次正式转换；已有 READY_FOR_GPT_REVIEW.json 时按 SKILL.md 的 resume / diagnose 路线继续。不要把 PDF 直接传给 one_click_convert.py，它是 DOCX → AxMath 转换入口，不是 MinerU 上传器。
