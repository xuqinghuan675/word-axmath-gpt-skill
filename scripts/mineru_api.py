from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
import shutil
import stat
import sys
import time
import zipfile
from pathlib import Path, PurePosixPath
from urllib.parse import quote

import pymupdf as fitz
import requests

ROOT = Path(__file__).resolve().parents[1]
TOKEN_PATH = ROOT / "secrets" / "mineru_token.txt"
API_ROOT = "https://mineru.net/api/v4"
MAX_FILE_BYTES = 200 * 1024 * 1024
MAX_PAGES = 600


class MinerUError(RuntimeError):
    pass


def now():
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_json(path: Path, payload: dict):
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + ".tmp")
    temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temp, path)


def read_json(path: Path) -> dict:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise MinerUError(f"无法读取任务记录：{path.name}（{type(exc).__name__}）") from None
    if not isinstance(value, dict):
        raise MinerUError(f"任务记录格式错误：{path.name}")
    return value


def load_token() -> str:
    if not TOKEN_PATH.is_file():
        raise MinerUError("未找到 GPT 专属 MinerU Token 文件：" + str(TOKEN_PATH))
    token = TOKEN_PATH.read_text(encoding="utf-8-sig").strip()
    if not token or any(ch.isspace() for ch in token):
        raise MinerUError("GPT 专属 MinerU Token 文件为空或格式异常。")
    return token


def validate_pdf(path: Path) -> tuple[str, int, int]:
    if not path.is_file():
        raise MinerUError(f"PDF 不存在：{path}")
    size = path.stat().st_size
    if size <= 0:
        raise MinerUError("PDF 是空文件。")
    if size > MAX_FILE_BYTES:
        raise MinerUError("PDF 超过 MinerU API 的 200 MB 单文件上限；先拆分到新工作区。")
    try:
        with fitz.open(path) as pdf:
            pages = len(pdf)
            if pdf.is_encrypted:
                raise MinerUError("PDF 已加密，MinerU 无法直接解析。")
    except MinerUError:
        raise
    except Exception as exc:
        raise MinerUError(f"PDF 无法读取：{type(exc).__name__}") from None
    if pages <= 0:
        raise MinerUError("PDF 没有可解析页面。")
    if pages > MAX_PAGES:
        raise MinerUError(
            f"PDF 共 {pages} 页，超过当前 MinerU API 文档的 200 页单文件限制；"
            "按页段拆分到独立工作区，不要重复提交整份文件。"
        )
    return sha256_file(path), size, pages


def preserve_source(pdf: Path, workspace: Path, source_hash: str):
    target = workspace / pdf.name
    if target.resolve() == pdf.resolve():
        return target
    if target.exists():
        if sha256_file(target) != source_hash:
            raise MinerUError(f"工作区内已有同名但哈希不同的源 PDF，拒绝覆盖：{target}")
        return target
    shutil.copy2(pdf, target)
    if sha256_file(target) != source_hash:
        raise MinerUError("复制源 PDF 后 SHA-256 不一致，停止提交。")
    return target


def api_json(session: requests.Session, method: str, url: str, *, body=None, timeout=(20, 60)) -> dict:
    try:
        response = session.request(method, url, json=body, timeout=timeout)
    except requests.RequestException as exc:
        raise MinerUError(
            f"MinerU API 网络请求失败（{type(exc).__name__}）。"
            "若是提交请求超时，先检查 mineru_job.json，不要盲目重新提交。"
        ) from None
    try:
        payload = response.json()
    except ValueError:
        payload = {}
    code = payload.get("code") if isinstance(payload, dict) else None
    message = (payload.get("msg") or payload.get("message") or response.reason) if isinstance(payload, dict) else response.reason
    if response.status_code in (401, 403) or str(code) in ("A0202", "A0211"):
        raise MinerUError(
            f"MinerU 认证失败：HTTP {response.status_code} / {code}。"
            "A0202 表示 Token 错误，A0211 表示 Token 过期；停止重试并核对本机凭据。"
        )
    if not response.ok:
        raise MinerUError(f"MinerU API 请求失败：HTTP {response.status_code} / {code} / {message}")
    if not isinstance(payload, dict) or str(code) != "0":
        raise MinerUError(f"MinerU API 返回业务错误：{code} / {message}")
    return payload


def build_payload(pdf: Path, data_id: str, model_version: str) -> dict:
    return {
        "files": [{"name": pdf.name, "data_id": data_id, "is_ocr": True}],
        "model_version": model_version,
        "language": "ch",
        "enable_formula": True,
        "enable_table": True,
        "extra_formats": ["docx"],
    }


def save_manifest(path: Path, manifest: dict, **updates):
    manifest.update(updates)
    manifest["updated_at"] = now()
    write_json(path, manifest)


def submit_task(pdf: Path, manifest_path: Path, manifest: dict, session: requests.Session):
    # Persist before POST. If the response is lost, never create a duplicate task automatically.
    save_manifest(manifest_path, manifest, state="submitting", submit_started_at=now())
    payload = build_payload(pdf, manifest["data_id"], manifest["model_version"])
    try:
        response = session.post(API_ROOT + "/file-urls/batch", json=payload, timeout=(20, 60))
    except requests.RequestException as exc:
        save_manifest(manifest_path, manifest, state="submission_unknown",
                      last_error_type=type(exc).__name__)
        raise MinerUError(
            "MinerU 提交请求结果不确定，任务记录已保留。先到 MinerU 任务列表核实，"
            "不要直接重发；确认没有任务后再用全新的工作区提交。"
        ) from None
    try:
        result = response.json()
    except ValueError:
        result = {}
    code = result.get("code") if isinstance(result, dict) else None
    message = (result.get("msg") or result.get("message") or response.reason) if isinstance(result, dict) else response.reason
    if response.status_code in (401, 403) or str(code) in ("A0202", "A0211"):
        save_manifest(manifest_path, manifest, state="submit_failed", api_code=code)
        raise MinerUError(f"MinerU 认证失败：HTTP {response.status_code} / {code}。检查本机 GPT 专属 Token 文件。")
    if not response.ok or not isinstance(result, dict) or str(code) != "0":
        save_manifest(manifest_path, manifest, state="submit_failed",
                      http_status=response.status_code, api_code=code, api_message=str(message)[:500])
        raise MinerUError(f"MinerU 创建上传任务失败：HTTP {response.status_code} / {code} / {message}")
    data = result.get("data") or {}
    batch_id = data.get("batch_id")
    urls = data.get("file_urls") or []
    if not batch_id or len(urls) != 1 or not urls[0]:
        save_manifest(manifest_path, manifest, state="submit_failed", api_code=code)
        raise MinerUError("MinerU 返回的 batch_id 或上传 URL 数量不符合单文件任务预期。")
    save_manifest(manifest_path, manifest, batch_id=batch_id, upload_url=urls[0],
                  upload_attempts=0, state="upload_pending", trace_id=result.get("trace_id"))


def upload_file(pdf: Path, manifest_path: Path, manifest):
    upload_url = manifest.get("upload_url")
    if not upload_url:
        raise MinerUError("任务记录缺少签名上传 URL；保留现场，不重新创建任务。")
    attempts = int(manifest.get("upload_attempts") or 0)
    if attempts >= 2:
        raise MinerUError(
            "同一 MinerU 上传 URL 已尝试两次仍未成功。检查 mineru_job.json 和 MinerU 任务列表；"
            "不要在原工作区盲目重提。"
        )
    attempts += 1
    save_manifest(manifest_path, manifest, state="uploading", upload_attempts=attempts)
    upload_session = requests.Session()  # Deliberately no Authorization header on signed OSS PUT.
    try:
        with pdf.open("rb") as stream:
            response = upload_session.put(upload_url, data=stream, timeout=(30, 600))
    except requests.RequestException as exc:
        save_manifest(manifest_path, manifest, state="upload_failed",
                      last_upload_error_type=type(exc).__name__)
        raise MinerUError(
            f"MinerU 文件上传失败（{type(exc).__name__}）。签名 URL 已保存在本地任务记录；"
            "再次运行同一命令会恢复同一个 batch，不会新建任务。"
        ) from None
    finally:
        upload_session.close()
    if not response.ok:
        save_manifest(manifest_path, manifest, state="upload_failed",
                      last_upload_http_status=response.status_code)
        raise MinerUError(
            f"MinerU 文件上传失败：HTTP {response.status_code}。"
            "签名 URL 有效期为 24 小时；再次运行同一命令会恢复同一任务。"
        )
    save_manifest(manifest_path, manifest, state="uploaded", upload_url=None, uploaded_at=now())


def query_batch(session: requests.Session, batch_id: str):
    url = API_ROOT + "/extract-results/batch/" + quote(batch_id, safe="")
    payload = api_json(session, "GET", url, timeout=(20, 60))
    data = payload.get("data") or {}
    rows = data.get("extract_result") or []
    if isinstance(rows, dict):
        rows = [rows]
    if not isinstance(rows, list) or not rows:
        raise MinerUError("MinerU 批量状态接口没有返回 extract_result。")
    return data, rows


def find_file_result(rows: list, manifest: dict) -> dict:
    for row in rows:
        if row.get("data_id") == manifest.get("data_id"):
            return row
        if row.get("file_name") == manifest.get("source_name"):
            return row
    if len(rows) == 1:
        return rows[0]
    raise MinerUError("MinerU 批量结果中找不到本次 PDF 对应的 data_id / file_name。")


def poll_until_done(session: requests.Session, manifest_path: Path, manifest: dict,
                    poll_interval: int, max_wait_seconds: int):
    deadline = time.monotonic() + max_wait_seconds
    last_state = None
    while True:
        data, rows = query_batch(session, manifest["batch_id"])
        row = find_file_result(rows, manifest)
        state = str(row.get("state") or "unknown")
        progress = row.get("extract_progress") or {}
        if state != last_state:
            print(f"MinerU 状态：{state}", flush=True)
            last_state = state
        save_manifest(
            manifest_path, manifest,
            state="polling" if state not in ("done", "failed") else state,
            last_state=state,
            last_progress=progress,
            last_error=str(row.get("err_msg") or "")[:500] or None,
        )
        if state == "done":
            zip_url = row.get("full_zip_url")
            if not zip_url:
                save_manifest(manifest_path, manifest, state="polling", last_state="done",
                              last_error="done response did not include full_zip_url")
                raise MinerUError("MinerU 状态为 done，但没有 full_zip_url；任务记录保留为 polling，可稍后恢复查询。")
            save_manifest(manifest_path, manifest, state="done", full_zip_url=zip_url)
            return
        if state == "failed":
            raise MinerUError(
                "MinerU 解析失败：" + str(row.get("err_msg") or "未提供错误原因")
                + "。保留 mineru_job.json，不自动创建新任务。"
            )
        if state not in ("waiting-file", "pending", "running", "converting"):
            raise MinerUError(f"MinerU 返回未知任务状态：{state}。查看 mineru_job.json。")
        if time.monotonic() >= deadline:
            raise MinerUError(
                "MinerU 任务仍在队列或解析中。任务 batch_id 已保存；"
                "再次运行同一命令会继续轮询，不会重复提交。"
            )
        time.sleep(poll_interval)


def recover_interrupted_upload(session: requests.Session, manifest_path: Path, manifest: dict,
                               poll_interval: int):
    deadline = time.monotonic() + min(60, max(15, poll_interval * 12))
    while True:
        data, rows = query_batch(session, manifest["batch_id"])
        row = find_file_result(rows, manifest)
        state = str(row.get("state") or "unknown")
        if state != "waiting-file":
            save_manifest(manifest_path, manifest, state="uploaded",
                          upload_url=None, last_state=state)
            return
        if time.monotonic() >= deadline:
            upload_file(Path(manifest["source"]), manifest_path, manifest)
            return
        time.sleep(poll_interval)


def download_zip(manifest_path: Path, manifest: dict, workspace: Path) -> Path:
    archive = workspace / "mineru_result.zip"
    if archive.is_file() and zipfile.is_zipfile(archive):
        with zipfile.ZipFile(archive) as zf:
            bad = zf.testzip()
        if bad is None:
            return archive
    url = manifest.get("full_zip_url")
    if not url:
        raise MinerUError("任务已完成，但任务记录没有 full_zip_url。")
    temp = workspace / "mineru_result.zip.part"
    download_session = requests.Session()  # No API token on the signed result URL.
    try:
        response = download_session.get(url, stream=True, timeout=(30, 180))
        if not response.ok:
            raise MinerUError(f"MinerU 结果 ZIP 下载失败：HTTP {response.status_code}。")
        with temp.open("wb") as target:
            for chunk in response.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    target.write(chunk)
    except MinerUError:
        raise
    except requests.RequestException as exc:
        raise MinerUError(f"MinerU 结果 ZIP 下载失败（{type(exc).__name__}）；任务记录已保留。") from None
    finally:
        download_session.close()
    if not zipfile.is_zipfile(temp):
        temp.unlink(missing_ok=True)
        raise MinerUError("MinerU 返回的结果不是有效 ZIP；任务记录已保留，稍后可继续下载。")
    with zipfile.ZipFile(temp) as zf:
        bad = zf.testzip()
    if bad is not None:
        temp.unlink(missing_ok=True)
        raise MinerUError(f"MinerU ZIP 完整性检查失败：{bad}")
    os.replace(temp, archive)
    save_manifest(manifest_path, manifest, zip_sha256=sha256_file(archive))
    return archive


def safe_extract_zip(archive: Path, output_dir: Path):
    if output_dir.exists():
        raise MinerUError(f"原始解压目录已存在，拒绝覆盖：{output_dir}。检查目录后再决定如何继续。")
    temp_dir = output_dir.with_name(output_dir.name + ".tmp")
    if temp_dir.exists():
        raise MinerUError(f"检测到未完成的临时解压目录：{temp_dir}；先检查现场，不自动删除。")
    temp_dir.mkdir(parents=True)
    with zipfile.ZipFile(archive) as zf:
        for info in zf.infolist():
            name = info.filename.replace("\\", "/")
            rel = PurePosixPath(name)
            if rel.is_absolute() or not rel.parts or any(part in ("", ".", "..") for part in rel.parts):
                raise MinerUError(f"ZIP 含不安全路径，拒绝解压：{info.filename}")
            mode = info.external_attr >> 16
            if stat.S_ISLNK(mode):
                raise MinerUError(f"ZIP 含符号链接，拒绝解压：{info.filename}")
            target = temp_dir.joinpath(*rel.parts)
            resolved = target.resolve()
            if temp_dir.resolve() not in resolved.parents and resolved != temp_dir.resolve():
                raise MinerUError(f"ZIP 路径越界，拒绝解压：{info.filename}")
            if info.is_dir():
                target.mkdir(parents=True, exist_ok=True)
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, target.open("wb") as dst:
                shutil.copyfileobj(src, dst)
    os.replace(temp_dir, output_dir)


def choose_docx(output_dir: Path) -> Path:
    files = sorted(output_dir.rglob("*.docx"))
    if not files:
        raise MinerUError(
            "结果 ZIP 中没有 DOCX。原始 ZIP 和 mineru_output 已保留；"
            "检查任务是否请求 extra_formats=['docx']，不要把 Markdown 冒充 Word。"
        )
    full = [p for p in files if p.name.lower() == "full.docx"]
    if len(full) == 1:
        return full[0]
    if len(files) == 1:
        return files[0]
    raise MinerUError(
        "结果 ZIP 含多个 DOCX，无法自动判断主文件：" + ", ".join(p.name for p in files)
        + "。保留原始输出，需人工指定。"
    )


def finish_outputs(workspace: Path, manifest_path: Path, manifest: dict):
    archive = download_zip(manifest_path, manifest, workspace)
    output_dir = workspace / "mineru_output"
    if not output_dir.exists():
        safe_extract_zip(archive, output_dir)
    docx_source = choose_docx(output_dir)
    raw_docx = workspace / "mineru_original.docx"
    if raw_docx.exists():
        if sha256_file(raw_docx) != sha256_file(docx_source):
            raise MinerUError("mineru_original.docx 已存在但哈希不同，拒绝覆盖。")
    else:
        shutil.copy2(docx_source, raw_docx)
    ready = {
        "status": "ready_for_axmath",
        "source_pdf": manifest["source"],
        "source_pdf_sha256": manifest["source_sha256"],
        "mineru_docx": str(raw_docx.resolve()),
        "mineru_docx_sha256": sha256_file(raw_docx),
        "raw_zip": str(archive.resolve()),
        "raw_output_dir": str(output_dir.resolve()),
        "batch_id": manifest["batch_id"],
        "model_version": manifest["model_version"],
        "completed_at": now(),
    }
    write_json(workspace / "mineru_ready.json", ready)
    save_manifest(manifest_path, manifest, state="completed", mineru_docx=str(raw_docx.resolve()),
                  mineru_docx_sha256=ready["mineru_docx_sha256"])
    print("MinerU 完成。")
    print("原始 MinerU DOCX：", raw_docx)
    print("原始 ZIP：", archive)
    print("解压结果：", output_dir)
    print("交接记录：", workspace / "mineru_ready.json")


def run_parse(args) -> int:
    pdf = Path(args.input).expanduser().resolve()
    workspace = Path(args.outdir).expanduser().resolve()
    workspace.mkdir(parents=True, exist_ok=True)
    manifest_path = workspace / "mineru_job.json"
    source_hash, source_size, pages = validate_pdf(pdf)
    preserve_source(pdf, workspace, source_hash)

    if manifest_path.exists():
        manifest = read_json(manifest_path)
        if manifest.get("source_sha256") != source_hash:
            raise MinerUError(
                "该工作区已有其他 PDF 的 MinerU 任务记录，哈希不匹配；"
                "不要覆盖或复用，给新 PDF 使用新的工作区。"
            )
        state = manifest.get("state")
        if state in ("submission_unknown", "submitting"):
            raise MinerUError(
                "上次创建任务时请求结果不确定。先到 MinerU 任务列表核实，"
                "不要重新提交；确认没有任务后再用全新的工作区。"
            )
        if state in ("submit_failed", "failed"):
            raise MinerUError(
                "已有任务明确失败。先读取 mineru_job.json 的错误信息；"
                "确认原因并决定新任务后，使用新的工作区，不在原记录上覆盖。"
            )
        if state == "completed" and (workspace / "mineru_original.docx").is_file():
            print("MinerU 已完成，复用现有结果：", workspace / "mineru_original.docx")
            return 0
    else:
        data_id = source_hash[:32]
        manifest = {
            "schema": "mineru-precision-api-job/v1",
            "source": str(pdf),
            "source_name": pdf.name,
            "source_sha256": source_hash,
            "source_bytes": source_size,
            "source_pages": pages,
            "data_id": data_id,
            "model_version": args.model_version,
            "extra_formats": ["docx"],
            "state": "created",
            "created_at": now(),
            "updated_at": now(),
        }
        write_json(manifest_path, manifest)
        state = "created"

    # state=created means the process stopped before the POST began.
    # submit_task writes state=submitting before network I/O; that state is never resubmitted.
    if state == "created":
        token = load_token()
        api = requests.Session()
        api.headers.update({"Authorization": "Bearer " + token, "Accept": "application/json"})
        try:
            submit_task(pdf, manifest_path, manifest, api)
        finally:
            api.close()
        token = None
        state = manifest.get("state")

    # Once the batch is done, the signed result URL is sufficient; no Token is needed to finish download.
    if state in ("upload_pending", "upload_failed", "uploading", "uploaded", "polling"):
        token = load_token()
        api = requests.Session()
        api.headers.update({"Authorization": "Bearer " + token, "Accept": "application/json"})
        try:
            if state in ("upload_pending", "upload_failed"):
                upload_file(pdf, manifest_path, manifest)
            elif state == "uploading":
                recover_interrupted_upload(api, manifest_path, manifest, args.poll_interval)
            if manifest.get("state") not in ("done", "completed"):
                poll_until_done(api, manifest_path, manifest, args.poll_interval, args.max_wait_seconds)
        finally:
            api.close()
            token = None

    if manifest.get("state") not in ("done", "completed"):
        raise MinerUError(f"任务记录处于未处理状态：{manifest.get('state')}。查看 mineru_job.json。")
    finish_outputs(workspace, manifest_path, manifest)
    return 0

def run_status(args) -> int:
    workspace = Path(args.workspace).expanduser().resolve()
    path = workspace / "mineru_job.json"
    if not path.is_file():
        raise MinerUError(f"找不到 MinerU 任务记录：{path}")
    manifest = read_json(path)
    summary = {
        key: manifest.get(key)
        for key in (
            "state", "source", "source_sha256", "source_pages", "model_version",
            "batch_id", "last_state", "last_progress", "last_error",
            "mineru_docx", "mineru_docx_sha256", "updated_at",
        )
        if key in manifest
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


def make_parser():
    parser = argparse.ArgumentParser(
        description="GPT 专属 MinerU Precision API：本地 PDF → 原始可编辑 DOCX + 完整 ZIP。"
    )
    sub = parser.add_subparsers(dest="command", required=True)
    parse = sub.add_parser("parse", help="首次提交或恢复同一 MinerU 任务")
    parse.add_argument("--input", required=True, help="源 PDF")
    parse.add_argument("--outdir", required=True, help="独立任务工作区")
    parse.add_argument("--model-version", choices=("vlm", "pipeline"), default="vlm")
    parse.add_argument("--poll-interval", type=int, default=5)
    parse.add_argument("--max-wait-seconds", type=int, default=3600)
    status = sub.add_parser("status", help="只读查看本地任务记录，不调用 API")
    status.add_argument("--workspace", required=True)
    return parser


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    args = make_parser().parse_args(argv)
    try:
        if args.command == "parse":
            if args.poll_interval < 1 or args.max_wait_seconds < 1:
                raise MinerUError("poll-interval 和 max-wait-seconds 必须大于 0。")
            return run_parse(args)
        if args.command == "status":
            return run_status(args)
        raise MinerUError("未知命令。")
    except MinerUError as exc:
        print("MinerU：", str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
