from __future__ import annotations

import csv
import gc
import json
import os
import re
import shutil
import sys
import time as time_module
import uuid
from contextlib import contextmanager
from dataclasses import asdict
from datetime import datetime, time
from pathlib import Path
from typing import Any, Callable, Iterator
from urllib.parse import unquote, urlparse

from .models import Candidate, MessageSnapshot
from .msg_reader import MsgFileSession
from .outlook_reader import OutlookSession
from .pst import work_root_for_output
from .renderer import render_message, write_metadata
from .state import ArchiveState
from .utils import ensure_within, iso, parse_years, read_json, safe_name, sha256_file, sha256_text, write_json_atomic

LAYOUT_NAMES = {"folder_first": "文件夹优先", "year_first": "年份优先"}


def load_config(path: str | Path) -> dict[str, Any]:
    config_path = Path(path).resolve()
    config = json.loads(config_path.read_text(encoding="utf-8-sig"))
    required = ("source_mode", "output_root", "layout")
    missing = [name for name in required if name not in config]
    if missing:
        raise ValueError("配置缺少字段：" + "、".join(missing))
    if config["source_mode"] not in {"outlook", "pst", "msg"}:
        raise ValueError("source_mode 必须是 outlook、pst 或 msg。")
    if config["source_mode"] == "pst" and not config.get("pst_path"):
        raise ValueError("PST 来源必须提供 pst_path。")
    if config["source_mode"] == "msg" and not config.get("msg_paths"):
        raise ValueError("MSG 来源必须选择至少一个 .msg/.msd 文件或文件夹。")
    if config["layout"] not in LAYOUT_NAMES:
        raise ValueError("layout 必须是 folder_first 或 year_first。")
    if config["source_mode"] != "msg" and (not isinstance(config.get("folder_ids"), list) or not config["folder_ids"]):
        raise ValueError("至少选择一个邮件文件夹。")
    config.setdefault("store_id", "msg-import")
    config.setdefault("folder_ids", [])
    if config["source_mode"] == "pst" and not config.get("pst_work_root"):
        config["pst_work_root"] = str(work_root_for_output(config["output_root"]))
    if config.get("years") and (config.get("start_date") or config.get("end_date")):
        raise ValueError("年份筛选和自定义日期范围不能同时使用。")
    config["years"] = parse_years(config.get("years"))
    if bool(config.get("start_date")) != bool(config.get("end_date")):
        raise ValueError("自定义日期范围必须同时提供开始和结束日期。")
    config.setdefault("save_msg", True)
    config.setdefault("index_mode", "incremental")
    if config["index_mode"] not in {"incremental", "full"}:
        raise ValueError("index_mode 必须为 incremental 或 full。")
    config.setdefault("save_attachments", True)
    config.setdefault("save_inline_images", True)
    config.setdefault("save_body_sources", any(bool(config[name]) for name in ("save_msg", "save_attachments", "save_inline_images")))
    config["_config_path"] = str(config_path)
    return config


def date_filter(config: dict[str, Any]) -> Callable[[datetime | None], bool]:
    years = set(config.get("years") or [])
    if years:
        return lambda value: value is not None and value.year in years
    if config.get("start_date"):
        start = datetime.combine(datetime.fromisoformat(config["start_date"]).date(), time.min).astimezone()
        end = datetime.combine(datetime.fromisoformat(config["end_date"]).date(), time.min).astimezone()
        if start >= end:
            raise ValueError("结束日期必须晚于开始日期；结束日期不包含当天零点之后。")
        return lambda value: value is not None and start <= value < end
    return lambda _value: True


def _emit_default(event: dict[str, Any]) -> None:
    print("EVENT " + json.dumps(event, ensure_ascii=False, default=str), flush=True)


def _make_session(config: dict[str, Any], session_factory: Callable[[str | None], Any]):
    if config["source_mode"] == "msg":
        return MsgFileSession(config.get("msg_paths", []))
    if session_factory is OutlookSession:
        return OutlookSession(
            config.get("pst_path") if config["source_mode"] == "pst" else None,
            config.get("pst_work_root") if config["source_mode"] == "pst" else None,
            [config["output_root"]] if config["source_mode"] == "pst" else None,
        )
    return session_factory(config.get("pst_path") if config["source_mode"] == "pst" else None)


@contextmanager
def archive_lock(root: Path) -> Iterator[None]:
    root.mkdir(parents=True, exist_ok=True)
    path = root / ".archive.lock"
    stream = path.open("a+b")
    if path.stat().st_size == 0:
        stream.write(b"0")
        stream.flush()
    try:
        if os.name == "nt":
            import msvcrt

            stream.seek(0)
            try:
                msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError as exc:
                raise RuntimeError("另一个导出进程正在使用这个归档根，请等待其结束。") from exc
        yield
    finally:
        if os.name == "nt":
            import msvcrt

            stream.seek(0)
            try:
                msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
            except OSError:
                pass
        stream.close()


def _content_profile(config: dict[str, Any]) -> dict[str, bool]:
    return {
        "save_msg": bool(config.get("save_msg", True)),
        "save_attachments": bool(config.get("save_attachments", True)),
        "save_inline_images": bool(config.get("save_inline_images", True)),
        "save_body_sources": bool(config.get("save_body_sources", True)),
    }


def initialize_root(root: Path, layout: str, content_profile: dict[str, bool] | None = None) -> None:
    marker = root / "archive.json"
    if marker.exists():
        existing = read_json(marker)
        if existing.get("layout") != layout:
            raise RuntimeError(
                f"这个归档根已固定为“{LAYOUT_NAMES.get(existing.get('layout'), existing.get('layout'))}”。"
                "若要改变目录顺序，请选择新的保存位置。"
            )
        if content_profile is not None:
            existing_profile = existing.get("content_profile")
            if existing_profile is None:
                existing_profile = {"save_msg": True, "save_attachments": True, "save_inline_images": True, "save_body_sources": True}
            if existing_profile != content_profile:
                raise RuntimeError("这个归档根已使用另一组‘导出内容’选项。为避免混合和重复，请选择新的保存位置。")
    else:
        write_json_atomic(
            marker,
            {
                "archive_version": 2,
                "layout": layout,
                "content_profile": content_profile or {"save_msg": True, "save_attachments": True, "save_inline_images": True, "save_body_sources": True},
                "created_at": iso(datetime.now().astimezone()),
            },
        )
    for name in ("state", "logs", "indexes"):
        (root / name).mkdir(parents=True, exist_ok=True)
    start = root / "00_开始这里.md"
    if not start.exists():
        start.write_text(
            "# Outlook 邮件归档\n\n"
            "- 邮件按所选 Outlook/PST 来源、原文件夹和年份保存。\n"
            "- ‘完整归档’按单封目录保存；‘仅 Markdown’直接生成日期开头的 `.md` 文件。\n"
            "- 是否保存 `original.msg`、普通附件和正文图片，以创建归档根时的界面选项为准。\n"
            "- `indexes/emails.csv` 是邮件索引；`logs/` 保存每次结果报告和失败清单。\n"
            "- 移动整个归档目录不会破坏邮件目录内的相对链接。\n",
            encoding="utf-8-sig",
        )


def _source_component(snapshot: MessageSnapshot) -> str:
    label = snapshot.store_name or (Path(snapshot.store_path).stem if snapshot.store_path else "Outlook")
    return f"{safe_name(label, 48)}_s{sha256_text(snapshot.store_id)[:8]}"


def _folder_components(snapshot: MessageSnapshot) -> list[str]:
    output = []
    cumulative: list[str] = []
    for index, segment in enumerate(snapshot.folder_segments):
        cumulative.append(segment)
        identity = snapshot.segment_ids[index] if index < len(snapshot.segment_ids) else snapshot.store_id + "\0" + "/".join(cumulative)
        output.append(f"{safe_name(segment, 44)}_f{sha256_text(identity)[:6]}")
    return output


def message_relative_path(snapshot: MessageSnapshot, layout: str, root: Path, *, markdown_only: bool = False) -> Path:
    source = _source_component(snapshot)
    folders = _folder_components(snapshot)
    year = str(snapshot.effective_at.year) if snapshot.effective_at else "年份未知"
    stamp = snapshot.effective_at.strftime("%Y-%m-%d") if snapshot.effective_at else "日期未知"
    subject = safe_name(snapshot.subject, 72, "无主题")
    suffix = ".md" if markdown_only else ""
    message = f"{stamp}_{subject}{suffix}"

    def compose(folder_parts: list[str], message_part: str) -> list[str]:
        return [source, *folder_parts, year, message_part] if layout == "folder_first" else [source, year, *folder_parts, message_part]

    parts = compose(folders, message)
    relative = Path(*parts)
    path_probe = root / relative if markdown_only else root / relative / "attachments" / ("x" * 72)
    if len(str(path_probe)) > 238:
        subject = safe_name(snapshot.subject, 28, "无主题")
        message = f"{stamp}_{subject}{suffix}"
        parts = compose(folders, message)
        relative = Path(*parts)
        path_probe = root / relative if markdown_only else root / relative / "attachments" / ("x" * 72)
    if len(str(path_probe)) > 238:
        short_folder = f"_deep_path_f{sha256_text(snapshot.folder_path)[:12]}"
        parts = compose([short_folder], message)
        relative = Path(*parts)
        path_probe = root / relative if markdown_only else root / relative / "attachments" / ("x" * 72)
    if len(str(path_probe)) > 238:
        raise OSError("保存位置本身过长，无法安全生成附件路径；请选择更短的归档根，例如 D:\\MailArchive。")
    return relative


def _available_collision_path(path: Path) -> Path:
    """Keep the simple name and add (2), (3)... only when the name is occupied."""
    if not path.exists():
        return path
    stem, suffix = (path.stem, path.suffix) if path.suffix else (path.name, "")
    for number in range(2, 10000):
        candidate = path.with_name(f"{stem} ({number}){suffix}")
        if not candidate.exists():
            return candidate
    raise OSError(f"同名项目过多，无法分配安全文件名：{path.name}")


def build_manifest(
    directory: Path,
    status: str,
    warnings: list[str],
    errors: list[str],
    *,
    required_files: list[str] | None = None,
) -> dict[str, Any]:
    files = []
    for path in sorted(directory.rglob("*")):
        if path.is_file() and path.name != ".archive-manifest.json":
            files.append({"path": path.relative_to(directory).as_posix(), "size": path.stat().st_size, "sha256": sha256_file(path)})
    return {
        "version": 2,
        "status": status,
        "created_at": iso(datetime.now().astimezone()),
        "warnings": warnings,
        "errors": errors,
        "required_files": required_files or ["email.md", "metadata.json"],
        "files": files,
    }


def _publish_directory(staging: Path, final: Path, attempts: int = 4) -> None:
    last_error: OSError | None = None
    for attempt in range(attempts):
        try:
            os.replace(staging, final)
            return
        except (PermissionError, OSError) as exc:
            last_error = exc
            if attempt + 1 >= attempts:
                break
            gc.collect()
            time_module.sleep(0.3 * (attempt + 1))
    assert last_error is not None
    raise last_error


def _publish_file(staging_file: Path, final: Path, attempts: int = 4) -> None:
    final.parent.mkdir(parents=True, exist_ok=True)
    last_error: OSError | None = None
    for attempt in range(attempts):
        try:
            os.replace(staging_file, final)
            return
        except (PermissionError, OSError) as exc:
            last_error = exc
            if attempt + 1 >= attempts:
                break
            gc.collect()
            time_module.sleep(0.3 * (attempt + 1))
    assert last_error is not None
    raise last_error


def verify_message_dir(directory: Path, full_hash: bool = False) -> tuple[bool, list[str]]:
    issues: list[str] = []
    manifest_path = directory / ".archive-manifest.json"
    if not manifest_path.exists():
        return False, ["缺少 .archive-manifest.json"]
    try:
        manifest = read_json(manifest_path)
    except Exception as exc:
        return False, [f"清单无法读取：{exc}"]
    if manifest.get("status") not in {"complete", "warning"}:
        issues.append(f"状态不是可跳过终态：{manifest.get('status')}")
    required = set(manifest.get("required_files") or ["original.msg", "email.md", "metadata.json", "body.txt"])
    listed = {item.get("path") for item in manifest.get("files", [])}
    for name in sorted(required - listed):
        issues.append(f"清单缺少必要文件：{name}")
    for record in manifest.get("files", []):
        path = directory / str(record.get("path", ""))
        try:
            ensure_within(directory, path)
        except ValueError as exc:
            issues.append(str(exc))
            continue
        if not path.is_file():
            issues.append(f"文件缺失：{record.get('path')}")
        elif path.stat().st_size != int(record.get("size", -1)):
            issues.append(f"大小不符：{record.get('path')}")
        elif full_hash and sha256_file(path) != record.get("sha256"):
            issues.append(f"哈希不符：{record.get('path')}")
    markdown_path = directory / "email.md"
    if markdown_path.exists():
        markdown = markdown_path.read_text(encoding="utf-8-sig", errors="replace")
        for match in re.finditer(r"!?\[[^\]]*\]\(([^)]+)\)", markdown):
            target = unquote(match.group(1).strip().split("#", 1)[0])
            if not target or urlparse(target).scheme in {"http", "https", "mailto"}:
                continue
            linked = directory / target.replace("/", os.sep)
            try:
                ensure_within(directory, linked)
            except ValueError:
                issues.append(f"Markdown 链接越过邮件目录：{target}")
                continue
            if not linked.exists():
                issues.append(f"Markdown 本地链接失效：{target}")
    return not issues, issues


def verify_archive_record(root: Path, archive_path: str, expected_hash: str = "", *, full_hash: bool = False) -> tuple[bool, list[str]]:
    target = ensure_within(root, root / archive_path)
    if target.suffix.lower() == ".md":
        if not target.is_file():
            return False, ["Markdown 文件缺失"]
        if target.stat().st_size == 0:
            return False, ["Markdown 文件为空"]
        if expected_hash and sha256_file(target) != expected_hash:
            return False, ["Markdown 文件哈希不符"]
        return True, []
    return verify_message_dir(target, full_hash=full_hash)


def _failure_record(candidate: Candidate, source_key: str, stage: str, exc: Exception) -> dict[str, str]:
    return {
        "source_key": source_key[:12],
        "subject": candidate.subject.strip() or "(无主题)",
        "source": candidate.entry_id if Path(candidate.entry_id).suffix.lower() in {".msg", ".msd"} else candidate.store_name,
        "effective_at": iso(candidate.effective_at) or "未知",
        "folder": candidate.folder_path,
        "stage": stage,
        "error": f"{type(exc).__name__}: {exc}",
        "retry": "点击‘继续’可重试",
    }


def preview(config: dict[str, Any], *, limit: int | None = None, emit: Callable[[dict[str, Any]], None] = _emit_default, session_factory: Callable[[str | None], Any] = OutlookSession) -> dict[str, Any]:
    with _make_session(config, session_factory) as session:
        candidates, stats = session.scan(
            store_id=config["store_id"],
            folder_ids=config["folder_ids"],
            folder_paths=config.get("folder_paths", []),
            include_subfolders=bool(config.get("include_subfolders", True)),
            include_date=date_filter(config),
            limit=limit,
            progress=emit,
        )
    result = {"selected": len(candidates), **stats}
    emit({"event": "preview_complete", **result})
    return result


def _record_jsonl(path: Path, event: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as stream:
        stream.write(json.dumps(event, ensure_ascii=False, default=str) + "\n")


def export_archive(
    config: dict[str, Any],
    *,
    limit: int | None = None,
    stop_file: Path | None = None,
    emit: Callable[[dict[str, Any]], None] = _emit_default,
    session_factory: Callable[[str | None], Any] = OutlookSession,
) -> dict[str, Any]:
    root = Path(config["output_root"]).expanduser().resolve()
    save_msg = bool(config.get("save_msg", True))
    save_attachments = bool(config.get("save_attachments", True))
    save_inline_images = bool(config.get("save_inline_images", True))
    markdown_only = not (save_msg or save_attachments or save_inline_images)
    save_body_sources = bool(config.get("save_body_sources", not markdown_only)) and not markdown_only
    profile = {
        "save_msg": save_msg,
        "save_attachments": save_attachments,
        "save_inline_images": save_inline_images,
        "save_body_sources": save_body_sources,
    }
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S_") + uuid.uuid4().hex[:8]
    with archive_lock(root):
        initialize_root(root, config["layout"], profile)
        log_path = root / "logs" / f"run_{run_id}.jsonl"
        failures: list[dict[str, str]] = []
        counts = {"success": 0, "warning": 0, "skipped": 0, "failed": 0, "partial": 0, "source_unavailable": 0, "unprocessed": 0}
        started = datetime.now().astimezone()
        with ArchiveState(root / "state" / "archive.sqlite") as state:
            state.set_setting("layout", config["layout"])
            state.connection.execute(
                "INSERT INTO runs(run_id,started_at,mode,status) VALUES(?,?,?,?)",
                (run_id, iso(started), config["source_mode"], "scanning"),
            )
            state.connection.commit()
            with _make_session(config, session_factory) as session:
                candidates, scan_stats = session.scan(
                    store_id=config["store_id"],
                    folder_ids=config["folder_ids"],
                    folder_paths=config.get("folder_paths", []),
                    include_subfolders=bool(config.get("include_subfolders", True)),
                    include_date=date_filter(config),
                    limit=limit,
                    progress=emit,
                    stop_requested=(lambda: bool(stop_file and stop_file.exists())),
                )
                total = len(candidates)
                state.connection.execute("UPDATE runs SET selected_count=?,status='running' WHERE run_id=?", (total, run_id))
                state.connection.commit()
                snapshot_config = {key: value for key, value in config.items() if not key.startswith("_")}
                write_json_atomic(root / "state" / f"run_{run_id}_config.json", snapshot_config)
                candidate_path = root / "state" / f"run_{run_id}_candidates.jsonl"
                for candidate in candidates:
                    _record_jsonl(candidate_path, {**asdict(candidate), "effective_at": iso(candidate.effective_at), "last_modified": iso(candidate.last_modified)})
                emit({"event": "export_start", "run_id": run_id, "total": total, "root": str(root)})

                for number, candidate in enumerate(candidates, start=1):
                    if stop_file and stop_file.exists():
                        counts["unprocessed"] = total - number + 1
                        emit({"event": "stopped", "unprocessed": counts["unprocessed"]})
                        break
                    minimum_free = int(config.get("minimum_free_bytes", 256 * 1024 * 1024))
                    if shutil.disk_usage(root).free < minimum_free:
                        counts["unprocessed"] = total - number + 1
                        scan_stats["warnings"].append(
                            f"归档磁盘可用空间低于安全阈值 {minimum_free / 1024**2:.0f} MB，已暂停；请释放空间后继续。"
                        )
                        emit({"event": "stopped", "unprocessed": counts["unprocessed"], "reason": "disk_space"})
                        break
                    staging = root / "state" / "staging" / run_id / sha256_text(candidate.store_id + "\0" + candidate.entry_id)[:16]
                    try:
                        source_key, signature, internet_id = session.inspect_candidate(candidate)
                    except Exception as exc:
                        counts["source_unavailable"] += 1
                        failure = _failure_record(candidate, sha256_text(candidate.entry_id), "读取源邮件", exc)
                        failures.append(failure)
                        _record_jsonl(log_path, {"event": "source_unavailable", **failure})
                        emit({"event": "progress", "current": number, "total": total, **counts, "folder": candidate.folder_path})
                        continue
                    latest = state.latest(source_key)
                    if latest and latest["content_signature"] == signature and latest["folder_id"] == candidate.folder_id and latest["status"] in {"complete", "warning"}:
                        valid, issues = verify_archive_record(
                            root, latest["archive_path"], latest["manifest_sha256"] or "", full_hash=False
                        )
                        if valid:
                            counts["skipped"] += 1
                            _record_jsonl(log_path, {"event": "item_skipped_verified", "source_key": source_key[:12], "archive_path": latest["archive_path"]})
                            emit({"event": "progress", "current": number, "total": total, **counts, "folder": candidate.folder_path})
                            continue
                        _record_jsonl(log_path, {"event": "item_repair_needed", "source_key": source_key[:12], "issues": issues})

                    if latest and latest["content_signature"] == signature and latest["folder_id"] == candidate.folder_id and latest["status"] in {"partial", "failed"}:
                        version = int(latest["version"])
                    else:
                        version = state.next_version(source_key)
                    try:
                        if staging.exists():
                            shutil.rmtree(staging)
                        last_prepare_error = None
                        for attempt in range(2):
                            try:
                                snapshot = session.prepare_message(
                                    candidate,
                                    staging,
                                    save_msg=save_msg,
                                    save_attachments=save_attachments,
                                    save_inline_images=save_inline_images,
                                )
                                break
                            except Exception as exc:
                                last_prepare_error = exc
                                if attempt == 0:
                                    _record_jsonl(log_path, {"event": "item_retry", "source_key": source_key[:12], "error": f"{type(exc).__name__}: {exc}"})
                                    if staging.exists():
                                        shutil.rmtree(staging)
                                    time_module.sleep(0.5)
                                else:
                                    raise
                        if snapshot.source_key != source_key or snapshot.content_signature != signature:
                            snapshot.warnings.append("邮件在扫描与导出之间发生变化；本次按实际导出版本保存。")
                        _, warnings = render_message(
                            snapshot,
                            staging,
                            bool(config.get("include_addresses", True)),
                            save_msg=save_msg,
                            save_body_sources=save_body_sources,
                            save_inline_images=save_inline_images,
                        )
                        status = "partial" if snapshot.errors else ("warning" if warnings else "complete")
                        for error_text in snapshot.errors:
                            detail = _failure_record(candidate, source_key, "附件或邮件内容导出", RuntimeError(error_text))
                            if detail["subject"] == "(无主题)" and snapshot.subject.strip():
                                detail["subject"] = snapshot.subject.strip()
                            failures.append(detail)
                        desired = ensure_within(
                            root,
                            root / message_relative_path(snapshot, config["layout"], root, markdown_only=markdown_only),
                        )
                        repairing_latest = bool(
                            latest
                            and latest["content_signature"] == signature
                            and latest["folder_id"] == candidate.folder_id
                            and (root / latest["archive_path"]).exists()
                        )
                        final = ensure_within(root, root / latest["archive_path"]) if repairing_latest else _available_collision_path(desired)
                        if final.exists():
                            recovery = root / "state" / "failed-artifacts" / run_id / final.name
                            recovery.parent.mkdir(parents=True, exist_ok=True)
                            recovery = _available_collision_path(recovery)
                            shutil.move(str(final), str(recovery))
                        relative = final.relative_to(root)

                        # Durable intent before publishing: replay after interruption, even if
                        # the process dies between publishing the message and updating indexes.
                        index_metadata = (root / "state" / "items" / f"{snapshot.source_key}_v{version}.json") if markdown_only else final / "metadata.json"
                        _record_jsonl(root / "state" / "index-pending.jsonl", {"metadata": index_metadata.relative_to(root).as_posix()})

                        if markdown_only:
                            if status == "partial":
                                # Pure Markdown has no selected binary outputs, so only a true message/body failure can be partial.
                                raise RuntimeError("邮件正文未能完整读取：" + "; ".join(snapshot.errors))
                            _publish_file(staging / "email.md", final)
                            manifest_hash = sha256_file(final)
                            sidecar = {
                                "format": "markdown_only",
                                "source_key": snapshot.source_key,
                                "store_id": snapshot.store_id,
                                "folder_id": snapshot.folder_id,
                                "store": snapshot.store_name,
                                "folder": snapshot.folder_path,
                                "subject": snapshot.subject or "(无主题)",
                                "sender": snapshot.sender_name,
                                "effective_at": iso(snapshot.effective_at) or "",
                                "status": status,
                                "version": version,
                                "archive_path": relative.as_posix(),
                                "sha256": manifest_hash,
                            }
                            sidecar_path = root / "state" / "items" / f"{snapshot.source_key}_v{version}.json"
                            write_json_atomic(sidecar_path, sidecar)
                            if staging.exists():
                                shutil.rmtree(staging)
                        else:
                            write_metadata(snapshot, staging, status=status, warnings=warnings, version=version)
                            required_files = ["email.md", "metadata.json"]
                            if save_msg:
                                required_files.append("original.msg")
                            if save_body_sources:
                                required_files.append("body.txt")
                            manifest = build_manifest(
                                staging,
                                status,
                                warnings,
                                snapshot.errors,
                                required_files=required_files,
                            )
                            write_json_atomic(staging / ".archive-manifest.json", manifest)
                            final.parent.mkdir(parents=True, exist_ok=True)
                            _publish_directory(staging, final)
                            manifest_hash = sha256_file(final / ".archive-manifest.json")
                        state.upsert_message(
                            source_key=snapshot.source_key,
                            version=version,
                            store_id=snapshot.store_id,
                            folder_id=snapshot.folder_id,
                            entry_id=snapshot.entry_id,
                            internet_message_id=internet_id,
                            content_signature=snapshot.content_signature,
                            status=status,
                            archive_path=relative.as_posix(),
                            manifest_sha256=manifest_hash,
                            updated_at=iso(datetime.now().astimezone()) or "",
                            last_error="; ".join(snapshot.errors),
                        )
                        counts[status if status in counts else "success"] += 1
                        _record_jsonl(log_path, {"event": "item_published", "source_key": source_key[:12], "status": status, "archive_path": relative.as_posix(), "warnings": len(warnings)})
                    except Exception as exc:
                        counts["failed"] += 1
                        failure = _failure_record(candidate, source_key, "导出或写入", exc)
                        failures.append(failure)
                        _record_jsonl(log_path, {"event": "item_failed", **failure})
                        if staging.exists():
                            recovery = root / "state" / "failed-artifacts" / run_id / staging.name
                            recovery.parent.mkdir(parents=True, exist_ok=True)
                            if recovery.exists():
                                recovery = recovery.with_name(recovery.name + "_" + uuid.uuid4().hex[:6])
                            shutil.move(str(staging), str(recovery))
                    emit({"event": "progress", "current": number, "total": total, **counts, "folder": candidate.folder_path})

            candidate_total = len(candidates)
            scan_stats.setdefault("warnings", []).extend(getattr(session, "cleanup_warnings", []))
            emit({"event": "finalizing"})
            terminal = counts["success"] + counts["warning"] + counts["skipped"] + counts["failed"] + counts["partial"] + counts["source_unavailable"] + counts["unprocessed"]
            if terminal != candidate_total:
                counts["unprocessed"] += candidate_total - terminal
            finished = datetime.now().astimezone()
            report_path = root / "logs" / f"run_{run_id}_report.md"
            failure_path = root / "logs" / f"run_{run_id}_failures.csv"
            _write_failure_csv(failure_path, failures)
            _write_run_report(report_path, run_id, config, scan_stats, counts, candidate_total, started, finished, log_path, failure_path)
            emit({"event": "report_ready", "report": str(report_path), "failures": str(failure_path)})
            index_started = time_module.perf_counter()
            index_result = rebuild_indexes(root, emit=emit, mode=config.get("index_mode", "incremental"))
            index_seconds = round(time_module.perf_counter() - index_started, 3)
            _record_jsonl(log_path, {"event": "indexes_complete", "seconds": index_seconds, **index_result})
            with report_path.open("a", encoding="utf-8") as report_stream:
                report_stream.write(f"\n## 收尾情况\n\n- 索引更新：{index_result['mode']}；共 {index_result['messages']} 封；本次读取记录 {index_result['scanned']} 条；耗时 {index_seconds:.3f} 秒。\n")
            status = "stopped" if counts["unprocessed"] or scan_stats.get("scan_stopped") else ("completed_with_errors" if counts["failed"] or counts["partial"] or counts["source_unavailable"] else "completed")
            state.connection.execute(
                "UPDATE runs SET finished_at=?,completed_count=?,status=?,report_path=? WHERE run_id=?",
                (iso(finished), candidate_total - counts["unprocessed"], status, report_path.relative_to(root).as_posix(), run_id),
            )
            state.connection.commit()
        result = {"run_id": run_id, "root": str(root), "report": str(report_path), "failures": str(failure_path), "selected": candidate_total, "scan_stopped": bool(scan_stats.get("scan_stopped")), **counts}
        _record_jsonl(log_path, {"event": "run_complete", **result})
        emit({"event": "run_complete", **result})
        return result


def _write_failure_csv(path: Path, failures: list[dict[str, str]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(
            stream,
            fieldnames=["source_key", "subject", "source", "effective_at", "folder", "stage", "error", "retry"],
        )
        writer.writeheader()
        writer.writerows(failures)


def _write_run_report(path: Path, run_id: str, config: dict[str, Any], scan: dict[str, Any], counts: dict[str, int], selected: int, started: datetime, finished: datetime, log_path: Path, failure_path: Path) -> None:
    total_terminal = sum(counts.values())
    lines = [
        f"# Outlook 邮件归档结果：{run_id}", "",
        f"- 开始：{iso(started)}", f"- 完成：{iso(finished)}",
        f"- 来源：{('PST 工作副本' if config['source_mode'] == 'pst' else ('独立 MSG/MSD 文件' if config['source_mode'] == 'msg' else '当前经典 Outlook 可见内容'))}",
        f"- 目录布局：{LAYOUT_NAMES[config['layout']]}",
        f"- 导出内容：{('仅 Markdown（每封邮件直接保存为一个 .md 文件）' if not any(bool(config.get(name, True)) for name in ('save_msg', 'save_attachments', 'save_inline_images')) else '完整/自定义归档')}",
        f"- 已选普通邮件候选：{selected}", f"- 互斥终态合计：{total_terminal}", "",
        "## 结果", "", "| 状态 | 数量 |", "| --- | ---: |",
        f"| 新完成 | {counts['success']} |", f"| 完成但有警告 | {counts['warning']} |",
        f"| 已有且核验后跳过 | {counts['skipped']} |", f"| 部分完成 | {counts['partial']} |",
        f"| 失败 | {counts['failed']} |", f"| 源已不可访问 | {counts['source_unavailable']} |",
        f"| 停止后未处理 | {counts['unprocessed']} |", "",
        "## 扫描范围", "",
        f"- 所选可访问文件夹：{scan.get('folders_selected', 0)}",
        f"- 枚举项目总数：{scan.get('items_total', 0)}",
        f"- 普通 MailItem：{scan.get('mail_total', 0)}",
        f"- 非 MailItem（未计入候选）：{scan.get('non_mail', 0)}",
        f"- 非 MailItem 类型统计：{json.dumps(scan.get('unsupported_types', {}), ensure_ascii=False)}",
        f"- 日期范围外：{scan.get('date_outside', 0)}",
        f"- 时间未知：{scan.get('unknown_time', 0)}", "",
        "## 说明", "",
        "- Outlook 来源仅代表本次运行时经典 Outlook 能枚举和读取的当前可见内容，不代表服务器全部邮件。",
        "- 程序没有发送、删除、移动、标记已读或修改邮件，也没有关闭 Outlook。",
        "- 归档不会因源邮件以后被删除而自动删除。",
        "- 外链图片默认未下载；附件未执行，压缩包未解压；OCR 和附件正文提取未启用。",
        f"- 技术日志：`{log_path.name}`",
        f"- 失败清单：`{failure_path.name}`", "",
    ]
    if failure_path.exists():
        with failure_path.open(encoding="utf-8-sig", newline="") as stream:
            failure_rows = list(csv.DictReader(stream))
        if failure_rows:
            lines.extend(["## 失败与未导出条目", "", "| 主题 | 来源文件夹 | 阶段 | 原因 |", "| --- | --- | --- | --- |"]) 
            for row in failure_rows[:50]:
                values = [str(row.get(key, "")).replace("|", "\\|").replace("\n", " ") for key in ("subject", "folder", "stage", "error")]
                lines.append(f"| {values[0]} | {values[1]} | {values[2]} | {values[3]} |")
            if len(failure_rows) > 50:
                lines.append(f"\n其余 {len(failure_rows) - 50} 条请打开失败 CSV 清单。")
            lines.append("")
    warnings = list(scan.get("warnings", []))
    folder_rows = list(scan.get("folders", []))
    if folder_rows:
        lines.extend(["## 文件夹核对", "", "| 文件夹 | 总项目 | MailItem | 选中 | 非 MailItem | 最早 | 最晚 |", "| --- | ---: | ---: | ---: | ---: | --- | --- |"]) 
        for row in folder_rows:
            folder_name = str(row.get("folder", "")).replace("|", "\\|")
            lines.append(f"| {folder_name} | {row.get('items', 0)} | {row.get('mail', 0)} | {row.get('selected', 0)} | {row.get('non_mail', 0)} | {row.get('earliest') or ''} | {row.get('latest') or ''} |")
        lines.append("")
    if warnings:
        lines.extend(["## 扫描警告", ""] + [f"- {value}" for value in warnings] + [""])
    path.write_text("\n".join(lines), encoding="utf-8-sig")


def rebuild_indexes(root: Path, *, emit: Callable[[dict[str, Any]], None] = lambda event: None, mode: str = "full") -> dict[str, Any]:
    if mode not in {"incremental", "full"}:
        raise ValueError("未知索引更新方式")
    records: list[dict[str, Any]] = []
    folders: dict[tuple[str, str], dict[str, Any]] = {}
    pending = root / "state" / "index-pending.jsonl"
    changed: set[Path] = set()
    incremental = mode == "incremental" and (root / "indexes" / "emails.csv").exists() and (root / "indexes" / "folders.csv").exists()
    if incremental:
        try:
            with (root / "indexes" / "emails.csv").open(encoding="utf-8-sig", newline="") as stream:
                reader = csv.DictReader(stream)
                if not {"archive_path", "email_md", "year", "effective_at", "subject", "folder"}.issubset(reader.fieldnames or []):
                    raise ValueError("索引列不完整")
                records = list(reader)
                if any(None in row or any(value is None for value in row.values()) for row in records):
                    raise ValueError("索引行不完整")
            with (root / "indexes" / "folders.csv").open(encoding="utf-8-sig", newline="") as stream:
                for row in csv.DictReader(stream):
                    folders[(row["store_id_hash"], row["folder_id_hash"])] = row
            if pending.exists():
                for line in pending.read_text(encoding="utf-8").splitlines():
                    changed.add(ensure_within(root, root / json.loads(line)["metadata"]))
            # Upgrade/recovery: an older release may have published messages but
            # stopped before indexing, without our new pending journal.
            database = root / "state" / "archive.sqlite"
            if database.exists():
                import sqlite3
                indexed_paths = {row["archive_path"] for row in records}
                connection = sqlite3.connect(database.as_uri() + "?mode=ro", uri=True)
                try:
                    for key, version, archive_path in connection.execute("SELECT source_key,version,archive_path FROM messages"):
                        if archive_path not in indexed_paths:
                            metadata = (root / "state" / "items" / f"{key}_v{version}.json") if archive_path.lower().endswith('.md') else root / archive_path / "metadata.json"
                            changed.add(ensure_within(root, metadata))
                finally:
                    connection.close()
        except (OSError, ValueError, KeyError, TypeError):
            incremental = False
            records = []; folders = {}
    emit({"event": "index_mode", "mode": "增量更新" if incremental else "完整重建（首次或手动选择）"})
    profile = read_json(root / "archive.json").get("content_profile", {}) if (root / "archive.json").exists() else {}
    pure_md = bool(profile) and not any(profile.get(key, True) for key in ("save_msg", "save_attachments", "save_inline_images"))
    def metadata_files():
        if incremental:
            yield from sorted(path for path in changed if path.name == "metadata.json")
            return
        if pure_md:
            return
        for directory, children, files in os.walk(root):
            if Path(directory) == root:
                children[:] = [name for name in children if name not in {"state", "logs", "indexes"}]
            if "metadata.json" in files:
                yield Path(directory) / "metadata.json"
    scanned = 0
    replaced_paths: set[str] = set()
    old_records = records
    records = []
    def progress():
        nonlocal scanned
        scanned += 1
        if scanned % 500 == 0:
            emit({"event": "index_progress", "current": scanned})
    for metadata_path in metadata_files():
        progress()
        if incremental and not metadata_path.exists():
            continue  # interrupted before publication; no new message to index
        if "state" in metadata_path.parts or "logs" in metadata_path.parts:
            continue
        try:
            data = read_json(metadata_path)
            source = data["source"]
            message = data["message"]
            relative = metadata_path.parent.relative_to(root).as_posix()
            replaced_paths.add(relative)
            records.append({
                "effective_at": message.get("effective_at") or "",
                "year": (message.get("effective_at") or "unknown")[:4],
                "subject": message.get("subject") or "(无主题)",
                "sender": message.get("sender_name") or "",
                "store": source.get("store_name") or "",
                "folder": source.get("folder_path") or "",
                "status": data.get("status") or "",
                "version": data.get("version") or 1,
                "archive_path": relative,
                "email_md": f"{relative}/email.md",
            })
            folders[(sha256_text(source.get("store_id", ""))[:12], sha256_text(source.get("folder_id", ""))[:12])] = {
                "store": source.get("store_name") or "", "folder": source.get("folder_path") or "",
                "store_id_hash": sha256_text(source.get("store_id", ""))[:12],
                "folder_id_hash": sha256_text(source.get("folder_id", ""))[:12],
            }
        except Exception:
            if incremental:
                raise
            continue
    item_dir = root / "state" / "items"
    if item_dir.exists():
        for sidecar_path in (sorted(path for path in changed if path.parent == item_dir) if incremental else item_dir.glob("*.json")):
            progress()
            if incremental and not sidecar_path.exists():
                continue
            try:
                data = read_json(sidecar_path)
                if data.get("format") != "markdown_only":
                    continue
                archive_path = str(data["archive_path"])
                if not (root / archive_path).is_file():
                    continue
                replaced_paths.add(archive_path)
                records.append({
                    "effective_at": data.get("effective_at") or "",
                    "year": (data.get("effective_at") or "unknown")[:4],
                    "subject": data.get("subject") or "(无主题)",
                    "sender": data.get("sender") or "",
                    "store": data.get("store") or "",
                    "folder": data.get("folder") or "",
                    "status": data.get("status") or "",
                    "version": data.get("version") or 1,
                    "archive_path": archive_path,
                    "email_md": archive_path,
                })
                folders[(sha256_text(data.get("store_id", ""))[:12], sha256_text(data.get("folder_id", ""))[:12])] = {
                    "store": data.get("store") or "",
                    "folder": data.get("folder") or "",
                    "store_id_hash": sha256_text(data.get("store_id", ""))[:12],
                    "folder_id_hash": sha256_text(data.get("folder_id", ""))[:12],
                }
            except Exception:
                if incremental:
                    raise
                continue
    records.extend(row for row in old_records if row["archive_path"] not in replaced_paths)
    records.sort(key=lambda value: (value["effective_at"], value["archive_path"]))
    index_dir = root / "indexes"
    index_dir.mkdir(parents=True, exist_ok=True)
    with (index_dir / "emails.csv.tmp").open("w", encoding="utf-8-sig", newline="") as stream:
        fields = ["effective_at", "year", "subject", "sender", "store", "folder", "status", "version", "archive_path", "email_md"]
        writer = csv.DictWriter(stream, fieldnames=fields)
        writer.writeheader(); writer.writerows(records)
    with (index_dir / "folders.csv.tmp").open("w", encoding="utf-8-sig", newline="") as stream:
        fields = ["store", "folder", "store_id_hash", "folder_id_hash"]
        writer = csv.DictWriter(stream, fieldnames=fields)
        writer.writeheader(); writer.writerows(sorted(folders.values(), key=lambda value: (value["store"], value["folder"])))
    years_dir = index_dir / "years"
    years_dir.mkdir(parents=True, exist_ok=True)
    grouped: dict[str, list[dict[str, Any]]] = {}
    for record in records:
        grouped.setdefault(record["year"], []).append(record)
    for year, values in grouped.items():
        lines = [f"# {year} 年邮件索引", ""]
        for value in values:
            from .utils import relative_link
            rel = relative_link(os.path.relpath(root / value["email_md"], years_dir).replace("\\", "/"))
            lines.append(f"- {value['effective_at'] or '时间未知'} · [{value['subject']}]({rel}) · {value['folder']}")
        target = years_dir / f"{safe_name(year, 12)}.md"
        temporary = target.with_suffix(".md.tmp")
        temporary.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
        os.replace(temporary, target)
    os.replace(index_dir / "folders.csv.tmp", index_dir / "folders.csv")
    os.replace(index_dir / "emails.csv.tmp", index_dir / "emails.csv")
    if pending.exists():
        pending.unlink()
    return {"messages": len(records), "folders": len(folders), "scanned": scanned, "mode": "增量更新" if incremental else "完整重建"}


def verify_archive(root: str | Path, *, full_hash: bool = True, open_msg: bool = False) -> dict[str, Any]:
    archive_root = Path(root).resolve()
    checked = 0; passed = 0; issues: list[dict[str, Any]] = []
    for manifest in archive_root.rglob(".archive-manifest.json"):
        if "state" in manifest.parts:
            continue
        checked += 1
        valid, found = verify_message_dir(manifest.parent, full_hash=full_hash)
        if valid and open_msg and (manifest.parent / "original.msg").exists():
            try:
                with OutlookSession() as session:
                    item = session.namespace.OpenSharedItem(str(manifest.parent / "original.msg"))
                    if int(getattr(item, "Class", -1)) != 43:
                        found.append("original.msg 可打开但不是 MailItem。")
            except Exception as exc:
                found.append(f"original.msg 无法由 Outlook 打开：{type(exc).__name__}: {exc}")
            valid = not found
        if valid:
            passed += 1
        else:
            issues.append({"directory": str(manifest.parent.relative_to(archive_root)), "issues": found})
    item_dir = archive_root / "state" / "items"
    if item_dir.exists():
        for sidecar_path in item_dir.glob("*.json"):
            try:
                data = read_json(sidecar_path)
                if data.get("format") != "markdown_only":
                    continue
                checked += 1
                valid, found = verify_archive_record(
                    archive_root,
                    str(data.get("archive_path", "")),
                    str(data.get("sha256", "")) if full_hash else "",
                    full_hash=full_hash,
                )
                if valid:
                    passed += 1
                else:
                    issues.append({"directory": str(data.get("archive_path", "")), "issues": found})
            except Exception as exc:
                checked += 1
                issues.append({"directory": str(sidecar_path.relative_to(archive_root)), "issues": [f"状态记录无法读取：{exc}"]})
    result = {"root": str(archive_root), "checked": checked, "passed": passed, "failed": checked - passed, "issues": issues}
    write_json_atomic(archive_root / "logs" / f"verify_{datetime.now():%Y%m%d_%H%M%S}.json", result)
    return result


def render_archive(root: str | Path, *, include_addresses: bool = True) -> dict[str, Any]:
    archive_root = Path(root).resolve()
    rebuilt = 0; failed: list[dict[str, str]] = []
    for metadata_path in archive_root.rglob("metadata.json"):
        if "state" in metadata_path.parts:
            continue
        try:
            data = read_json(metadata_path)
            source = data["source"]; message = data["message"]
            from .models import AttachmentRecord
            snapshot = MessageSnapshot(
                source_key=source["source_key"], content_signature=source["content_signature"],
                store_id=source["store_id"], store_name=source["store_name"], store_path=source.get("store_path", ""),
                folder_id=source["folder_id"], folder_path=source["folder_path"], folder_segments=source.get("folder_segments", []),
                folder_kind="custom", entry_id=source["entry_id"], internet_message_id=source.get("internet_message_id", ""),
                conversation_id=source.get("conversation_id", ""), subject=message.get("subject", ""),
                sender_name=message.get("sender_name", ""), sender_email=message.get("sender_email", ""),
                to=message.get("to", []), cc=message.get("cc", []), bcc=message.get("bcc", []),
                sent_at=_parse_iso(message.get("sent_at")), received_at=_parse_iso(message.get("received_at")),
                created_at=_parse_iso(message.get("created_at")), modified_at=_parse_iso(message.get("modified_at")),
                effective_at=_parse_iso(message.get("effective_at")), effective_basis=message.get("effective_basis", "unknown"),
                direction=message.get("direction", "unknown"),
                html_body=(metadata_path.parent / "body.source.html").read_text(encoding="utf-8-sig") if (metadata_path.parent / "body.source.html").exists() else "",
                text_body=(metadata_path.parent / "body.txt").read_text(encoding="utf-8-sig") if (metadata_path.parent / "body.txt").exists() else "",
                segment_ids=source.get("segment_ids", []),
                attachments=[AttachmentRecord(**item) for item in data.get("attachments", [])], warnings=data.get("warnings", []), errors=data.get("errors", []),
            )
            output_name = f"email.regenerated_{datetime.now():%Y%m%d_%H%M%S}.md"
            render_message(snapshot, metadata_path.parent, include_addresses, output_name=output_name)
            rebuilt += 1
        except Exception as exc:
            failed.append({"metadata": str(metadata_path), "error": f"{type(exc).__name__}: {exc}"})
    rebuild_indexes(archive_root)
    return {"rebuilt": rebuilt, "failed": len(failed), "failures": failed}


def _parse_iso(value: str | None) -> datetime | None:
    return datetime.fromisoformat(value) if value else None
