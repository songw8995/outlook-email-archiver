from __future__ import annotations

import os
import gc
import shutil
import tempfile
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable, Iterator

import pythoncom

from .models import Candidate
from .outlook_reader import OL_MAIL, OutlookSession, com_value, effective_time, valid_datetime
from .utils import iso, sha256_file, sha256_text


class MsgFileSession:
    """通过经典 Outlook 只读打开独立 MSG 文件；源文件始终原样复制。"""

    def __init__(self, paths: Iterable[str]):
        self.input_paths = [Path(value).expanduser().resolve() for value in paths]
        self.outlook: OutlookSession | None = None
        self._temp_context: tempfile.TemporaryDirectory[str] | None = None
        self._temp_dir: Path | None = None

    def __enter__(self) -> "MsgFileSession":
        self._temp_context = tempfile.TemporaryDirectory(prefix="outlook-msd-open-")
        self._temp_dir = Path(self._temp_context.name)
        self.outlook = OutlookSession().__enter__()
        return self

    def __exit__(self, *args: object) -> None:
        if self.outlook:
            self.outlook.__exit__(*args)
        self.outlook = None
        gc.collect()
        if self._temp_context:
            self._temp_context.cleanup()
        self._temp_context = None
        self._temp_dir = None

    def _files(self) -> list[tuple[Path, Path]]:
        found: dict[str, tuple[Path, Path]] = {}
        for selected in self.input_paths:
            if selected.is_dir():
                root = selected
                for path in selected.rglob("*"):
                    if path.is_file() and path.suffix.lower() in {".msg", ".msd"}:
                        found[os.path.normcase(str(path))] = (path, root)
            elif selected.is_file() and selected.suffix.lower() in {".msg", ".msd"}:
                found[os.path.normcase(str(selected))] = (selected, selected.parent)
        return sorted(found.values(), key=lambda pair: os.path.normcase(str(pair[0])))

    @contextmanager
    def _open_item(self, source: Path) -> Iterator[Any]:
        if not self.outlook:
            raise RuntimeError("MSG Outlook 会话尚未初始化。")
        if source.suffix.lower() == ".msg":
            item = self.outlook.namespace.OpenSharedItem(str(source))
            try:
                yield item
            finally:
                item = None
                gc.collect()
            return
        # “.msd”仅兼容实际内容为 Outlook MSG、但扩展名写错的文件。
        if self._temp_dir is None:
            raise RuntimeError("MSD 临时读取目录尚未初始化。")
        renamed = self._temp_dir / (sha256_text(os.path.normcase(str(source)))[:16] + ".msg")
        if not renamed.exists():
            shutil.copy2(source, renamed)
        item = self.outlook.namespace.OpenSharedItem(str(renamed))
        try:
            yield item
        finally:
            item = None
            gc.collect()
            pythoncom.CoFreeUnusedLibraries()

    def scan(
        self,
        *,
        include_date: Callable[[datetime | None], bool],
        limit: int | None = None,
        progress: Callable[[dict[str, Any]], None] | None = None,
        stop_requested: Callable[[], bool] | None = None,
        **_: Any,
    ) -> tuple[list[Candidate], dict[str, Any]]:
        files = self._files()
        candidates: list[Candidate] = []
        stats: dict[str, Any] = {
            "folders_selected": len({str(root) for _, root in files}), "items_total": len(files),
            "mail_total": 0, "non_mail": 0, "date_outside": 0, "unknown_time": 0,
            "earliest": None, "latest": None, "unsupported_types": {}, "scan_stopped": False,
            "warnings": [], "folders": [],
        }
        folder_stats: dict[str, dict[str, Any]] = {}
        for index, (path, root) in enumerate(files, start=1):
            if stop_requested and stop_requested():
                stats["scan_stopped"] = True
                stats["warnings"].append("用户在 MSG 文件扫描期间请求停止；后续文件尚未枚举。")
                break
            relative_parent = path.parent.relative_to(root)
            segments = [root.name, *relative_parent.parts] if relative_parent.parts else [root.name]
            folder_path = " / ".join(segments)
            folder_id = "msg-folder:" + sha256_text(os.path.normcase(str(path.parent)))
            row = folder_stats.setdefault(folder_path, {"folder": folder_path, "items": 0, "mail": 0, "selected": 0, "non_mail": 0, "earliest": None, "latest": None})
            row["items"] += 1
            at = None; basis = "unknown"; item_class = OL_MAIL; subject = ""
            try:
                with self._open_item(path) as item:
                    item_class = int(com_value(item, "Class", -1))
                    subject = str(com_value(item, "Subject", "") or "")
                    if item_class != OL_MAIL:
                        stats["non_mail"] += 1; row["non_mail"] += 1
                        key = str(item_class); stats["unsupported_types"][key] = int(stats["unsupported_types"].get(key, 0)) + 1
                        stats["warnings"].append(f"不是普通 MailItem，导出时将记录失败：{path}")
                    else:
                        stats["mail_total"] += 1; row["mail"] += 1
                    at, basis = effective_time(item, "custom")
            except Exception as exc:
                stats["mail_total"] += 1; row["mail"] += 1
                stats["warnings"].append(f"MSG/MSD 元数据暂时无法读取，导出时会重试：{path}：{type(exc).__name__}: {exc}")
            if at is None:
                stats["unknown_time"] += 1
            else:
                at_iso = iso(at)
                row["earliest"] = min(filter(None, [row["earliest"], at_iso])); row["latest"] = max(filter(None, [row["latest"], at_iso]))
                stats["earliest"] = min(filter(None, [stats["earliest"], at_iso])); stats["latest"] = max(filter(None, [stats["latest"], at_iso]))
            if not include_date(at):
                stats["date_outside"] += 1
                continue
            normalized = os.path.normcase(str(path))
            source_store_key = "msg-file:" + sha256_text(normalized)
            candidates.append(
                Candidate(
                    store_id="msg-import", store_name="MSG 文件导入", store_path=str(root),
                    folder_id=folder_id, folder_path=folder_path, folder_segments=segments,
                    folder_kind="custom", entry_id=str(path), effective_at=at, effective_basis=basis,
                    item_class=item_class, last_modified=valid_datetime(datetime.fromtimestamp(path.stat().st_mtime)),
                    size=path.stat().st_size,
                    segment_ids=["msg-segment:" + sha256_text(os.path.normcase(str(root.joinpath(*segments[1:index + 1])))) for index in range(len(segments))],
                    source_store_key=source_store_key, archive_store_id="msg-import",
                    subject=subject,
                )
            )
            row["selected"] += 1
            if progress and (index % 25 == 0 or index == len(files)):
                progress({"event": "scan_progress", "folder": folder_path, "current": index, "total": len(files)})
            if limit and len(candidates) >= limit:
                break
        stats["folders"] = list(folder_stats.values())
        candidates.sort(key=lambda value: (value.effective_at or datetime.min.replace(tzinfo=datetime.now().astimezone().tzinfo), os.path.normcase(value.entry_id)))
        return candidates, stats

    @staticmethod
    def inspect_candidate(candidate: Candidate) -> tuple[str, str, str]:
        path = Path(candidate.entry_id)
        if not path.is_file():
            raise FileNotFoundError(f"MSG 源文件已不可访问：{path}")
        source_key = sha256_text((candidate.source_store_key or "msg-file") + "\0" + os.path.normcase(str(path)))
        return source_key, sha256_file(path), ""

    def prepare_message(
        self, candidate: Candidate, staging: Path, *, retries: int = 2,
        save_msg: bool = True, save_attachments: bool = True, save_inline_images: bool = True,
    ):
        if not self.outlook:
            raise RuntimeError("MSG Outlook 会话尚未初始化。")
        source = Path(candidate.entry_id)
        source_key, signature, _ = self.inspect_candidate(candidate)
        with self._open_item(source) as item:
            return self.outlook.prepare_open_item(
                item, candidate, staging, retries=retries, original_path=source,
                source_key_override=source_key, signature_override=signature,
                save_msg=save_msg, save_attachments=save_attachments,
                save_inline_images=save_inline_images,
            )
