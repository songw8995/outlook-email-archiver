from __future__ import annotations

import mimetypes
import os
import shutil
import time
from contextlib import suppress
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable

import pythoncom
import win32com.client

from .models import AttachmentRecord, Candidate, FolderInfo, MessageSnapshot
from .pst import PstCopy, prepare_pst_copy
from .utils import ensure_within, iso, sha256_file, sha256_text, unique_actual_name

OL_MAIL = 43
OL_MAIL_ITEM = 0
OL_MSG_UNICODE = 9
OL_STORE_UNICODE = 3
DEFAULT_FOLDERS = {
    6: "inbox",
    5: "sent",
    4: "outbox",
    3: "deleted",
    16: "drafts",
    23: "junk",
}
PROP_INTERNET_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
PROP_ATTACH_CID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PROP_ATTACH_LOCATION = "http://schemas.microsoft.com/mapi/proptag/0x3713001F"
PROP_ATTACH_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
PROP_ATTACH_MIME = "http://schemas.microsoft.com/mapi/proptag/0x370E001F"


def release(obj: Any) -> None:
    if obj is not None:
        with suppress(Exception):
            del obj


def com_value(obj: Any, name: str, default: Any = "") -> Any:
    try:
        return getattr(obj, name)
    except Exception:
        return default


def mapi_value(obj: Any, schema: str, default: Any = "") -> Any:
    try:
        return obj.PropertyAccessor.GetProperty(schema)
    except Exception:
        return default


def valid_datetime(value: Any) -> datetime | None:
    try:
        result = value if isinstance(value, datetime) else datetime.fromisoformat(str(value))
        if result.year < 1900 or result.year > 9998:
            return None
        return result.astimezone() if result.tzinfo else result.replace(tzinfo=datetime.now().astimezone().tzinfo)
    except Exception:
        return None


def effective_time(item: Any, folder_kind: str) -> tuple[datetime | None, str]:
    sent = valid_datetime(com_value(item, "SentOn", None))
    received = valid_datetime(com_value(item, "ReceivedTime", None))
    created = valid_datetime(com_value(item, "CreationTime", None))
    if folder_kind == "sent":
        choices = ((sent, "SentOn"), (received, "ReceivedTime fallback"), (created, "CreationTime fallback"))
    elif folder_kind in {"outbox", "drafts"} and not bool(com_value(item, "Sent", False)):
        choices = ((created, "CreationTime"), (sent, "SentOn fallback"), (received, "ReceivedTime fallback"))
    else:
        choices = ((received, "ReceivedTime"), (sent, "SentOn fallback"), (created, "CreationTime fallback"))
    return next(((value, basis) for value, basis in choices if value is not None), (None, "unknown"))


class OutlookSession:
    """一个进程内的串行 Outlook COM 会话；从不调用 Quit。"""

    def __init__(
        self,
        pst_path: str | None = None,
        pst_work_root: str | Path | None = None,
        pst_reuse_roots: list[str | Path] | None = None,
    ):
        self.pst_path = pst_path
        self.pst_work_root = Path(pst_work_root).expanduser().resolve() if pst_work_root else None
        self.pst_reuse_roots = list(pst_reuse_roots or [])
        self.app = None
        self.namespace = None
        self.pst_copy: PstCopy | None = None
        self._added_pst_root = None
        self.cleanup_warnings: list[str] = []

    def __enter__(self) -> "OutlookSession":
        pythoncom.CoInitialize()
        try:
            # Outlook startup can reopen the source PST registered in its profile.
            # Validate/reuse or exclusively copy BEFORE activating Outlook COM.
            if self.pst_path:
                self.pst_copy = prepare_pst_copy(
                    self.pst_path,
                    work_root=self.pst_work_root,
                    reuse_roots=self.pst_reuse_roots,
                )
            self.app = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.app.GetNamespace("MAPI")
            if self.pst_path:
                self._mount_pst_work_copy(self.pst_path)
            return self
        except Exception:
            self.__exit__()
            raise

    def __exit__(self, *_: object) -> None:
        if self._added_pst_root is not None:
            error = None
            for attempt in range(3):
                try:
                    self.namespace.RemoveStore(self._added_pst_root)
                    error = None
                    break
                except Exception as exc:
                    error = exc
                    if attempt < 2:
                        time.sleep(0.3)
            if error is not None:
                warning = f"临时 PST 工作副本卸载失败：{self.pst_copy.copy if self.pst_copy else ''}；{error}。请在 Outlook 数据文件设置中手动关闭此工作副本。原有数据文件未卸载。"
                self.cleanup_warnings.append(warning)
                print(warning, flush=True)
        self._added_pst_root = None
        self.namespace = None
        self.app = None
        pythoncom.CoUninitialize()

    def _store_file_paths(self) -> dict[str, Any]:
        found: dict[str, Any] = {}
        stores = self.namespace.Stores
        for index in range(1, stores.Count + 1):
            store = stores.Item(index)
            path = str(com_value(store, "FilePath", ""))
            if path:
                found[os.path.normcase(os.path.abspath(path))] = store
        return found

    def _mount_pst_work_copy(self, source: str) -> None:
        before = self._store_file_paths()
        if self.pst_copy is None:
            raise RuntimeError("PST 工作副本尚未准备完成，不能挂载。")
        copy_norm = os.path.normcase(os.path.abspath(self.pst_copy.copy))
        if copy_norm not in before:
            self.namespace.AddStoreEx(str(self.pst_copy.copy), OL_STORE_UNICODE)
            after = self._store_file_paths()
            store = after.get(copy_norm)
            if store is None:
                raise RuntimeError("PST 工作副本已请求挂载，但 Outlook 未返回对应 Store。")
            self._added_pst_root = store.GetRootFolder()

    def stores(self) -> list[Any]:
        stores = self.namespace.Stores
        result = []
        wanted = os.path.normcase(os.path.abspath(self.pst_copy.copy)) if self.pst_copy else ""
        for index in range(1, stores.Count + 1):
            store = stores.Item(index)
            if wanted:
                path = str(com_value(store, "FilePath", ""))
                if os.path.normcase(os.path.abspath(path)) != wanted:
                    continue
            result.append(store)
        return result

    def _default_ids(self, store: Any) -> dict[str, str]:
        output: dict[str, str] = {}
        for value, kind in DEFAULT_FOLDERS.items():
            try:
                output[str(store.GetDefaultFolder(value).EntryID)] = kind
            except Exception:
                continue
        return output

    def catalog(self) -> tuple[list[dict[str, Any]], list[FolderInfo], list[str]]:
        stores_output: list[dict[str, Any]] = []
        folders_output: list[FolderInfo] = []
        warnings: list[str] = []
        for store in self.stores():
            store_id = str(com_value(store, "StoreID", ""))
            store_name = str(com_value(store, "DisplayName", "未命名邮箱/PST"))
            store_path = str(com_value(store, "FilePath", ""))
            stores_output.append({"store_id": store_id, "store_name": store_name, "store_path": store_path})
            defaults = self._default_ids(store)
            root = store.GetRootFolder()
            visited: set[str] = set()

            def walk(parent: Any, segments: list[str], segment_ids: list[str], inherited_kind: str = "custom") -> None:
                try:
                    children = parent.Folders
                    count = int(children.Count)
                except Exception as exc:
                    warnings.append(f"文件夹无法访问：{' / '.join(segments) or store_name}：{exc}")
                    return
                for index in range(1, count + 1):
                    child = None
                    try:
                        child = children.Item(index)
                        folder_id = str(child.EntryID)
                        if folder_id in visited:
                            warnings.append(f"检测到重复文件夹引用，已停止重复遍历：{child.Name}")
                            continue
                        visited.add(folder_id)
                        name = str(child.Name)
                        child_segments = segments + [name]
                        child_segment_ids = segment_ids + [folder_id]
                        kind = defaults.get(folder_id, inherited_kind)
                        selectable = int(com_value(child, "DefaultItemType", -1)) == OL_MAIL_ITEM
                        is_search = False
                        with suppress(Exception):
                            is_search = bool(store.IsSearchFolder(child))
                        if is_search:
                            selectable = False
                        folders_output.append(
                            FolderInfo(
                                store_id=store_id,
                                store_name=store_name,
                                store_path=store_path,
                                folder_id=folder_id,
                                folder_name=name,
                                folder_path=" / ".join(child_segments),
                                segments=child_segments,
                                parent_id=str(com_value(parent, "EntryID", "")) or None,
                                kind=kind,
                                selectable=selectable,
                                warning="搜索文件夹默认排除" if is_search else ("非邮件文件夹" if not selectable else ""),
                                segment_ids=child_segment_ids,
                            )
                        )
                        walk(child, child_segments, child_segment_ids, kind)
                    except Exception as exc:
                        warnings.append(f"枚举子文件夹失败：{' / '.join(segments)} / #{index}：{exc}")

            walk(root, [], [])
        return stores_output, folders_output, warnings

    def selected_folders(self, store_id: str, folder_ids: Iterable[str], include_subfolders: bool, folder_paths: Iterable[str] = ()) -> tuple[list[FolderInfo], list[str]]:
        _, catalog, warnings = self.catalog()
        same_store = list(catalog) if self.pst_copy is not None else [folder for folder in catalog if folder.store_id == store_id]
        if self.pst_copy is None and any(folder.store_path.lower().endswith(".pst") for folder in same_store):
            raise RuntimeError("所选 Store 是 PST。请在窗口选择“PST 文件”，让程序使用经过校验的完整工作副本。")
        selected_ids = set(folder_ids)
        wanted_paths = set(folder_paths)
        if wanted_paths:
            path_matches = {folder.folder_id for folder in same_store if folder.folder_path in wanted_paths}
            selected_ids = path_matches if self.pst_copy is not None else (selected_ids | path_matches)
        known = {folder.folder_id for folder in same_store}
        missing = selected_ids - known
        warnings.extend(f"所选文件夹当前不可访问：{value[:12]}…" for value in missing)
        if include_subfolders:
            changed = True
            while changed:
                changed = False
                for folder in same_store:
                    if folder.parent_id in selected_ids and folder.folder_id not in selected_ids:
                        selected_ids.add(folder.folder_id)
                        changed = True
        output = [folder for folder in same_store if folder.folder_id in selected_ids and folder.selectable]
        output.sort(key=lambda value: value.folder_path.casefold())
        return output, warnings

    def scan(
        self,
        *,
        store_id: str,
        folder_ids: Iterable[str],
        include_subfolders: bool,
        folder_paths: Iterable[str] = (),
        include_date: Callable[[datetime | None], bool],
        limit: int | None = None,
        progress: Callable[[dict[str, Any]], None] | None = None,
        stop_requested: Callable[[], bool] | None = None,
    ) -> tuple[list[Candidate], dict[str, Any]]:
        folders, warnings = self.selected_folders(store_id, folder_ids, include_subfolders, folder_paths)
        candidates: list[Candidate] = []
        stats = {"folders_selected": len(folders), "items_total": 0, "mail_total": 0, "non_mail": 0, "date_outside": 0, "unknown_time": 0, "earliest": None, "latest": None, "unsupported_types": {}, "scan_stopped": False, "warnings": warnings, "folders": []}
        for folder_info in folders:
            if stop_requested and stop_requested():
                stats["scan_stopped"] = True
                stats["warnings"].append("用户在扫描期间请求停止；后续文件夹尚未枚举。")
                break
            folder_stat = {"folder": folder_info.folder_path, "items": 0, "mail": 0, "selected": 0, "non_mail": 0, "earliest": None, "latest": None}
            try:
                folder = self.namespace.GetFolderFromID(folder_info.folder_id, folder_info.store_id)
                items = folder.Items
                count = int(items.Count)
            except Exception as exc:
                stats["warnings"].append(f"无法读取文件夹 {folder_info.folder_path}：{exc}")
                continue
            for index in range(1, count + 1):
                if stop_requested and stop_requested():
                    stats["scan_stopped"] = True
                    stats["warnings"].append(f"用户在扫描期间请求停止：{folder_info.folder_path} 后续项目尚未枚举。")
                    break
                item = None
                stats["items_total"] += 1
                folder_stat["items"] += 1
                try:
                    item = items.Item(index)
                    item_class = int(com_value(item, "Class", -1))
                    if item_class != OL_MAIL:
                        stats["non_mail"] += 1
                        folder_stat["non_mail"] += 1
                        key = str(item_class)
                        stats["unsupported_types"][key] = int(stats["unsupported_types"].get(key, 0)) + 1
                        continue
                    stats["mail_total"] += 1
                    folder_stat["mail"] += 1
                    at, basis = effective_time(item, folder_info.kind)
                    if at is None:
                        stats["unknown_time"] += 1
                    else:
                        at_iso = iso(at)
                        folder_stat["earliest"] = min(filter(None, [folder_stat["earliest"], at_iso]))
                        folder_stat["latest"] = max(filter(None, [folder_stat["latest"], at_iso]))
                        stats["earliest"] = min(filter(None, [stats["earliest"], at_iso]))
                        stats["latest"] = max(filter(None, [stats["latest"], at_iso]))
                    if not include_date(at):
                        stats["date_outside"] += 1
                        continue
                    candidates.append(
                        Candidate(
                            store_id=folder_info.store_id,
                            store_name=folder_info.store_name,
                            store_path=folder_info.store_path,
                            folder_id=folder_info.folder_id,
                            folder_path=folder_info.folder_path,
                            folder_segments=folder_info.segments,
                            folder_kind=folder_info.kind,
                            entry_id=str(com_value(item, "EntryID", "")),
                            effective_at=at,
                            effective_basis=basis,
                            item_class=item_class,
                            last_modified=valid_datetime(com_value(item, "LastModificationTime", None)),
                            size=int(com_value(item, "Size", 0) or 0),
                            segment_ids=folder_info.segment_ids,
                            source_store_key=(
                                "pst:" + sha256_text(os.path.normcase(str(self.pst_copy.source)) + "\0" + self.pst_copy.sha256)
                                if self.pst_copy else folder_info.store_id
                            ),
                            archive_store_id=(
                                "pst:" + sha256_text(os.path.normcase(str(self.pst_copy.source)) + "\0" + self.pst_copy.sha256)
                                if self.pst_copy else folder_info.store_id
                            ),
                            subject=str(com_value(item, "Subject", "") or ""),
                        )
                    )
                    folder_stat["selected"] += 1
                    if limit and len(candidates) >= limit:
                        break
                except Exception as exc:
                    stats["warnings"].append(f"读取候选项目失败：{folder_info.folder_path} #{index}：{type(exc).__name__}: {exc}")
                finally:
                    item = None
                if progress and index % 100 == 0:
                    progress({"event": "scan_progress", "folder": folder_info.folder_path, "current": index, "total": count})
            stats["folders"].append(folder_stat)
            if stats["scan_stopped"] or (limit and len(candidates) >= limit):
                break
        candidates.sort(key=lambda item: ((item.effective_at or datetime.min.replace(tzinfo=datetime.now().astimezone().tzinfo)), item.folder_path.casefold(), item.entry_id))
        return candidates, stats

    def prepare_message(
        self, candidate: Candidate, staging: Path, *, retries: int = 2,
        save_msg: bool = True, save_attachments: bool = True, save_inline_images: bool = True,
    ) -> MessageSnapshot:
        item = self.namespace.GetItemFromID(candidate.entry_id, candidate.store_id)
        return self.prepare_open_item(
            item, candidate, staging, retries=retries, save_msg=save_msg,
            save_attachments=save_attachments, save_inline_images=save_inline_images,
        )

    def prepare_open_item(
        self,
        item: Any,
        candidate: Candidate,
        staging: Path,
        *,
        retries: int = 2,
        original_path: Path | None = None,
        source_key_override: str = "",
        signature_override: str = "",
        save_msg: bool = True,
        save_attachments: bool = True,
        save_inline_images: bool = True,
    ) -> MessageSnapshot:
        if int(com_value(item, "Class", -1)) != OL_MAIL:
            raise RuntimeError("源项目不再是普通 MailItem。")
        entry_id = candidate.entry_id if original_path else str(com_value(item, "EntryID", candidate.entry_id))
        source_identity = candidate.source_store_key or candidate.store_id
        source_key = source_key_override or sha256_text(source_identity + "\0" + entry_id)
        sent = valid_datetime(com_value(item, "SentOn", None))
        received = valid_datetime(com_value(item, "ReceivedTime", None))
        created = valid_datetime(com_value(item, "CreationTime", None))
        modified = valid_datetime(com_value(item, "LastModificationTime", None))
        effective_at, basis = effective_time(item, candidate.folder_kind)
        internet_id = str(mapi_value(item, PROP_INTERNET_ID, "") or "")
        conversation_id = str(com_value(item, "ConversationID", "") or "")
        subject = str(com_value(item, "Subject", "") or "")
        signature_material = "\0".join(
            [
                source_key,
                iso(modified) or "",
                str(com_value(item, "Size", 0) or 0),
                subject,
                str(com_value(com_value(item, "Attachments", None), "Count", 0) or 0),
            ]
        )
        content_signature = signature_override or sha256_text(signature_material)
        staging.mkdir(parents=True, exist_ok=True)
        msg_path = ensure_within(staging, staging / "original.msg")
        if save_msg:
            if original_path:
                if original_path.resolve() != msg_path.resolve():
                    shutil.copy2(original_path, msg_path)
            else:
                item.SaveAs(str(msg_path), OL_MSG_UNICODE)
            if not msg_path.exists() or msg_path.stat().st_size == 0:
                raise OSError("Outlook 未生成有效的 original.msg。")

        to, cc, bcc = self._recipients(item)
        attachments: list[AttachmentRecord] = []
        collection = com_value(item, "Attachments", None)
        count = int(com_value(collection, "Count", 0) or 0)
        for index in range(1, count + 1):
            attachment = collection.Item(index)
            original = str(com_value(attachment, "FileName", "") or f"attachment_{index:03d}")
            cid = str(mapi_value(attachment, PROP_ATTACH_CID, "") or "")
            location = str(mapi_value(attachment, PROP_ATTACH_LOCATION, "") or "")
            hidden = bool(mapi_value(attachment, PROP_ATTACH_HIDDEN, False))
            mime = str(mapi_value(attachment, PROP_ATTACH_MIME, "") or mimetypes.guess_type(original)[0] or "")
            inline = bool(cid or location or hidden)
            attachment_type = int(com_value(attachment, "Type", 1) or 1)
            should_save = save_inline_images if inline else save_attachments
            subdir = staging / ("images" if inline else "attachments")
            actual = unique_actual_name(subdir, original, index)
            target = ensure_within(subdir, subdir / actual)
            error = ""
            if should_save:
                subdir.mkdir(parents=True, exist_ok=True)
                for attempt in range(retries + 1):
                    try:
                        attachment.SaveAsFile(str(target))
                        if not target.exists():
                            raise OSError("Outlook 未写出附件文件。")
                        break
                    except Exception as exc:
                        error = f"{type(exc).__name__}: {exc}"
                        if attempt < retries:
                            time.sleep(0.2 * (attempt + 1))
                status = "saved" if target.exists() else "failed"
            else:
                status = "not_selected"
            attachments.append(
                AttachmentRecord(
                    index=index,
                    original_name=original,
                    actual_name=actual,
                    relative_path=f"{'images' if inline else 'attachments'}/{actual}",
                    content_id=cid,
                    content_location=location,
                    mime_type=mime,
                    hidden=hidden,
                    inline=inline,
                    size=target.stat().st_size if target.exists() else 0,
                    sha256=sha256_file(target) if target.exists() else "",
                    status=status,
                    error=error,
                    attachment_type=attachment_type,
                )
            )

        folder_kind = candidate.folder_kind
        sent_flag = bool(com_value(item, "Sent", False))
        if folder_kind == "sent":
            direction = "outgoing"
        elif folder_kind in {"outbox", "drafts"} and not sent_flag:
            direction = "unsent"
        elif folder_kind in {"inbox", "junk", "deleted"}:
            direction = "incoming"
        else:
            direction = "unknown"
        snapshot = MessageSnapshot(
            source_key=source_key,
            content_signature=content_signature,
            store_id=candidate.archive_store_id or source_identity,
            store_name=candidate.store_name,
            store_path=str(self.pst_copy.source) if self.pst_copy else candidate.store_path,
            folder_id=candidate.folder_id,
            folder_path=candidate.folder_path,
            folder_segments=candidate.folder_segments,
            folder_kind=folder_kind,
            entry_id=entry_id,
            internet_message_id=internet_id,
            conversation_id=conversation_id,
            subject=subject,
            sender_name=str(com_value(item, "SenderName", "") or ""),
            sender_email=str(com_value(item, "SenderEmailAddress", "") or ""),
            to=to,
            cc=cc,
            bcc=bcc,
            sent_at=sent,
            received_at=received,
            created_at=created,
            modified_at=modified,
            effective_at=effective_at,
            effective_basis=basis,
            direction=direction,
            html_body=str(com_value(item, "HTMLBody", "") or "").replace("\x00", ""),
            text_body=str(com_value(item, "Body", "") or "").replace("\x00", ""),
            segment_ids=candidate.segment_ids,
            attachments=attachments,
        )
        snapshot.errors.extend(record.error for record in attachments if record.status == "failed" and record.error)
        if not snapshot.html_body and not snapshot.text_body:
            suffix = "；请核对 original.msg。" if save_msg else "。"
            snapshot.warnings.append("邮件正文为空或当前权限下不可读取" + suffix)
        for record in attachments:
            if record.status == "saved" and record.attachment_type != 1:
                snapshot.warnings.append(
                    f"附件“{record.original_name}”类型为引用/嵌入/OLE（Outlook Type={record.attachment_type}）；"
                    "已按实际可读取结果保存，不能据此认定云端原文件已备份。"
                )
        return snapshot

    def inspect_candidate(self, candidate: Candidate) -> tuple[str, str, str]:
        """读取跳过判断所需的最小元数据，不读取正文或附件内容。"""
        item = self.namespace.GetItemFromID(candidate.entry_id, candidate.store_id)
        if int(com_value(item, "Class", -1)) != OL_MAIL:
            raise RuntimeError("源项目不再是普通 MailItem。")
        entry_id = str(com_value(item, "EntryID", candidate.entry_id))
        source_key = sha256_text((candidate.source_store_key or candidate.store_id) + "\0" + entry_id)
        modified = valid_datetime(com_value(item, "LastModificationTime", None))
        subject = str(com_value(item, "Subject", "") or "")
        attachments = com_value(item, "Attachments", None)
        signature_material = "\0".join(
            [
                source_key,
                iso(modified) or "",
                str(com_value(item, "Size", 0) or 0),
                subject,
                str(com_value(attachments, "Count", 0) or 0),
            ]
        )
        internet_id = str(mapi_value(item, PROP_INTERNET_ID, "") or "")
        return source_key, sha256_text(signature_material), internet_id

    @staticmethod
    def _recipients(item: Any) -> tuple[list[str], list[str], list[str]]:
        values = {1: [], 2: [], 3: []}
        recipients = com_value(item, "Recipients", None)
        count = int(com_value(recipients, "Count", 0) or 0)
        for index in range(1, count + 1):
            recipient = recipients.Item(index)
            name = str(com_value(recipient, "Name", "") or "")
            address = str(com_value(recipient, "Address", "") or "")
            display = f"{name} <{address}>" if name and address else (name or address)
            values.setdefault(int(com_value(recipient, "Type", 0) or 0), []).append(display)
        return values[1], values[2], values[3]


def doctor() -> dict[str, Any]:
    result: dict[str, Any] = {
        "python": os.sys.version.splitlines()[0],
        "python_executable": os.sys.executable,
        "python_bitness": 64 if os.sys.maxsize > 2**32 else 32,
        "pywin32": False,
        "tkinter": False,
        "outlook_com": False,
        "profile_access": False,
        "stores": 0,
        "errors": [],
    }
    try:
        import tkinter

        result["tkinter"] = bool(tkinter.TkVersion)
    except Exception as exc:
        result["errors"].append(f"Tkinter：{exc}")
    try:
        result["pywin32"] = True
        with OutlookSession() as session:
            result["outlook_com"] = True
            result["outlook_name"] = str(com_value(session.app, "Name", "Microsoft Outlook"))
            result["outlook_version"] = str(com_value(session.app, "Version", "未知"))
            stores, _, warnings = session.catalog()
            result["profile_access"] = True
            result["stores"] = len(stores)
            result["warnings"] = warnings
    except Exception as exc:
        result["errors"].append(f"经典 Outlook COM：{type(exc).__name__}: {exc}")
    return result
