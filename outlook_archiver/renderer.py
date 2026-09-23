from __future__ import annotations

import base64
import binascii
import html
import re
from pathlib import Path
from urllib.parse import unquote, urlparse

from bs4 import BeautifulSoup, NavigableString
from markdownify import markdownify

from .models import AttachmentRecord, MessageSnapshot
from .utils import iso, relative_link, safe_name, sha256_file, write_json_atomic

MAX_DATA_IMAGE_BYTES = 10 * 1024 * 1024
SAFE_DATA_IMAGE_TYPES = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/gif": ".gif",
    "image/webp": ".webp",
    "image/bmp": ".bmp",
}


def _normalize_resource(value: str) -> str:
    value = unquote((value or "").strip()).strip("<>")
    return value.lower()


def _escape_md(value: str) -> str:
    return (value or "").replace("\\", "\\\\").replace("|", "\\|").replace("\r", " ").replace("\n", " ")


def _display_person(name: str, address: str, include_addresses: bool) -> str:
    if include_addresses and address:
        return f"{name} <{address}>" if name else address
    return name or (address if include_addresses else "")


def _replace_data_image(tag, images_dir: Path, number: int, warnings: list[str]) -> str | None:
    src = str(tag.get("src", ""))
    match = re.match(r"^data:([^;,]+);base64,(.*)$", src, flags=re.I | re.S)
    if not match:
        warnings.append("正文中的 data 图片格式不受支持，未保存。")
        return None
    mime = match.group(1).lower()
    suffix = SAFE_DATA_IMAGE_TYPES.get(mime)
    if not suffix:
        warnings.append(f"正文中的 data 图片类型不受支持：{mime}")
        return None
    try:
        raw = base64.b64decode(match.group(2), validate=True)
    except (ValueError, binascii.Error):
        warnings.append("正文中的 data 图片编码损坏，未保存。")
        return None
    if len(raw) > MAX_DATA_IMAGE_BYTES:
        warnings.append("正文中的 data 图片超过 10 MB，未保存。")
        return None
    name = f"embedded_{number:03d}{suffix}"
    (images_dir / name).write_bytes(raw)
    return f"images/{name}"


def html_to_markdown(
    html_body: str,
    message_dir: Path,
    attachments: list[AttachmentRecord],
    warnings: list[str],
    save_inline_images: bool = True,
) -> str:
    soup = BeautifulSoup(html_body, "html.parser")
    for element in soup(["script", "style", "iframe", "object", "embed", "form"]):
        element.decompose()

    cid_map: dict[str, AttachmentRecord] = {}
    location_map: dict[str, AttachmentRecord] = {}
    for attachment in attachments:
        if attachment.content_id:
            cid_map[_normalize_resource(attachment.content_id)] = attachment
        if attachment.content_location:
            location_map[_normalize_resource(attachment.content_location)] = attachment

    images_dir = message_dir / "images"
    data_count = 0
    for tag in soup.find_all("img"):
        src = str(tag.get("src", "")).strip()
        alt = str(tag.get("alt") or "正文图片")
        normalized = _normalize_resource(src[4:] if src.lower().startswith("cid:") else src)
        record = None
        if src.lower().startswith("cid:"):
            record = cid_map.get(normalized)
        elif normalized in location_map:
            record = location_map[normalized]
        if record:
            if record.status == "saved":
                tag["src"] = relative_link(record.relative_path)
                tag["alt"] = alt
            elif record.status == "not_selected":
                tag.replace_with(NavigableString(f"[正文图片未保存（按当前设置）：{alt}]"))
            else:
                tag.replace_with(NavigableString(f"[正文图片保存失败：{alt}]"))
            continue
        if src.lower().startswith("data:"):
            if not save_inline_images:
                tag.replace_with(NavigableString(f"[正文图片未保存（按当前设置）：{alt}]"))
                continue
            images_dir.mkdir(parents=True, exist_ok=True)
            data_count += 1
            rel = _replace_data_image(tag, images_dir, data_count, warnings)
            if rel:
                tag["src"] = relative_link(rel)
                tag["alt"] = alt
            else:
                tag.replace_with(NavigableString("[正文图片无法恢复]"))
            continue
        scheme = urlparse(src).scheme.lower()
        if scheme in {"http", "https"}:
            safe_url = html.escape(src, quote=True)
            replacement = soup.new_tag("span")
            replacement.append(NavigableString("[外链图片未下载："))
            link = soup.new_tag("a", href=safe_url)
            link.string = alt
            replacement.append(link)
            replacement.append(NavigableString("]"))
            tag.replace_with(replacement)
            warnings.append(f"外链图片未下载：{src[:200]}")
            continue
        tag.replace_with(NavigableString(f"[正文图片无法恢复：{alt}]"))
        warnings.append(f"正文图片无法匹配本地资源：{src[:200] or '(空地址)'}")

    for tag in soup.find_all("a"):
        href = str(tag.get("href", "")).strip()
        scheme = urlparse(href).scheme.lower()
        if scheme and scheme not in {"http", "https", "mailto"}:
            tag.attrs.pop("href", None)
            warnings.append(f"已移除不安全或不支持的正文链接协议：{scheme}")

    if soup.find(attrs={"rowspan": True}) or soup.find(attrs={"colspan": True}):
        warnings.append("HTML 表格含合并单元格，Markdown 中可能出现结构降级；请核对 original.msg。")
    try:
        result = markdownify(
            str(soup),
            heading_style="ATX",
            bullets="-",
            table_infer_header=True,
            keep_inline_images_in=["td", "th"],
        )
    except TypeError:
        result = markdownify(str(soup), heading_style="ATX", bullets="-", table_infer_header=True)
    return result.strip()


def render_message(
    snapshot: MessageSnapshot,
    message_dir: Path,
    include_addresses: bool = True,
    output_name: str = "email.md",
    *,
    save_msg: bool = True,
    save_body_sources: bool = True,
    save_inline_images: bool = True,
) -> tuple[str, list[str]]:
    warnings = list(snapshot.warnings)
    if save_body_sources:
        (message_dir / "body.txt").write_text(snapshot.text_body or "", encoding="utf-8-sig")
    if snapshot.html_body:
        if save_body_sources:
            (message_dir / "body.source.html").write_text(snapshot.html_body, encoding="utf-8-sig")
        try:
            body_md = html_to_markdown(
                snapshot.html_body, message_dir, snapshot.attachments, warnings,
                save_inline_images=save_inline_images,
            )
        except Exception as exc:  # 保留纯文本回退，不因派生格式丢失原始层
            references = []
            if save_msg:
                references.append("original.msg")
            if save_body_sources:
                references.append("body.source.html")
            suffix = "，请查阅 " + " 和 ".join(references) if references else ""
            body_md = snapshot.text_body or f"[正文转换失败{suffix}]"
            warnings.append(f"HTML 转 Markdown 失败，已回退纯文本：{type(exc).__name__}: {exc}")
    else:
        body_md = snapshot.text_body or ("[没有可读取的正文，请查阅 original.msg]" if save_msg else "[没有可读取的正文]")

    title = snapshot.subject.strip() or "(无主题)"
    sender = _display_person(snapshot.sender_name, snapshot.sender_email, include_addresses)
    lines = [
        f"# {title.replace(chr(10), ' ').replace(chr(13), ' ')}",
        "",
        "## 邮件信息",
        "",
        "| 字段 | 内容 |",
        "| --- | --- |",
        f"| 主题 | {_escape_md(title)} |",
        f"| 发件人 | {_escape_md(sender)} |",
        f"| 收件人 | {_escape_md('；'.join(snapshot.to))} |",
        f"| 抄送 | {_escape_md('；'.join(snapshot.cc))} |",
        f"| 密送 | {_escape_md('；'.join(snapshot.bcc)) if snapshot.bcc else '未提供或不可读取'} |",
        f"| 发送时间 | {iso(snapshot.sent_at) or '未知'} |",
        f"| 接收时间 | {iso(snapshot.received_at) or '未知'} |",
        f"| 归档时间口径 | {iso(snapshot.effective_at) or '未知'}（{snapshot.effective_basis}） |",
        f"| 来源文件夹 | {_escape_md(snapshot.store_name + ' / ' + snapshot.folder_path)} |",
        f"| 方向 | {snapshot.direction} |",
        f"| 转换状态 | {'有警告' if warnings else '完成'}；OCR 未启用 |",
        "",
    ]
    if save_msg:
        lines.append("- [Outlook 原始导出](original.msg)")
    if save_body_sources:
        lines.append("- [纯文本正文源](body.txt)")
    if snapshot.html_body and save_body_sources:
        lines.append("- [HTML 正文源](body.source.html)")
    lines.extend(["", "## 正文", "", body_md, "", "## 附件", ""])
    ordinary = [item for item in snapshot.attachments if not item.inline]
    if not ordinary:
        lines.append("无普通附件（正文内嵌资源如有，已在正文位置显示）。")
    for item in ordinary:
        if item.status == "saved":
            label = item.original_name.replace("\\", "\\\\").replace("[", "\\[").replace("]", "\\]")
            lines.append(f"- [{label}]({relative_link(item.relative_path)})（保存为 `{item.actual_name}`）")
        elif item.status == "not_selected":
            lines.append(f"- {item.original_name}：未保存（按当前设置）")
        else:
            lines.append(f"- {item.original_name}：保存失败—{item.error}")
    if warnings:
        lines.extend(["", "## 转换警告", ""])
        lines.extend(f"- {warning}" for warning in warnings)
    lines.append("")
    markdown = "\n".join(lines)
    (message_dir / output_name).write_text(markdown, encoding="utf-8-sig")
    return markdown, warnings


def write_metadata(snapshot: MessageSnapshot, message_dir: Path, *, status: str, warnings: list[str], version: int) -> Path:
    files: list[dict[str, object]] = []
    for path in sorted(message_dir.rglob("*")):
        if path.is_file() and path.name not in {"metadata.json", ".complete.json"}:
            files.append(
                {
                    "path": path.relative_to(message_dir).as_posix(),
                    "size": path.stat().st_size,
                    "sha256": sha256_file(path),
                }
            )
    metadata = {
        "archive_version": 1,
        "converter_version": "1.2.1",
        "status": status,
        "version": version,
        "source": {
            "store_id": snapshot.store_id,
            "store_name": snapshot.store_name,
            "store_path": snapshot.store_path,
            "folder_id": snapshot.folder_id,
            "folder_path": snapshot.folder_path,
            "folder_segments": snapshot.folder_segments,
            "segment_ids": snapshot.segment_ids,
            "entry_id": snapshot.entry_id,
            "internet_message_id": snapshot.internet_message_id,
            "conversation_id": snapshot.conversation_id,
            "source_key": snapshot.source_key,
            "content_signature": snapshot.content_signature,
        },
        "message": {
            "subject": snapshot.subject,
            "sender_name": snapshot.sender_name,
            "sender_email": snapshot.sender_email,
            "to": snapshot.to,
            "cc": snapshot.cc,
            "bcc": snapshot.bcc,
            "sent_at": iso(snapshot.sent_at),
            "received_at": iso(snapshot.received_at),
            "created_at": iso(snapshot.created_at),
            "modified_at": iso(snapshot.modified_at),
            "effective_at": iso(snapshot.effective_at),
            "effective_basis": snapshot.effective_basis,
            "direction": snapshot.direction,
        },
        "attachments": [item.to_dict() for item in snapshot.attachments],
        "warnings": warnings,
        "errors": snapshot.errors,
        "files": files,
    }
    path = message_dir / "metadata.json"
    write_json_atomic(path, metadata)
    return path
