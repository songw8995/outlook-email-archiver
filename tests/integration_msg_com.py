"""真实经典 Outlook COM 的无邮箱数据 MSG 集成测试。

仅创建内存中的虚构 MailItem 并 SaveAs 到系统临时目录；不调用 MailItem.Save，
因此不会进入草稿箱，也不发送、移动或修改用户邮件。
"""

from __future__ import annotations

import json
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

from outlook_archiver.archive import export_archive, verify_archive


def main() -> int:
    with tempfile.TemporaryDirectory(prefix="outlook-msg-integration-") as temp:
        base = Path(temp)
        source_dir = base / "虚构MSG来源"
        source_dir.mkdir()
        text_attachment = source_dir / "同名附件.txt"
        text_attachment.write_text("虚构附件，不含用户数据。", encoding="utf-8")
        png_attachment = source_dir / "正文图片.png"
        png_attachment.write_bytes(
            bytes.fromhex("89504E470D0A1A0A0000000D49484452000000010000000108060000001F15C4890000000D49444154789C6360F8CFC0000004010100A9F6451A0000000049454E44AE426082")
        )
        msg_path = source_dir / "中文 English 表格和图片.msg"
        pythoncom.CoInitialize()
        try:
            app = win32com.client.Dispatch("Outlook.Application")
            mail = app.CreateItem(0)
            mail.Subject = "虚构测试：中文 English HTML table CID"
            mail.To = "example@example.invalid"
            mail.Body = "虚构正文\r\nQuoted history remains."
            mail.Attachments.Add(str(text_attachment))
            inline = mail.Attachments.Add(str(png_attachment))
            inline.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "integration-image")
            inline.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", True)
            mail.HTMLBody = "<p>虚构正文 <b>English</b></p><table><tr><th>项目</th><th>值</th></tr><tr><td>A</td><td>42</td></tr></table><p><img src='cid:integration-image' alt='正文图片'></p>"
            mail.SaveAs(str(msg_path), 9)
            mail = None
            app = None
        finally:
            pythoncom.CoUninitialize()
        msd_path = source_dir / "同内容错误扩展名.msd"
        shutil.copy2(msg_path, msd_path)
        output = base / "archive"
        config = {
            "source_mode": "msg", "msg_paths": [str(source_dir)], "store_id": "msg-import",
            "folder_ids": [], "include_subfolders": True, "years": [], "output_root": str(output),
            "layout": "folder_first", "include_addresses": True,
        }
        events = []
        result = export_archive(config, emit=events.append)
        repeated = export_archive(config, emit=events.append)
        verified = verify_archive(output, full_hash=True, open_msg=True)
        if result["selected"] != 2 or result["failed"] or result["partial"] or repeated["skipped"] != 2 or verified["failed"]:
            failure_text = Path(result["failures"]).read_text(encoding="utf-8-sig") if Path(result["failures"]).exists() else ""
            print(json.dumps({"export": result, "repeat": repeated, "verify": verified, "failure_csv": failure_text}, ensure_ascii=False, indent=2))
            return 1
        markdown_files = list(output.rglob("email.md"))
        if len(markdown_files) != 2 or not all("| 项目 | 值 |" in path.read_text(encoding="utf-8-sig") for path in markdown_files):
            return 2
        if not all("images/%E6%AD%A3%E6%96%87%E5%9B%BE%E7%89%87.png" in path.read_text(encoding="utf-8-sig") for path in markdown_files):
            return 3
        pure_output = base / "markdown-only"
        pure_config = dict(
            config,
            output_root=str(pure_output),
            save_msg=False,
            save_attachments=False,
            save_inline_images=False,
            save_body_sources=False,
        )
        pure_result = export_archive(pure_config, emit=events.append)
        pure_repeat = export_archive(pure_config, emit=events.append)
        pure_verified = verify_archive(pure_output, full_hash=True)
        pure_files = [path for path in pure_output.rglob("*.md") if path.name != "00_开始这里.md" and "logs" not in path.parts and "indexes" not in path.parts]
        if pure_result["selected"] != 2 or pure_repeat["skipped"] != 2 or pure_verified["failed"] or len(pure_files) != 2:
            return 4
        if list(pure_output.rglob("original.msg")) or list(pure_output.rglob("metadata.json")):
            return 5
        if not all(path.name[:10].count("-") == 2 and path.suffix == ".md" for path in pure_files):
            return 6
        print(json.dumps({"exported": 2, "repeat_skipped": repeated["skipped"], "verified": verified["passed"], "msg": 1, "msd_compatible": 1, "cid": "passed", "table": "passed", "markdown_only": 2, "markdown_only_repeat_skipped": pure_repeat["skipped"]}, ensure_ascii=False))
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
