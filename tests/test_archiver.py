from __future__ import annotations

import csv
import json
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path

from outlook_archiver.archive import date_filter, export_archive, load_config, verify_archive
from outlook_archiver.models import AttachmentRecord, Candidate, MessageSnapshot
from outlook_archiver.pst import prepare_pst_copy, work_root_for_output
from outlook_archiver.utils import ensure_within, parse_years, safe_name, sha256_file, sha256_text, unique_actual_name


class FakeSession:
    fail_attachment = False

    def __init__(self, _pst: str | None = None):
        self.candidates = []
        for index in range(1, 9):
            year = 2025 if index <= 4 else 2026
            subject = "" if index == 3 else ("中文 English 邮件 " + f"entry-{index}")
            self.candidates.append(
                Candidate(
                    store_id="fake-store", store_name="虚构测试邮箱", store_path="",
                    folder_id="folder-" + ("child" if index > 4 else "root"),
                    folder_path="收件箱 / 项目资料" if index > 4 else "收件箱",
                    folder_segments=["收件箱", "项目资料"] if index > 4 else ["收件箱"],
                    folder_kind="inbox", entry_id=f"entry-{index}",
                    effective_at=datetime(year, 1, min(index, 28), 10, 0, tzinfo=timezone.utc),
                    effective_basis="ReceivedTime", item_class=43,
                    last_modified=datetime(2026, 1, 1, tzinfo=timezone.utc), size=100 + index,
                    subject=subject,
                )
            )

    def __enter__(self): return self
    def __exit__(self, *_): return None

    def scan(self, *, include_date, limit=None, **_):
        selected = [value for value in self.candidates if include_date(value.effective_at)]
        if limit:
            selected = selected[:limit]
        stats = {
            "folders_selected": 2, "items_total": 9, "mail_total": 8, "non_mail": 1,
            "date_outside": 8 - len(selected), "unknown_time": 0, "warnings": [], "folders": [],
        }
        return selected, stats

    @staticmethod
    def inspect_candidate(candidate):
        source_key = sha256_text(candidate.store_id + "\0" + candidate.entry_id)
        return source_key, sha256_text(source_key + "-signature"), "same-internet-id" if candidate.entry_id in {"entry-1", "entry-2"} else ""

    def prepare_message(
        self, candidate, staging, *, save_msg=True, save_attachments=True, save_inline_images=True, **_
    ):
        source_key, signature, internet_id = self.inspect_candidate(candidate)
        staging.mkdir(parents=True, exist_ok=True)
        if save_msg:
            (staging / "original.msg").write_bytes(b"FAKE-MSG\0" + candidate.entry_id.encode())
        attachments = staging / "attachments"; images = staging / "images"
        attachments.mkdir(); images.mkdir()
        records = []
        if candidate.entry_id in {"entry-2", "entry-6"}:
            for number in (1, 2):
                actual = f"{number:03d}_同名附件.txt"
                status = "not_selected" if not save_attachments else ("failed" if candidate.entry_id == "entry-6" and number == 2 and self.fail_attachment else "saved")
                target = attachments / actual
                if status == "saved": target.write_text(f"attachment {number}", encoding="utf-8")
                records.append(AttachmentRecord(number, "同名附件.txt", actual, f"attachments/{actual}", size=target.stat().st_size if target.exists() else 0, sha256=sha256_file(target) if target.exists() else "", status=status, error="模拟附件读取失败" if status == "failed" else ""))
        if candidate.entry_id == "entry-4":
            target = images / "001_签名图片.png"; target.write_bytes(b"\x89PNG\r\n\x1a\nFAKE")
            status = "saved" if save_inline_images else "not_selected"
            if not save_inline_images:
                target.unlink()
            records.append(AttachmentRecord(1, "签名图片.png", target.name, f"images/{target.name}", content_id="logo-1", mime_type="image/png", hidden=True, inline=True, size=target.stat().st_size if target.exists() else 0, sha256=sha256_file(target) if target.exists() else "", status=status))
        subject = candidate.subject
        html = "<p>完整正文 <b>bold</b></p>"
        if candidate.entry_id == "entry-1": html += '<img src="https://example.invalid/tracker.png" alt="远程图">'
        if candidate.entry_id == "entry-4": html += '<p>图：</p><img src="cid:logo-1" alt="签名">'
        if candidate.entry_id == "entry-5": html += "<table><tr><th>项目</th><th>值</th></tr><tr><td>A</td><td>42</td></tr></table>"
        snapshot = MessageSnapshot(
            source_key=source_key, content_signature=signature, store_id=candidate.store_id,
            store_name=candidate.store_name, store_path="", folder_id=candidate.folder_id,
            folder_path=candidate.folder_path, folder_segments=candidate.folder_segments,
            folder_kind="inbox", entry_id=candidate.entry_id, internet_message_id=internet_id,
            conversation_id="conversation", subject=subject, sender_name="测试发件人",
            sender_email="sender@example.invalid", to=["收件人 <to@example.invalid>"], cc=[], bcc=[],
            sent_at=candidate.effective_at, received_at=candidate.effective_at,
            created_at=candidate.effective_at, modified_at=datetime(2026, 1, 1, tzinfo=timezone.utc),
            effective_at=candidate.effective_at, effective_basis="ReceivedTime", direction="incoming",
            html_body=html, text_body="完整正文\r\n历史引用保留", attachments=records,
        )
        snapshot.errors.extend(record.error for record in records if record.status == "failed")
        return snapshot


class ArchiverTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="outlook-archiver-test-")
        self.root = Path(self.temp.name) / "archive"
        self.config = {
            "source_mode": "outlook", "store_id": "fake-store", "folder_ids": ["folder-root"],
            "include_subfolders": True, "output_root": str(self.root), "layout": "folder_first",
            "include_addresses": True, "years": [],
        }
        FakeSession.fail_attachment = False

    def tearDown(self): self.temp.cleanup()

    def factory(self, pst=None): return FakeSession(pst)

    def test_batch_repeat_partial_recovery_and_verification(self):
        events = []
        first = export_archive(self.config, limit=5, emit=events.append, session_factory=self.factory)
        self.assertEqual(first["selected"], 5)
        self.assertEqual(first["success"] + first["warning"], 5)
        self.assertEqual(len(list(self.root.rglob("metadata.json"))), 5)

        FakeSession.fail_attachment = True
        second = export_archive(self.config, emit=events.append, session_factory=self.factory)
        self.assertEqual(second["selected"], 8)  # 正式模式未永久限制为 5 封
        self.assertEqual(second["skipped"], 5)
        self.assertEqual(second["partial"], 1)
        self.assertEqual(second["success"] + second["warning"], 2)
        with Path(second["failures"]).open(encoding="utf-8-sig", newline="") as stream:
            failures = list(csv.DictReader(stream))
        self.assertEqual(len(failures), 1)
        self.assertIn("中文 English 邮件 entry-6", failures[0]["subject"])
        self.assertEqual(failures[0]["stage"], "附件或邮件内容导出")
        self.assertIn("模拟附件读取失败", failures[0]["error"])

        FakeSession.fail_attachment = False
        third = export_archive(self.config, emit=events.append, session_factory=self.factory)
        self.assertEqual(third["skipped"], 7)
        self.assertEqual(third["success"], 1)
        result = verify_archive(self.root, full_hash=True)
        self.assertEqual(result["failed"], 0)
        self.assertEqual(result["checked"], 8)

        with (self.root / "indexes" / "emails.csv").open(encoding="utf-8-sig") as stream:
            emails = list(csv.DictReader(stream))
        self.assertEqual(len(emails), 8)
        self.assertEqual(sum(row["subject"] == "(无主题)" for row in emails), 1)
        # 相同 InternetMessageID 的两个不同源项目均保留。
        stored = [json.loads(path.read_text(encoding="utf-8-sig")) for path in self.root.rglob("metadata.json") if "state" not in path.parts]
        self.assertEqual(sum(item["source"]["internet_message_id"] == "same-internet-id" for item in stored), 2)
        cid_md = next(path for path in self.root.rglob("email.md") if "entry-4" in (path.parent / "metadata.json").read_text(encoding="utf-8-sig"))
        text = cid_md.read_text(encoding="utf-8-sig")
        self.assertIn("images/001_", text)
        self.assertNotIn("cid:logo-1", text)
        table_md = next(path for path in self.root.rglob("email.md") if "entry-5" in (path.parent / "metadata.json").read_text(encoding="utf-8-sig"))
        self.assertIn("| 项目 | 值 |", table_md.read_text(encoding="utf-8-sig"))
        first_dir = next(path.parent for path in self.root.rglob("metadata.json") if "entry-1" in path.read_text(encoding="utf-8-sig"))
        self.assertEqual(first_dir.name, "2025-01-01_中文 English 邮件 entry-1")

    def test_markdown_only_direct_files_and_same_name_suffix(self):
        class SameNameSession(FakeSession):
            def __init__(self, pst=None):
                super().__init__(pst)
                self.candidates[1].effective_at = self.candidates[0].effective_at
                self.candidates[1].folder_id = self.candidates[0].folder_id
                self.candidates[1].folder_path = self.candidates[0].folder_path
                self.candidates[1].folder_segments = list(self.candidates[0].folder_segments)
                self.candidates[0].subject = self.candidates[1].subject = "同名邮件"

        config = dict(
            self.config,
            save_msg=False,
            save_attachments=False,
            save_inline_images=False,
            save_body_sources=False,
        )
        factory = lambda pst=None: SameNameSession(pst)
        first = export_archive(config, limit=2, emit=lambda _: None, session_factory=factory)
        self.assertEqual(first["success"] + first["warning"], 2)
        files = sorted(path for path in self.root.rglob("*.md") if "state" not in path.parts and "logs" not in path.parts and "indexes" not in path.parts and path.name != "00_开始这里.md")
        self.assertEqual([path.name for path in files], ["2025-01-01_同名邮件 (2).md", "2025-01-01_同名邮件.md"])
        self.assertFalse(list(self.root.rglob("original.msg")))
        self.assertFalse(list(self.root.rglob("metadata.json")))
        self.assertFalse(list(self.root.rglob("body.txt")))
        repeated = export_archive(config, limit=2, emit=lambda _: None, session_factory=factory)
        self.assertEqual(repeated["skipped"], 2)
        verified = verify_archive(self.root, full_hash=True)
        self.assertEqual((verified["checked"], verified["failed"]), (2, 0))
        with (self.root / "indexes" / "emails.csv").open(encoding="utf-8-sig", newline="") as stream:
            self.assertEqual(len(list(csv.DictReader(stream))), 2)

    def test_content_profile_is_fixed_for_one_archive_root(self):
        export_archive(self.config, limit=1, emit=lambda _: None, session_factory=self.factory)
        changed = dict(self.config, save_msg=False, save_attachments=False, save_inline_images=False, save_body_sources=False)
        with self.assertRaisesRegex(RuntimeError, "另一组.*导出内容"):
            export_archive(changed, limit=1, emit=lambda _: None, session_factory=self.factory)

    def test_stop_before_first_item_leaves_all_unprocessed(self):
        stop = Path(self.temp.name) / "stop.request"; stop.write_text("stop")
        result = export_archive(self.config, stop_file=stop, emit=lambda _: None, session_factory=self.factory)
        self.assertEqual(result["unprocessed"], 8)
        self.assertEqual(result["success"], 0)

    def test_layout_is_fixed(self):
        export_archive(self.config, limit=1, emit=lambda _: None, session_factory=self.factory)
        changed = dict(self.config, layout="year_first")
        with self.assertRaisesRegex(RuntimeError, "已固定"):
            export_archive(changed, limit=1, emit=lambda _: None, session_factory=self.factory)

    def test_date_filters_and_path_safety(self):
        self.assertEqual(parse_years("2026, 2025；2026"), [2025, 2026])
        fn = date_filter({"years": [2025]})
        self.assertTrue(fn(datetime(2025, 12, 31, tzinfo=timezone.utc)))
        self.assertFalse(fn(datetime(2026, 1, 1, tzinfo=timezone.utc)))
        self.assertNotRegex(safe_name("../../CON:<x>?"), r"[<>:\"/\\|?*]")
        with self.assertRaises(ValueError): ensure_within(self.root, self.root / ".." / "outside")
        attachment_dir = self.root / "attachment-names"
        attachment_dir.mkdir(parents=True)
        first = unique_actual_name(attachment_dir, "同名附件.txt", 1)
        (attachment_dir / first).write_text("one", encoding="utf-8")
        second = unique_actual_name(attachment_dir, "同名附件.txt", 2)
        self.assertEqual((first, second), ("同名附件.txt", "同名附件 (2).txt"))


class PstCopyTests(unittest.TestCase):
    def test_work_root_follows_output_drive_and_parent(self):
        with tempfile.TemporaryDirectory(prefix="pst-work-location-") as temp:
            output = Path(temp) / "archive-parent" / "MailArchive"
            self.assertEqual(work_root_for_output(output), output.resolve().parent / "PSTWork")

    def test_copy_is_verified_and_reused(self):
        with tempfile.TemporaryDirectory(prefix="pst-copy-test-") as temp:
            base = Path(temp); source = base / "源备份.pst"; source.write_bytes(b"!BDN" + b"PST-FAKE" * 1000)
            first = prepare_pst_copy(source, work_root=base / "work")
            self.assertFalse(first.reused); self.assertEqual(sha256_file(source), sha256_file(first.copy))
            second = prepare_pst_copy(source, work_root=base / "work", source_open_in_outlook=True)
            self.assertTrue(second.reused)

    def test_verified_copy_in_legacy_output_root_is_reused(self):
        with tempfile.TemporaryDirectory(prefix="pst-legacy-reuse-") as temp:
            base = Path(temp); source = base / "历史邮件.pst"; source.write_bytes(b"!BDN" + b"PST-LEGACY" * 1000)
            legacy_root = base / "MD-Archive"
            first = prepare_pst_copy(source, work_root=legacy_root)
            # Simulate legitimate Outlook internal changes after the initially verified copy was mounted.
            with first.copy.open("r+b") as stream:
                stream.seek(128)
                stream.write(b"OUTLOOK-UPDATED")
            self.assertNotEqual(sha256_file(first.copy), first.sha256)
            current_root = base / "PSTWork"
            reused = prepare_pst_copy(
                source,
                work_root=current_root,
                reuse_roots=[legacy_root],
                source_open_in_outlook=True,
            )
            self.assertTrue(reused.reused)
            self.assertEqual(reused.copy, first.copy)
            self.assertFalse(list(current_root.rglob("working_*.pst")))


if __name__ == "__main__":
    unittest.main()
