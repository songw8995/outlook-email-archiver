import csv
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from outlook_archiver.archive import rebuild_indexes


class IncrementalIndexTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        (self.root / 'state' / 'items').mkdir(parents=True)
        (self.root / 'archive.json').write_text(json.dumps({'content_profile': {
            'save_msg': False, 'save_attachments': False, 'save_inline_images': False}}))

    def item(self, name, subject, pending=True):
        (self.root / (name + '.md')).write_text(subject)
        data = {'format': 'markdown_only', 'archive_path': name + '.md',
                'subject': subject, 'effective_at': '2025-01-01', 'store_id': 's', 'folder_id': 'f'}
        p = self.root / 'state' / 'items' / (name + '.json')
        p.write_text(json.dumps(data))
        if pending:
            with (self.root / 'state' / 'index-pending.jsonl').open('a') as stream:
                stream.write(json.dumps({'metadata': p.relative_to(self.root).as_posix()}) + '\n')

    def rows(self):
        with (self.root / 'indexes' / 'emails.csv').open(encoding='utf-8-sig', newline='') as stream:
            return list(csv.DictReader(stream))

    def test_incremental_update_repair_and_repeat(self):
        self.item('old', 'old')
        self.assertEqual(rebuild_indexes(self.root, mode='incremental')['scanned'], 1)
        self.item('new', 'new')
        with patch('outlook_archiver.archive.os.walk', side_effect=AssertionError('full scan')):
            result = rebuild_indexes(self.root, mode='incremental')
        self.assertEqual((result['messages'], result['scanned']), (2, 1))
        self.item('old', 'changed')
        result = rebuild_indexes(self.root, mode='incremental')
        self.assertEqual(len(self.rows()), 2)
        self.assertIn('changed', [r['subject'] for r in self.rows()])
        self.assertEqual(rebuild_indexes(self.root, mode='incremental')['scanned'], 0)

    def test_failed_index_write_keeps_pending_and_replay_is_idempotent(self):
        self.item('old', 'old'); rebuild_indexes(self.root)
        self.item('new', 'new')
        import os
        replace = os.replace
        def fail(src, dst):
            if Path(dst).name == 'emails.csv':
                raise OSError('simulated interruption')
            return replace(src, dst)
        with patch('outlook_archiver.archive.os.replace', side_effect=fail):
            with self.assertRaises(OSError):
                rebuild_indexes(self.root, mode='incremental')
        self.assertTrue((self.root / 'state' / 'index-pending.jsonl').exists())
        self.assertEqual(rebuild_indexes(self.root, mode='incremental')['messages'], 2)
        self.assertEqual(len(self.rows()), 2)

    def test_full_rebuild_reflects_external_deletion(self):
        self.item('old', 'old'); rebuild_indexes(self.root)
        (self.root / 'old.md').unlink()
        self.assertEqual(rebuild_indexes(self.root, mode='incremental')['messages'], 1)
        self.assertEqual(rebuild_indexes(self.root, mode='full')['messages'], 0)

    def test_invalid_csv_falls_back_to_full(self):
        self.item('old', 'old'); rebuild_indexes(self.root)
        (self.root / 'indexes' / 'emails.csv').write_text('broken\n')
        self.assertEqual(rebuild_indexes(self.root, mode='incremental')['scanned'], 1)

    def test_noncanonical_root_keeps_new_index_entries(self):
        self.item('old', 'old'); rebuild_indexes(self.root)
        self.item('new', 'new')
        result = rebuild_indexes(self.root / 'state' / '..', mode='incremental')
        self.assertEqual((result['messages'], result['scanned']), (2, 1))
