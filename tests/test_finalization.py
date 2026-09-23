import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from outlook_archiver.archive import rebuild_indexes
from outlook_archiver.outlook_reader import OutlookSession
from outlook_archiver.gui import App


class FinalizationTests(unittest.TestCase):
    def test_preexisting_stores_are_never_removed(self):
        session = OutlookSession()
        namespace = session.namespace = MagicMock()
        with patch('outlook_archiver.outlook_reader.pythoncom.CoUninitialize'):
            session.__exit__()
        namespace.RemoveStore.assert_not_called()

    def test_owned_store_removed_and_transient_failure_retried(self):
        session = OutlookSession()
        namespace = session.namespace = MagicMock()
        owned = session._added_pst_root = object()
        namespace.RemoveStore.side_effect = [RuntimeError('busy'), None]
        with patch('outlook_archiver.outlook_reader.pythoncom.CoUninitialize'), \
             patch('outlook_archiver.outlook_reader.time.sleep'):
            session.__exit__()
        self.assertEqual(namespace.RemoveStore.call_count, 2)
        namespace.RemoveStore.assert_called_with(owned)
        self.assertEqual(session.cleanup_warnings, [])

    def test_cleanup_failure_is_not_silenced(self):
        session = OutlookSession()
        session.namespace = MagicMock()
        session._added_pst_root = object()
        session.namespace.RemoveStore.side_effect = RuntimeError('busy')
        with patch('outlook_archiver.outlook_reader.pythoncom.CoUninitialize'), \
             patch('outlook_archiver.outlook_reader.time.sleep'), patch('builtins.print'):
            session.__exit__()
        self.assertEqual(len(session.cleanup_warnings), 1)
        self.assertIn('busy', session.cleanup_warnings[0])

    def test_pure_markdown_does_not_walk_archive_tree(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / 'archive.json').write_text(json.dumps({'content_profile': {
                'save_msg': False, 'save_attachments': False, 'save_inline_images': False}}))
            with patch('outlook_archiver.archive.os.walk', side_effect=AssertionError('unneeded scan')):
                self.assertEqual(rebuild_indexes(root)['messages'], 0)

    def test_report_available_before_run_complete(self):
        app = MagicMock()
        app.pending_result = None
        App._handle_line(app, 'EVENT ' + json.dumps({'event': 'report_ready', 'report': 'early.md', 'failures': 'fail.csv'}))
        self.assertEqual(app.last_report, 'early.md')
        self.assertIsNone(app.pending_result)
        self.assertIn('索引', app.completion_status.set.call_args.args[0])
