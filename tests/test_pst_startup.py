"""Regression tests: do not start Outlook before securing a PST work copy."""
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from outlook_archiver.outlook_reader import OutlookSession
from outlook_archiver.pst import PstCopy, _exclusive_reader, prepare_pst_copy


class PstStartupTests(unittest.TestCase):
    def test_copy_precedes_outlook_activation(self):
        events = []
        copy = PstCopy(Path('source.pst'), Path('working.pst'), True, 'hash')
        app = MagicMock()
        with patch('outlook_archiver.outlook_reader.prepare_pst_copy',
                   side_effect=lambda *a, **k: (events.append('copy') or copy)), \
             patch('outlook_archiver.outlook_reader.win32com.client.Dispatch',
                   side_effect=lambda *a: (events.append('dispatch') or app)), \
             patch.object(OutlookSession, '_mount_pst_work_copy',
                          side_effect=lambda *a: events.append('mount')):
            with OutlookSession('source.pst') as session:
                self.assertIs(session.pst_copy, copy)
        self.assertEqual(events, ['copy', 'dispatch', 'mount'])

    def test_failed_copy_never_starts_outlook(self):
        with patch('outlook_archiver.outlook_reader.prepare_pst_copy',
                   side_effect=RuntimeError('locked')), \
             patch('outlook_archiver.outlook_reader.win32com.client.Dispatch') as dispatch:
            with self.assertRaisesRegex(RuntimeError, 'locked'):
                with OutlookSession('source.pst'):
                    pass
            dispatch.assert_not_called()

    @unittest.skipUnless(os.name == 'nt', 'Windows file sharing test')
    def test_real_windows_lock_and_copy_reuse(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / 'source.pst'
            source.write_bytes(b'!BDN' + b'fake PST test data' * 100)
            root = Path(tmp) / 'work'
            with _exclusive_reader(source):
                with self.assertRaisesRegex(RuntimeError, 'Windows.*32|Windows.*33'):
                    prepare_pst_copy(source, work_root=root)
            result = prepare_pst_copy(source, work_root=root)
            self.assertFalse(result.reused)
            self.assertEqual(source.read_bytes(), result.copy.read_bytes())
            with _exclusive_reader(source):
                self.assertTrue(prepare_pst_copy(source, work_root=root).reused)

    def test_write_error_is_not_reported_as_source_lock(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / 'source.pst'
            source.write_bytes(b'!BDN' + b'fake PST')
            original = Path.open
            def denied(path, *args, **kwargs):
                if path.suffix == '.partial':
                    raise PermissionError('target permission denied')
                return original(path, *args, **kwargs)
            with patch.object(Path, 'open', denied):
                with self.assertRaisesRegex(RuntimeError, 'target permission denied') as error:
                    prepare_pst_copy(source, work_root=Path(tmp) / 'work')
            self.assertNotIn('正被', str(error.exception))


if __name__ == '__main__':
    unittest.main()
