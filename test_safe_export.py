import contextlib
import io
import logging
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import export_vba_code as exporter


class SafeExportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / 'source.xlsm'
        self.source.write_bytes(b'synthetic source')
        self.dest = self.root / 'destination'
        self.dest.mkdir()
        (self.dest / 'Old.frm').write_bytes(b'old form')
        (self.dest / 'Keep.bas').write_bytes(b'old module')
        (self.dest / 'README.md').write_bytes(b'readme')
        (self.dest / 'Notes.txt').write_bytes(b'notes')
        (self.dest / 'nested').mkdir()
        (self.dest / 'nested' / 'Child.bas').write_bytes(b'child')
        self.parser = patch.object(exporter, 'VBA_Parser').start()
        self.addCleanup(patch.stopall)
        self.parser.return_value.extract_macros.return_value = [
            ('', '', 'Zebra.cls', b'new class'),
            ('', '', 'Keep.bas', 'new module'),
        ]

    def snapshot(self):
        return {str(p.relative_to(self.dest)): p.read_bytes() if p.is_file() else None
                for p in self.dest.rglob('*')}

    def run_export(self):
        return exporter._stage_and_replace(str(self.source), str(self.dest))

    def test_success_replaces_all_top_level_files_and_preserves_subdirectories(self):
        result = self.run_export()
        self.assertEqual(result, ['Keep.bas', 'Zebra.cls'])
        self.assertFalse((self.dest / 'Old.frm').exists())
        self.assertEqual((self.dest / 'Keep.bas').read_bytes(), b'new module')
        self.assertFalse((self.dest / 'README.md').exists())
        self.assertFalse((self.dest / 'Notes.txt').exists())
        self.assertEqual((self.dest / 'nested' / 'Child.bas').read_bytes(), b'child')

    def test_parser_txt_output_is_exported(self):
        self.parser.return_value.extract_macros.return_value = [
            ('', '', 'VBA_P-code.txt', b'p-code'),
        ]
        self.assertEqual(self.run_export(), ['VBA_P-code.txt'])
        self.assertEqual((self.dest / 'VBA_P-code.txt').read_bytes(), b'p-code')
        self.assertFalse((self.dest / 'Notes.txt').exists())
        self.assertEqual((self.dest / 'nested' / 'Child.bas').read_bytes(), b'child')
        self.parser.return_value.close.assert_called_once()

    def test_empty_untouched(self):
        before = self.snapshot()
        self.parser.return_value.extract_macros.return_value = []
        self.assertEqual(self.run_export(), [])
        self.assertEqual(before, self.snapshot())

    def test_extraction_failure_untouched(self):
        before = self.snapshot()
        self.parser.return_value.extract_macros.side_effect = RuntimeError('extraction failed')
        with self.assertRaises(RuntimeError):
            self.run_export()
        self.assertEqual(before, self.snapshot())

    def test_logged_extraction_failure_untouched(self):
        before = self.snapshot()
        def modules():
            logging.getLogger('oletools.olevba').error('incomplete project')
            yield ('', '', 'Keep.bas', b'partial')
        self.parser.return_value.extract_macros.side_effect = modules
        with self.assertRaises(ValueError):
            self.run_export()
        self.assertEqual(before, self.snapshot())

    def test_staging_write_failure_untouched(self):
        before = self.snapshot()
        with patch.object(exporter, 'open', side_effect=OSError('write failed'), create=True):
            with self.assertRaises(OSError):
                self.run_export()
        self.assertEqual(before, self.snapshot())

    def test_unsafe_and_duplicate_names(self):
        for names in [('..\\Escape.bas',), ('CON.bas',), ('A.bas', 'a.BAS')]:
            with self.subTest(names=names):
                before = self.snapshot()
                self.parser.return_value.extract_macros.return_value = [('', '', n, b'code') for n in names]
                with self.assertRaises(ValueError):
                    self.run_export()
                self.assertEqual(before, self.snapshot())

    def test_directory_collision_untouched(self):
        (self.dest / 'Zebra.cls').mkdir()
        before = self.snapshot()
        with self.assertRaises(ValueError):
            self.run_export()
        self.assertEqual(before, self.snapshot())

    def test_source_hardlink_untouched(self):
        os.link(self.source, self.dest / 'SourceAlias.bas')
        before = self.snapshot()
        with self.assertRaises(ValueError):
            self.run_export()
        self.assertEqual(before, self.snapshot())
        self.assertEqual(self.source.read_bytes(), b'synthetic source')

    def test_backup_failure_untouched(self):
        before = self.snapshot()
        with patch.object(exporter.shutil, 'copy2', side_effect=OSError('backup failed')):
            with self.assertRaises(OSError):
                self.run_export()
        self.assertEqual(before, self.snapshot())

    def test_replacement_failure_restores_originals(self):
        before = self.snapshot()
        real_copy = exporter.shutil.copyfileobj
        calls = 0
        def fail_once(source, dest):
            nonlocal calls
            calls += 1
            if calls == 2:
                dest.write(b'partial')
                raise OSError('replacement failed')
            return real_copy(source, dest)
        with patch.object(exporter.shutil, 'copyfileobj', side_effect=fail_once):
            with self.assertRaises(OSError):
                self.run_export()
        self.assertEqual(before, self.snapshot())

    def test_replacement_failure_restores_arbitrary_top_level_files(self):
        before = self.snapshot()
        real_copy = exporter.shutil.copyfileobj
        calls = 0

        def fail_once(source, dest):
            nonlocal calls
            calls += 1
            if calls == 2:
                dest.write(b'partial')
                raise OSError('replacement failed')
            return real_copy(source, dest)

        with patch.object(exporter.shutil, 'copyfileobj', side_effect=fail_once):
            with self.assertRaises(OSError):
                self.run_export()
        self.assertEqual(before, self.snapshot())
        self.assertEqual((self.dest / 'README.md').read_bytes(), b'readme')
        self.assertEqual((self.dest / 'Notes.txt').read_bytes(), b'notes')

    def test_failed_recovery_retains_backup(self):
        real_mkdtemp = tempfile.mkdtemp
        created = []
        def tracked_temp(**kwargs):
            path = real_mkdtemp(dir=self.root, **kwargs)
            created.append(Path(path))
            return path
        output = io.StringIO()
        with patch.object(exporter.tempfile, 'mkdtemp', side_effect=tracked_temp), \
                patch.object(exporter.shutil, 'copyfileobj', side_effect=OSError('disk failure')), \
                contextlib.redirect_stdout(output):
            with self.assertRaises(OSError):
                self.run_export()
        self.assertFalse(created[0].exists())
        self.assertTrue(created[1].exists())
        self.assertEqual((created[1] / 'Keep.bas').read_bytes(), b'old module')
        self.assertIn('Backup retained at', output.getvalue())

    def test_cli_commit_behavior(self):
        for tail in [[], ['A multi-word', 'message'], ['nOcOmMiT']]:
            with self.subTest(tail=tail), \
                    patch.object(exporter.sys, 'argv', ['export_vba_code.py', str(self.source), str(self.dest)] + tail), \
                    patch.object(exporter, '_stage_and_replace', return_value=['Keep.bas']), \
                    patch.object(exporter.subprocess, 'run') as git, \
                    contextlib.redirect_stdout(io.StringIO()):
                exporter.export_vba_code_internal()
                if tail == ['nOcOmMiT']:
                    git.assert_not_called()
                else:
                    self.assertEqual(git.call_count, 2)
                    message = git.call_args.args[0][-1]
                    if tail:
                        self.assertEqual(message, 'A multi-word message')
                    else:
                        self.assertRegex(message, r'^Auto-export VBA \d{4}-\d{2}-\d{2} \d{2}:\d{2}$')


if __name__ == '__main__':
    unittest.main()
