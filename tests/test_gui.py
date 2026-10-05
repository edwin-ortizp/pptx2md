from pathlib import Path
import tempfile
from threading import Event
import unittest
from unittest.mock import Mock, patch

from pptx2md.gui import Api


class GuiBridgeTests(unittest.TestCase):
    def setUp(self):
        self.api = Api()
        self.api._window = Mock()
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)

    def test_selection_multiple_dedup_removal_and_cancelled_dialog(self):
        first, second = self.root / 'one.pptx', self.root / 'two.pdf'
        self.api._window.create_file_dialog.return_value = [str(first), str(second), str(first)]
        state = self.api.add_files()
        self.assertEqual([f['path'] for f in state['selection']], [str(first), str(second)])
        self.assertTrue(self.api._window.create_file_dialog.call_args.kwargs['allow_multiple'])
        state = self.api.remove_file(str(first))
        self.assertEqual(len(state['selection']), 1)
        self.api._window.create_file_dialog.return_value = None
        self.assertEqual(len(self.api.add_files()['selection']), 1)
        self.assertFalse(self.api.get_state()['busy'])

    def test_choose_folder_and_open_only_registered_outputs(self):
        self.api._window.create_file_dialog.return_value = [str(self.root)]
        self.assertEqual(self.api.choose_output_dir()['output_dir'], str(self.root))
        output = self.root / 'done.md'
        output.write_text('done', encoding='utf-8')
        self.api._state['items'] = [{'output': str(output)}]
        self.api._state['output_dir'] = str(self.root)
        with patch.object(self.api, '_open', return_value={'ok': True}) as opener:
            self.assertTrue(self.api.open_result(str(output))['ok'])
            self.assertTrue(self.api.open_output_dir()['ok'])
            self.assertIn('error', self.api.open_result(str(self.root / 'other.md')))
            self.assertEqual(opener.call_count, 2)

    def test_background_conversion_lock_and_final_state(self):
        entered, release, finished = Event(), Event(), Event()
        self.api._sources = [self.root / 'one.pdf']
        def conversion(sources, options, update):
            entered.set()
            release.wait(5)
            update(dict(self.api._state, running=False, processed=1, succeeded=1))
            finished.set()
        with patch('pptx2md.gui.run_batch', side_effect=conversion):
            state = self.api.start_conversion({'mode': 'individual'})
            self.assertTrue(entered.wait(2))
            try:
                self.assertTrue(state['busy'])
                self.assertTrue(state['batch']['running'])
                self.assertIn('error', self.api.start_conversion({}))
                self.assertIn('error', self.api.remove_file(str(self.api._sources[0])))
                self.assertIn('error', self.api.add_files())
                self.assertIn('error', self.api.choose_output_dir())
                self.assertFalse(self.api._can_close())
                self.assertEqual(self.api.get_state()['batch']['processed'], 0)
            finally:
                release.set()
            self.assertTrue(finished.wait(2))
        # The worker clears busy just after publishing its final snapshot.
        import time
        for _ in range(100):
            if not self.api.get_state()['busy']:
                break
            time.sleep(.01)
        self.assertFalse(self.api.get_state()['busy'])
        self.assertTrue(self.api._can_close())
        self.assertEqual(self.api.get_state()['batch']['processed'], 1)

    def test_input_validation(self):
        self.assertIn('error', self.api.start_conversion({}))
        self.api._sources = [self.root / 'one.pdf']
        self.assertIn('error', self.api.start_conversion({'mode': 'combined', 'combined_name': '../bad.md'}))
        self.assertFalse(self.api.get_state()['busy'])
