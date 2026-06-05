"""Integration test for PromptsDialog._save_to_current().

Reproduces the reported bug: clicking "Save to Current" on a versioned prompt
(e.g. word_assistant) popped the "Enter version description" dialog and created
a NEW version, instead of overwriting the currently selected version in place.
"""
import os
import shutil
import sys
import tempfile
import unittest
from unittest.mock import patch

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from PySide6.QtWidgets import QApplication

app = QApplication.instance() or QApplication(sys.argv)

from icharlotte_core.prompt_manager import PromptManager
from icharlotte_core.ui import dialogs as dialogs_module
from icharlotte_core.ui.dialogs import PromptsDialog


class TestSaveToCurrentHandler(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.pm = PromptManager(prompts_dir=self.tmp)
        # A versioned prompt with no legacy file — the case that was broken.
        self.pm.create_version("word_assistant", "system_prompt", "ORIGINAL",
                               version="v1", set_as_current=True)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _make_dialog(self):
        dlg = PromptsDialog()
        # Point the dialog at our isolated, temp-backed prompt store.
        dlg.prompt_manager = self.pm
        dlg.current_agent = "word_assistant"
        dlg.current_pass = "system_prompt"
        dlg.current_version = "v1"
        # Drive editor content without touching the widget internals.
        dlg._get_editor_raw_content = lambda: "EDITED"
        # Refresh repopulates combos from the temp store; not under test here.
        dlg._populate_versions = lambda: None
        return dlg

    def test_save_to_current_overwrites_selected_version_without_prompting(self):
        dlg = self._make_dialog()
        try:
            with patch.object(dialogs_module.QInputDialog, "getText",
                              return_value=("desc", True)) as get_text, \
                 patch.object(dialogs_module.QMessageBox, "information"), \
                 patch.object(dialogs_module.QMessageBox, "warning"), \
                 patch.object(dialogs_module.QMessageBox, "critical"):
                dlg._save_to_current()

            # It must NOT ask for a new version description.
            get_text.assert_not_called()
            # The selected version was overwritten in place...
            self.assertEqual(
                self.pm.get_prompt("word_assistant", "system_prompt", "v1"), "EDITED")
            # ...and no new version was created.
            versions = self.pm.list_versions("word_assistant", "system_prompt")
            self.assertEqual(len(versions), 1)
            # ...and the active "current" prompt reflects the edit.
            self.assertEqual(
                self.pm.get_prompt("word_assistant", "system_prompt", "current"), "EDITED")
        finally:
            dlg.deleteLater()


if __name__ == "__main__":
    unittest.main()
