import unittest

import pytest

pytest.importorskip("pytestqt")
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from icharlotte_core.discovery.form_interrogatory_selection import ScannedInterrogatory
from icharlotte_core.ui.form_interrogatory_dialog import FormInterrogatorySelectionDialog


class FormInterrogatorySelectionDialogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._app = QApplication.instance() or QApplication([])

    def _items(self):
        return [
            ScannedInterrogatory("1.1", checked=True, text="State the name, ADDRESS..."),
            ScannedInterrogatory("2.1", checked=False, text="State your name..."),
            ScannedInterrogatory("2.2", checked=True, text="State the date and place of your birth."),
        ]

    def test_prechecks_autodetected_and_returns_selected(self):
        dlg = FormInterrogatorySelectionDialog(self._items())
        self.assertEqual(dlg.selected_numbers(), ["1.1", "2.2"])

    def test_toggling_an_item_changes_selection(self):
        dlg = FormInterrogatorySelectionDialog(self._items())
        dlg.checkboxes["2.1"].setChecked(True)
        self.assertEqual(dlg.selected_numbers(), ["1.1", "2.1", "2.2"])

    def test_select_none_then_all(self):
        dlg = FormInterrogatorySelectionDialog(self._items())
        dlg.set_all_checked(False)
        self.assertEqual(dlg.selected_numbers(), [])
        dlg.set_all_checked(True)
        self.assertEqual(dlg.selected_numbers(), ["1.1", "2.1", "2.2"])

    def test_numbers_sorted_naturally(self):
        items = [
            ScannedInterrogatory("12.1", checked=True),
            ScannedInterrogatory("2.10", checked=True),
            ScannedInterrogatory("2.2", checked=True),
        ]
        dlg = FormInterrogatorySelectionDialog(items)
        self.assertEqual(dlg.selected_numbers(), ["2.2", "2.10", "12.1"])


if __name__ == "__main__":
    unittest.main()
