import tempfile
import unittest
import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest

pytest.importorskip("pytestqt")
from PySide6.QtWidgets import QApplication

from icharlotte_core.ui.wizard.pages.output_page import OutputPage


def _make_docx(path: Path, text: str = "Generated response") -> None:
    from docx import Document

    doc = Document()
    doc.add_paragraph(text)
    doc.save(str(path))


class OutputPageSaveBehaviorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._app = QApplication.instance() or QApplication([])

    def test_save_button_is_actionable_after_output_load(self):
        with tempfile.TemporaryDirectory(dir="C:\\geminiterminal2") as tmp:
            path = Path(tmp) / "response.docx"
            _make_docx(path)
            page = OutputPage()

            page.load_output(str(path))

            self.assertTrue(page.save_btn.isEnabled())

    def test_save_without_edits_reports_already_saved(self):
        with tempfile.TemporaryDirectory(dir="C:\\geminiterminal2") as tmp:
            path = Path(tmp) / "response.docx"
            _make_docx(path)
            page = OutputPage()
            page.load_output(str(path))

            with patch(
                "icharlotte_core.ui.wizard.pages.output_page.QMessageBox.information"
            ) as mock_info:
                page.save_btn.click()

            mock_info.assert_called_once()
            self.assertIn("already saved", mock_info.call_args.args[2].lower())

    @patch("icharlotte_core.ui.wizard.pages.output_page.save_qtextdocument_as_docx")
    def test_save_with_edits_writes_and_confirms(self, mock_save):
        with tempfile.TemporaryDirectory(dir="C:\\geminiterminal2") as tmp:
            path = Path(tmp) / "response.docx"
            _make_docx(path)
            page = OutputPage()
            page.load_output(str(path))
            page.editor.setPlainText("Edited response")

            with patch(
                "icharlotte_core.ui.wizard.pages.output_page.QMessageBox.information"
            ) as mock_info:
                page.save_btn.click()

            mock_save.assert_called_once()
            self.assertFalse(page.is_dirty)
            mock_info.assert_called_once()
            self.assertIn("saved", mock_info.call_args.args[1].lower())

    @patch("icharlotte_core.ui.wizard.pages.output_page.shutil.copyfile")
    def test_save_as_mode_prompts_for_location_and_copies_preview(self, mock_copy):
        with tempfile.TemporaryDirectory(dir="C:\\geminiterminal2") as tmp:
            tmp_path = Path(tmp)
            preview = tmp_path / "preview.docx"
            responses = tmp_path / "DISCOVERY" / "RESPONSES"
            _make_docx(preview)
            page = OutputPage()
            page.load_output(str(preview))
            page.set_save_as_defaults(
                default_dir=str(responses),
                suggested_filename="Def Jones's Resp to SI(1).docx",
            )
            chosen = responses / "Custom Response.docx"

            with patch(
                "icharlotte_core.ui.wizard.pages.output_page.QFileDialog.getSaveFileName",
                return_value=(str(chosen), "Word Documents (*.docx)"),
            ) as mock_dialog:
                with patch(
                    "icharlotte_core.ui.wizard.pages.output_page.QMessageBox.information"
                ) as mock_info:
                    page.save_btn.click()

            self.assertEqual(
                mock_dialog.call_args.args[2],
                str(responses / "Def Jones's Resp to SI(1).docx"),
            )
            mock_copy.assert_called_once_with(str(preview), str(chosen))
            self.assertEqual(page.output_path, str(chosen))
            mock_info.assert_called_once()


if __name__ == "__main__":
    unittest.main()
