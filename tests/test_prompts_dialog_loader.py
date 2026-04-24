"""Tests for PromptsDialog._load_test_input multi-document loader.

Covers the bug fix (PDF loading no longer crashes with UTF-8 decode error)
and the feature (multiple documents concatenated as A/B test context).
"""
import os
import sys
import unittest
from unittest.mock import MagicMock, patch

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from PySide6.QtWidgets import QApplication
app = QApplication.instance() or QApplication(sys.argv)

from icharlotte_core.ui.dialogs import PromptsDialog
from icharlotte_core.document_processor import ExtractResult, ExtractionMethod


def _ok_result(text: str) -> ExtractResult:
    return ExtractResult(
        text=text,
        page_count=1,
        extraction_method=ExtractionMethod.NATIVE,
        char_count=len(text),
    )


def _fail_result(err: str) -> ExtractResult:
    return ExtractResult(
        text="",
        page_count=0,
        extraction_method=ExtractionMethod.FAILED,
        error=err,
    )


class TestExtractTestInputFromFiles(unittest.TestCase):
    """Pure-helper tests — no Qt dialog construction needed."""

    def test_single_pdf_does_not_raise_utf8_error(self):
        # Simulate a PDF: DocumentProcessor returns extracted text (the whole
        # point — we no longer open() binary bytes as UTF-8).
        processor = MagicMock()
        processor.extract_text.return_value = _ok_result("Page 1 text.\nPage 2 text.")

        text, errors = PromptsDialog._extract_test_input_from_files(
            [r"C:\fake\deposition.pdf"], processor
        )

        self.assertIn("Page 1 text.", text)
        self.assertIn("=== FILE: deposition.pdf ===", text)
        self.assertEqual(errors, [])
        processor.extract_text.assert_called_once_with(r"C:\fake\deposition.pdf")

    def test_multiple_docs_are_concatenated_with_headers(self):
        processor = MagicMock()
        processor.extract_text.side_effect = [
            _ok_result("Complaint body here."),
            _ok_result("Medical records body here."),
            _ok_result("Deposition Q&A here."),
        ]

        paths = [
            r"C:\case\complaint.docx",
            r"C:\case\med_records.pdf",
            r"C:\case\depo.pdf",
        ]
        text, errors = PromptsDialog._extract_test_input_from_files(paths, processor)

        self.assertEqual(errors, [])
        # All three headers present
        self.assertIn("=== FILE: complaint.docx ===", text)
        self.assertIn("=== FILE: med_records.pdf ===", text)
        self.assertIn("=== FILE: depo.pdf ===", text)
        # Order preserved
        self.assertLess(text.index("complaint.docx"), text.index("med_records.pdf"))
        self.assertLess(text.index("med_records.pdf"), text.index("depo.pdf"))
        # Contents preserved
        self.assertIn("Complaint body here.", text)
        self.assertIn("Medical records body here.", text)
        self.assertIn("Deposition Q&A here.", text)

    def test_extraction_failure_is_reported_as_error(self):
        processor = MagicMock()
        processor.extract_text.return_value = _fail_result("corrupt PDF")

        text, errors = PromptsDialog._extract_test_input_from_files(
            [r"C:\fake\broken.pdf"], processor
        )

        self.assertEqual(text, "")
        self.assertEqual(len(errors), 1)
        self.assertIn("broken.pdf", errors[0])
        self.assertIn("corrupt PDF", errors[0])

    def test_partial_failure_keeps_good_docs(self):
        processor = MagicMock()
        processor.extract_text.side_effect = [
            _ok_result("Good content."),
            _fail_result("could not read"),
            _ok_result("Also good."),
        ]

        text, errors = PromptsDialog._extract_test_input_from_files(
            [r"C:\a.pdf", r"C:\bad.pdf", r"C:\c.pdf"], processor
        )

        self.assertIn("Good content.", text)
        self.assertIn("Also good.", text)
        self.assertNotIn("bad.pdf", text)  # failed file not in test input
        self.assertEqual(len(errors), 1)
        self.assertIn("bad.pdf", errors[0])

    def test_processor_raising_exception_is_caught(self):
        processor = MagicMock()
        processor.extract_text.side_effect = RuntimeError("oh no")

        text, errors = PromptsDialog._extract_test_input_from_files(
            [r"C:\boom.pdf"], processor
        )

        self.assertEqual(text, "")
        self.assertEqual(len(errors), 1)
        self.assertIn("boom.pdf", errors[0])
        self.assertIn("oh no", errors[0])

    def test_empty_file_list_returns_empty(self):
        processor = MagicMock()
        text, errors = PromptsDialog._extract_test_input_from_files([], processor)
        self.assertEqual(text, "")
        self.assertEqual(errors, [])
        processor.extract_text.assert_not_called()


class TestLoadTestInputDialog(unittest.TestCase):
    """Integration-ish: patch the file dialog + processor, drive the handler."""

    @patch("icharlotte_core.ui.dialogs.QFileDialog.getOpenFileNames")
    @patch("icharlotte_core.document_processor.DocumentProcessor")
    def test_load_test_input_sets_test_input_text(self, mock_proc_cls, mock_dialog):
        mock_dialog.return_value = ([r"C:\fake\a.pdf"], "Documents (*.pdf)")
        mock_proc_cls.return_value.extract_text.return_value = _ok_result("Hello PDF.")

        dlg = PromptsDialog()
        try:
            dlg._load_test_input()
            loaded = dlg.test_input.toPlainText()
            self.assertIn("=== FILE: a.pdf ===", loaded)
            self.assertIn("Hello PDF.", loaded)
        finally:
            dlg.deleteLater()

    @patch("icharlotte_core.ui.dialogs.QFileDialog.getOpenFileNames")
    def test_load_test_input_noop_when_cancelled(self, mock_dialog):
        mock_dialog.return_value = ([], "")

        dlg = PromptsDialog()
        try:
            dlg.test_input.setPlainText("existing content")
            dlg._load_test_input()
            # Content unchanged when dialog cancelled
            self.assertEqual(dlg.test_input.toPlainText(), "existing content")
        finally:
            dlg.deleteLater()


if __name__ == "__main__":
    unittest.main()
