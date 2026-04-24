"""Tests for ChatTab._extract_doc_text — legacy .doc file extraction.

The helper uses Word COM automation and must NOT call word.Quit() or touch
word.Visible (the user may have Word open with their own documents).
"""
import os
import sys
import unittest
from unittest.mock import MagicMock, patch

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from PySide6.QtWidgets import QApplication
app = QApplication.instance() or QApplication(sys.argv)

from icharlotte_core.ui.tabs import ChatTab


def _mock_com_modules(doc_text: str = "Hello from .doc file."):
    """Build fake pythoncom + win32com.client modules and the fake Word COM tree.

    Returns (pythoncom_mock, win32_mock, word_mock, doc_mock) so tests can
    assert on any of them.
    """
    doc_mock = MagicMock()
    doc_mock.Content.Text = doc_text

    word_mock = MagicMock()
    word_mock.Documents.Open.return_value = doc_mock

    win32_client_mock = MagicMock()
    win32_client_mock.Dispatch.return_value = word_mock
    win32_mock = MagicMock()
    win32_mock.client = win32_client_mock

    pythoncom_mock = MagicMock()

    return pythoncom_mock, win32_mock, word_mock, doc_mock


def _patched_modules(pythoncom_mock, win32_mock):
    """Context manager for injecting fake COM modules into sys.modules."""
    return patch.dict(sys.modules, {
        "pythoncom": pythoncom_mock,
        "win32com": win32_mock,
        "win32com.client": win32_mock.client,
    })


class TestExtractDocText(unittest.TestCase):

    def test_extracts_text_from_doc(self):
        pythoncom_mock, win32_mock, _word, _doc = _mock_com_modules("Page 1 body.")
        with _patched_modules(pythoncom_mock, win32_mock):
            result = ChatTab._extract_doc_text(r"C:\fake\test.doc")
        self.assertIn("Page 1 body.", result)
        self.assertTrue(result.endswith("\n"))

    def test_opens_doc_readonly_and_without_recent_files_entry(self):
        pythoncom_mock, win32_mock, word_mock, _doc = _mock_com_modules()
        with _patched_modules(pythoncom_mock, win32_mock):
            ChatTab._extract_doc_text(r"C:\fake\test.doc")

        word_mock.Documents.Open.assert_called_once()
        _, kwargs = word_mock.Documents.Open.call_args
        self.assertEqual(kwargs.get("ReadOnly"), True)
        self.assertEqual(kwargs.get("AddToRecentFiles"), False)
        self.assertEqual(kwargs.get("ConfirmConversions"), False)

    def test_never_calls_word_quit(self):
        """Safety rule: we must never close the user's Word app."""
        pythoncom_mock, win32_mock, word_mock, _doc = _mock_com_modules()
        with _patched_modules(pythoncom_mock, win32_mock):
            ChatTab._extract_doc_text(r"C:\fake\test.doc")
        word_mock.Quit.assert_not_called()

    def test_never_touches_word_visible(self):
        """Don't hide the user's Word if it was already running."""
        pythoncom_mock, win32_mock, word_mock, _doc = _mock_com_modules()
        with _patched_modules(pythoncom_mock, win32_mock):
            ChatTab._extract_doc_text(r"C:\fake\test.doc")
        # No assignment to word_mock.Visible should have happened.
        # MagicMock tracks attribute setattr via _mock_children; check Visible
        # wasn't among explicit calls.
        visible_assignments = [
            call for call in word_mock.mock_calls
            if call[0].startswith("Visible")
        ]
        self.assertEqual(visible_assignments, [])

    def test_closes_document_without_saving(self):
        pythoncom_mock, win32_mock, _word, doc_mock = _mock_com_modules()
        with _patched_modules(pythoncom_mock, win32_mock):
            ChatTab._extract_doc_text(r"C:\fake\test.doc")
        doc_mock.Close.assert_called_once_with(SaveChanges=0)

    def test_document_closed_even_on_extraction_error(self):
        pythoncom_mock, win32_mock, word_mock, doc_mock = _mock_com_modules()

        # Replace Content with an object whose .Text property raises.
        class BrokenContent:
            @property
            def Text(self):
                raise RuntimeError("bad doc state")

        doc_mock.Content = BrokenContent()

        with _patched_modules(pythoncom_mock, win32_mock):
            result = ChatTab._extract_doc_text(r"C:\fake\test.doc")
        self.assertIn("Error reading .doc file", result)
        self.assertIn("bad doc state", result)
        doc_mock.Close.assert_called_once()
        word_mock.Quit.assert_not_called()

    def test_dispatch_failure_is_handled_gracefully(self):
        pythoncom_mock, win32_mock, _word, _doc = _mock_com_modules()
        win32_mock.client.Dispatch.side_effect = RuntimeError("no Word installed")
        with _patched_modules(pythoncom_mock, win32_mock):
            result = ChatTab._extract_doc_text(r"C:\fake\test.doc")
        self.assertIn("could not launch Word", result)
        # CoUninitialize should still run to avoid leaking COM state
        pythoncom_mock.CoUninitialize.assert_called_once()

    def test_open_failure_is_handled_gracefully(self):
        pythoncom_mock, win32_mock, word_mock, _doc = _mock_com_modules()
        word_mock.Documents.Open.side_effect = RuntimeError("file locked")
        with _patched_modules(pythoncom_mock, win32_mock):
            result = ChatTab._extract_doc_text(r"C:\fake\test.doc")
        self.assertIn("Error reading .doc file", result)
        self.assertIn("file locked", result)
        word_mock.Quit.assert_not_called()

    def test_com_initialization_pair(self):
        pythoncom_mock, win32_mock, _word, _doc = _mock_com_modules()
        with _patched_modules(pythoncom_mock, win32_mock):
            ChatTab._extract_doc_text(r"C:\fake\test.doc")
        pythoncom_mock.CoInitialize.assert_called_once()
        pythoncom_mock.CoUninitialize.assert_called_once()


if __name__ == "__main__":
    unittest.main()
