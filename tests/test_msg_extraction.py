"""Tests for .msg file extraction via DocumentProcessor."""
import os
import unittest
from unittest.mock import patch, MagicMock
from icharlotte_core.document_processor import DocumentProcessor, ExtractResult, ExtractionMethod


class TestMsgExtraction(unittest.TestCase):
    """Test .msg file extraction in DocumentProcessor."""

    @patch('os.path.exists', return_value=True)
    @patch('icharlotte_core.document_processor.MSG_AVAILABLE', True)
    @patch('icharlotte_core.document_processor.pythoncom', create=True)
    @patch('icharlotte_core.document_processor.win32com_client', create=True)
    def test_extract_msg_returns_subject_and_body(self, mock_win32, mock_pythoncom, mock_exists):
        """Extract text from .msg should return subject + body."""
        mock_item = MagicMock()
        mock_item.Subject = "Test Email Subject"
        mock_item.Body = "This is the email body.\nWith multiple lines."
        mock_item.SenderName = "John Doe"
        mock_item.SentOn = "2026-01-15"

        mock_namespace = MagicMock()
        mock_namespace.OpenSharedItem.return_value = mock_item
        mock_outlook = MagicMock()
        mock_outlook.GetNamespace.return_value = mock_namespace
        mock_win32.Dispatch.return_value = mock_outlook

        processor = DocumentProcessor()
        result = processor.extract_text("C:\\fake\\test_email.msg")

        assert result.success, f"Extraction failed: {result.error}"
        assert "Test Email Subject" in result.text
        assert "This is the email body." in result.text
        assert "John Doe" in result.text
        assert result.extraction_method == ExtractionMethod.NATIVE
        mock_item.Close.assert_called_once_with(0)

    @patch('os.path.exists', return_value=True)
    @patch('icharlotte_core.document_processor.MSG_AVAILABLE', True)
    @patch('icharlotte_core.document_processor.pythoncom', create=True)
    @patch('icharlotte_core.document_processor.win32com_client', create=True)
    def test_extract_msg_com_error_returns_failed(self, mock_win32, mock_pythoncom, mock_exists):
        """If Outlook COM fails, return a failed result, not an exception."""
        mock_win32.Dispatch.side_effect = Exception("Outlook not available")

        processor = DocumentProcessor()
        result = processor.extract_text("C:\\fake\\test.msg")

        assert not result.success
        assert "Outlook" in result.error or "error" in result.error.lower()

    @patch('os.path.exists', return_value=True)
    @patch('icharlotte_core.document_processor.MSG_AVAILABLE', True)
    @patch('icharlotte_core.document_processor.pythoncom', create=True)
    @patch('icharlotte_core.document_processor.win32com_client', create=True)
    def test_extract_msg_html_fallback(self, mock_win32, mock_pythoncom, mock_exists):
        """If Body is empty, fall back to HTMLBody stripped of tags."""
        mock_item = MagicMock()
        mock_item.Subject = "HTML Only Email"
        mock_item.Body = ""
        mock_item.HTMLBody = "<html><body><p>HTML content here</p></body></html>"
        mock_item.SenderName = "Jane"
        mock_item.SentOn = "2026-01-20"

        mock_namespace = MagicMock()
        mock_namespace.OpenSharedItem.return_value = mock_item
        mock_outlook = MagicMock()
        mock_outlook.GetNamespace.return_value = mock_namespace
        mock_win32.Dispatch.return_value = mock_outlook

        processor = DocumentProcessor()
        result = processor.extract_text("C:\\fake\\html_email.msg")

        assert result.success
        assert "HTML content here" in result.text
        mock_item.Close.assert_called_once_with(0)


if __name__ == "__main__":
    unittest.main()
