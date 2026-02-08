"""Tests for Word redline functionality."""

import unittest
from unittest.mock import Mock, MagicMock, patch, call
import sys


class TestWordRedlineMethod(unittest.TestCase):
    """Test the _apply_redlines method in isolation."""

    @patch('adeu.RedlineEngine')
    def test_apply_redlines_success(self, mock_engine_class):
        """Test successful redline application."""
        # Import the method implementation
        from icharlotte_core.word_hotkey import WordLLMPopup

        # Create mock popup object with just the attributes we need
        mock_popup = Mock()
        mock_popup.redline_settings = {
            "auto_enable_track_changes": True
        }
        mock_popup._original_range_start = 0
        mock_popup._original_range_end = 100

        # Setup Word mocks
        mock_word_app = Mock()
        mock_doc = Mock()
        mock_doc.TrackRevisions = False
        mock_range = Mock()
        mock_doc.Range = Mock(return_value=mock_range)
        mock_word_app.ActiveDocument = mock_doc

        mock_selection = Mock()
        mock_selection.Range = Mock()

        # Setup RedlineEngine mock
        mock_engine = mock_engine_class.return_value
        mock_engine.apply_redlines = Mock()

        # Call the actual method
        result = WordLLMPopup._apply_redlines(
            mock_popup,
            mock_word_app,
            mock_selection,
            "original text",
            "revised text"
        )

        # Verify
        self.assertTrue(result)
        self.assertTrue(mock_doc.TrackRevisions)  # Should be auto-enabled

        # Verify range reconstruction was attempted
        mock_doc.Range.assert_called_once_with(0, 100)

        # Verify RedlineEngine was instantiated and called
        mock_engine_class.assert_called_once()
        mock_engine.apply_redlines.assert_called_once_with(
            doc=mock_doc,
            range_obj=mock_range,
            original="original text",
            revised="revised text"
        )

    @patch('adeu.RedlineEngine')
    def test_apply_redlines_track_changes_already_enabled(self, mock_engine_class):
        """Test redline when Track Changes is already enabled."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        mock_popup = Mock()
        mock_popup.redline_settings = {
            "auto_enable_track_changes": True
        }
        mock_popup._original_range_start = 0
        mock_popup._original_range_end = 100

        # Track Changes already enabled
        mock_word_app = Mock()
        mock_doc = Mock()
        mock_doc.TrackRevisions = True  # Already enabled
        mock_range = Mock()
        mock_doc.Range = Mock(return_value=mock_range)
        mock_word_app.ActiveDocument = mock_doc

        mock_selection = Mock()
        mock_engine = mock_engine_class.return_value
        mock_engine.apply_redlines = Mock()

        result = WordLLMPopup._apply_redlines(
            mock_popup,
            mock_word_app,
            mock_selection,
            "original",
            "revised"
        )

        # Should still succeed
        self.assertTrue(result)
        # Track Changes should remain enabled (not changed)
        self.assertTrue(mock_doc.TrackRevisions)

    @patch('adeu.RedlineEngine')
    def test_apply_redlines_range_reconstruction_fails(self, mock_engine_class):
        """Test fallback to selection when range reconstruction fails."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        mock_popup = Mock()
        mock_popup.redline_settings = {"auto_enable_track_changes": True}
        mock_popup._original_range_start = 0
        mock_popup._original_range_end = 100

        mock_word_app = Mock()
        mock_doc = Mock()
        mock_doc.TrackRevisions = False
        # Range reconstruction fails
        mock_doc.Range = Mock(side_effect=Exception("Range error"))
        mock_word_app.ActiveDocument = mock_doc

        mock_selection = Mock()
        mock_selection_range = Mock()
        mock_selection.Range = mock_selection_range

        mock_engine = mock_engine_class.return_value
        mock_engine.apply_redlines = Mock()

        result = WordLLMPopup._apply_redlines(
            mock_popup,
            mock_word_app,
            mock_selection,
            "original",
            "revised"
        )

        # Should still succeed using selection fallback
        self.assertTrue(result)
        # Should have used selection.Range instead
        mock_engine.apply_redlines.assert_called_once_with(
            doc=mock_doc,
            range_obj=mock_selection_range,
            original="original",
            revised="revised"
        )

    def test_apply_redlines_adeu_not_available(self):
        """Test graceful failure when adeu is not available."""
        from icharlotte_core.word_hotkey import WordLLMPopup
        import builtins

        # Temporarily make adeu import fail
        original_import = builtins.__import__

        def mock_import(name, *args, **kwargs):
            if name == 'adeu':
                raise ImportError("No module named 'adeu'")
            return original_import(name, *args, **kwargs)

        with patch('builtins.__import__', side_effect=mock_import):
            mock_popup = Mock()
            mock_popup.redline_settings = {"auto_enable_track_changes": True}

            result = WordLLMPopup._apply_redlines(
                mock_popup,
                Mock(),
                Mock(),
                "original",
                "revised"
            )

            # Should return False to indicate fallback needed
            self.assertFalse(result)

    @patch('adeu.RedlineEngine')
    def test_apply_redlines_engine_error(self, mock_engine_class):
        """Test graceful failure when RedlineEngine raises an error."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        mock_popup = Mock()
        mock_popup.redline_settings = {"auto_enable_track_changes": True}
        mock_popup._original_range_start = 0
        mock_popup._original_range_end = 100

        mock_word_app = Mock()
        mock_doc = Mock()
        mock_doc.TrackRevisions = False
        mock_doc.Range = Mock(return_value=Mock())
        mock_word_app.ActiveDocument = mock_doc

        # RedlineEngine raises an error
        mock_engine = mock_engine_class.return_value
        mock_engine.apply_redlines = Mock(side_effect=Exception("Redline error"))

        result = WordLLMPopup._apply_redlines(
            mock_popup,
            mock_word_app,
            Mock(),
            "original",
            "revised"
        )

        # Should return False to indicate failure
        self.assertFalse(result)


if __name__ == '__main__':
    unittest.main()
