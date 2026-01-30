"""
Tests for Word Hotkey formatting and "Use All Text" functionality.

Run with: python tests/test_word_hotkey_formatting.py
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import unittest
from unittest.mock import MagicMock, patch

# Test imports
print("=" * 60)
print("Testing Word Hotkey Formatting Features")
print("=" * 60)


class TestFormatConstants(unittest.TestCase):
    """Test that format constants are defined correctly."""

    def test_format_constants_exist(self):
        """Test all format constants are defined."""
        from icharlotte_core.word_hotkey import (
            FORMAT_PLAIN, FORMAT_MATCH, FORMAT_MARKDOWN,
            FORMAT_DEFAULT, FORMAT_CUSTOM
        )

        self.assertEqual(FORMAT_PLAIN, "Plain Text")
        self.assertEqual(FORMAT_MATCH, "Match Selection")
        self.assertEqual(FORMAT_MARKDOWN, "Parse Markdown")
        self.assertEqual(FORMAT_DEFAULT, "Document Default")
        self.assertEqual(FORMAT_CUSTOM, "Custom...")
        print("  [PASS] Format constants defined correctly")


class TestCustomFormatDialog(unittest.TestCase):
    """Test CustomFormatDialog class."""

    @classmethod
    def setUpClass(cls):
        """Set up QApplication for widget tests."""
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance()
        if not cls.app:
            cls.app = QApplication([])

    def test_dialog_creation(self):
        """Test dialog can be created."""
        from icharlotte_core.word_hotkey import CustomFormatDialog

        dialog = CustomFormatDialog()
        self.assertIsNotNone(dialog)
        print("  [PASS] CustomFormatDialog created successfully")
        dialog.close()

    def test_dialog_default_values(self):
        """Test dialog has correct default values."""
        from icharlotte_core.word_hotkey import CustomFormatDialog

        dialog = CustomFormatDialog()
        settings = dialog.get_settings()

        self.assertIn('font_name', settings)
        self.assertIn('font_size', settings)
        self.assertIn('bold', settings)
        self.assertIn('italic', settings)
        self.assertIn('underline', settings)
        self.assertIn('first_indent', settings)
        self.assertIn('left_indent', settings)
        self.assertIn('line_spacing', settings)

        # Check default values
        self.assertEqual(settings['font_size'], 12)
        self.assertEqual(settings['bold'], False)
        self.assertEqual(settings['italic'], False)
        self.assertEqual(settings['underline'], False)
        print("  [PASS] CustomFormatDialog default values correct")
        dialog.close()

    def test_dialog_load_settings(self):
        """Test dialog can load existing settings."""
        from icharlotte_core.word_hotkey import CustomFormatDialog

        test_settings = {
            'font_name': 'Arial',
            'font_size': 14,
            'bold': True,
            'italic': False,
            'underline': True,
            'first_indent': 0.5,
            'left_indent': 1.0,
            'line_spacing': 'Double'
        }

        dialog = CustomFormatDialog(current_settings=test_settings)
        retrieved = dialog.get_settings()

        self.assertEqual(retrieved['font_size'], 14)
        self.assertEqual(retrieved['bold'], True)
        self.assertEqual(retrieved['underline'], True)
        self.assertEqual(retrieved['first_indent'], 0.5)
        print("  [PASS] CustomFormatDialog loads settings correctly")
        dialog.close()


class TestMarkdownParsing(unittest.TestCase):
    """Test markdown parsing functionality."""

    @classmethod
    def setUpClass(cls):
        """Set up QApplication for widget tests."""
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance()
        if not cls.app:
            cls.app = QApplication([])

    def test_parse_bold(self):
        """Test parsing bold markdown."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()
        segments = popup._parse_markdown_segments("This is **bold** text")

        # Should have 3 segments: "This is ", "bold", " text"
        self.assertEqual(len(segments), 3)
        self.assertEqual(segments[0]['text'], "This is ")
        self.assertEqual(segments[0]['bold'], False)
        self.assertEqual(segments[1]['text'], "bold")
        self.assertEqual(segments[1]['bold'], True)
        self.assertEqual(segments[2]['text'], " text")
        print("  [PASS] Bold markdown parsing works")
        popup.close()

    def test_parse_italic(self):
        """Test parsing italic markdown."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()
        segments = popup._parse_markdown_segments("This is *italic* text")

        self.assertEqual(len(segments), 3)
        self.assertEqual(segments[1]['text'], "italic")
        self.assertEqual(segments[1]['italic'], True)
        print("  [PASS] Italic markdown parsing works")
        popup.close()

    def test_parse_underline(self):
        """Test parsing underline markdown."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()
        segments = popup._parse_markdown_segments("This is _underlined_ text")

        self.assertEqual(len(segments), 3)
        self.assertEqual(segments[1]['text'], "underlined")
        self.assertEqual(segments[1]['underline'], True)
        print("  [PASS] Underline markdown parsing works")
        popup.close()

    def test_parse_code(self):
        """Test parsing code markdown."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()
        segments = popup._parse_markdown_segments("This is `code` text")

        self.assertEqual(len(segments), 3)
        self.assertEqual(segments[1]['text'], "code")
        self.assertEqual(segments[1]['code'], True)
        print("  [PASS] Code markdown parsing works")
        popup.close()

    def test_parse_plain_text(self):
        """Test parsing text without markdown."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()
        segments = popup._parse_markdown_segments("Plain text without formatting")

        self.assertEqual(len(segments), 1)
        self.assertEqual(segments[0]['text'], "Plain text without formatting")
        self.assertEqual(segments[0]['bold'], False)
        self.assertEqual(segments[0]['italic'], False)
        print("  [PASS] Plain text parsing works")
        popup.close()

    def test_parse_multiple_formats(self):
        """Test parsing text with multiple formats."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()
        segments = popup._parse_markdown_segments("**Bold** and *italic* and `code`")

        # Should have 5 segments
        bold_segment = next(s for s in segments if s.get('bold'))
        italic_segment = next(s for s in segments if s.get('italic'))
        code_segment = next(s for s in segments if s.get('code'))

        self.assertEqual(bold_segment['text'], "Bold")
        self.assertEqual(italic_segment['text'], "italic")
        self.assertEqual(code_segment['text'], "code")
        print("  [PASS] Multiple format parsing works")
        popup.close()


class TestWordLLMPopupUI(unittest.TestCase):
    """Test WordLLMPopup UI components."""

    @classmethod
    def setUpClass(cls):
        """Set up QApplication for widget tests."""
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance()
        if not cls.app:
            cls.app = QApplication([])

    def test_popup_creation(self):
        """Test popup can be created."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()
        self.assertIsNotNone(popup)
        print("  [PASS] WordLLMPopup created successfully")
        popup.close()

    def test_format_combo_exists(self):
        """Test format combo box exists with correct items."""
        from icharlotte_core.word_hotkey import (
            WordLLMPopup, FORMAT_PLAIN, FORMAT_MATCH,
            FORMAT_MARKDOWN, FORMAT_DEFAULT, FORMAT_CUSTOM
        )

        popup = WordLLMPopup()

        self.assertTrue(hasattr(popup, 'format_combo'))

        # Check all format options are in combo
        items = [popup.format_combo.itemText(i) for i in range(popup.format_combo.count())]
        self.assertIn(FORMAT_PLAIN, items)
        self.assertIn(FORMAT_MATCH, items)
        self.assertIn(FORMAT_MARKDOWN, items)
        self.assertIn(FORMAT_DEFAULT, items)
        self.assertIn(FORMAT_CUSTOM, items)
        print("  [PASS] Format combo box has all options")
        popup.close()

    def test_use_all_text_checkbox_exists(self):
        """Test 'Use All Text' checkbox exists."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()

        self.assertTrue(hasattr(popup, 'use_all_text_check'))
        self.assertFalse(popup.use_all_text_check.isChecked())  # Default unchecked
        print("  [PASS] 'Use All Text' checkbox exists and unchecked by default")
        popup.close()

    def test_format_settings_button_exists(self):
        """Test format settings button exists and is disabled by default."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()

        self.assertTrue(hasattr(popup, 'format_settings_btn'))
        self.assertFalse(popup.format_settings_btn.isEnabled())  # Disabled by default
        print("  [PASS] Format settings button exists and disabled by default")
        popup.close()

    def test_format_settings_button_enabled_for_custom(self):
        """Test format settings button is enabled when Custom is selected."""
        from icharlotte_core.word_hotkey import WordLLMPopup, FORMAT_CUSTOM

        popup = WordLLMPopup()

        # Find and select Custom format
        index = popup.format_combo.findText(FORMAT_CUSTOM)
        popup.format_combo.setCurrentIndex(index)

        self.assertTrue(popup.format_settings_btn.isEnabled())
        print("  [PASS] Format settings button enabled for Custom format")
        popup.close()

    def test_format_preview_updates(self):
        """Test format preview label updates when format changes."""
        from icharlotte_core.word_hotkey import (
            WordLLMPopup, FORMAT_PLAIN, FORMAT_MARKDOWN
        )

        popup = WordLLMPopup()

        # Select Plain Text
        index = popup.format_combo.findText(FORMAT_PLAIN)
        popup.format_combo.setCurrentIndex(index)
        self.assertIn("No formatting", popup.format_preview.text())

        # Select Parse Markdown
        index = popup.format_combo.findText(FORMAT_MARKDOWN)
        popup.format_combo.setCurrentIndex(index)
        self.assertIn("Markdown", popup.format_preview.text())

        print("  [PASS] Format preview updates correctly")
        popup.close()


class TestGetAllDocumentText(unittest.TestCase):
    """Test _get_all_document_text method."""

    @classmethod
    def setUpClass(cls):
        """Set up QApplication for widget tests."""
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance()
        if not cls.app:
            cls.app = QApplication([])

    def test_method_exists(self):
        """Test _get_all_document_text method exists."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()
        self.assertTrue(hasattr(popup, '_get_all_document_text'))
        self.assertTrue(callable(popup._get_all_document_text))
        print("  [PASS] _get_all_document_text method exists")
        popup.close()

    def test_returns_string(self):
        """Test _get_all_document_text returns a string."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()
        result = popup._get_all_document_text()
        self.assertIsInstance(result, str)
        print("  [PASS] _get_all_document_text returns string")
        popup.close()


class TestInsertWithFormat(unittest.TestCase):
    """Test _insert_with_format method."""

    @classmethod
    def setUpClass(cls):
        """Set up QApplication for widget tests."""
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance()
        if not cls.app:
            cls.app = QApplication([])

    def test_method_exists(self):
        """Test _insert_with_format method exists."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()
        self.assertTrue(hasattr(popup, '_insert_with_format'))
        self.assertTrue(callable(popup._insert_with_format))
        print("  [PASS] _insert_with_format method exists")
        popup.close()


class TestCaptureSelectionFormat(unittest.TestCase):
    """Test _capture_selection_format method."""

    @classmethod
    def setUpClass(cls):
        """Set up QApplication for widget tests."""
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance()
        if not cls.app:
            cls.app = QApplication([])

    def test_method_exists(self):
        """Test _capture_selection_format method exists."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()
        self.assertTrue(hasattr(popup, '_capture_selection_format'))
        self.assertTrue(callable(popup._capture_selection_format))
        print("  [PASS] _capture_selection_format method exists")
        popup.close()


class TestModelSelection(unittest.TestCase):
    """Test model selection functionality."""

    @classmethod
    def setUpClass(cls):
        """Set up QApplication for widget tests."""
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance()
        if not cls.app:
            cls.app = QApplication([])

    def test_available_models_defined(self):
        """Test AVAILABLE_MODELS is defined with models."""
        from icharlotte_core.word_hotkey import AVAILABLE_MODELS

        self.assertIsInstance(AVAILABLE_MODELS, list)
        self.assertGreater(len(AVAILABLE_MODELS), 0)
        print(f"  [PASS] AVAILABLE_MODELS has {len(AVAILABLE_MODELS)} models")

    def test_model_tuple_format(self):
        """Test each model is a tuple with (display_name, provider, model_id)."""
        from icharlotte_core.word_hotkey import AVAILABLE_MODELS

        for model in AVAILABLE_MODELS:
            self.assertIsInstance(model, tuple)
            self.assertEqual(len(model), 3)
            display_name, provider, model_id = model
            self.assertIsInstance(display_name, str)
            self.assertIsInstance(provider, str)
            self.assertIsInstance(model_id, str)
            self.assertIn(provider, ["Gemini", "Claude", "OpenAI"])

        print("  [PASS] All models have correct tuple format")

    def test_model_combo_exists(self):
        """Test model combo box exists in popup."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()
        self.assertTrue(hasattr(popup, 'model_combo'))
        print("  [PASS] Model combo box exists")
        popup.close()

    def test_model_combo_has_all_models(self):
        """Test model combo box has all available models."""
        from icharlotte_core.word_hotkey import WordLLMPopup, AVAILABLE_MODELS

        popup = WordLLMPopup()
        self.assertEqual(popup.model_combo.count(), len(AVAILABLE_MODELS))
        print(f"  [PASS] Model combo has {popup.model_combo.count()} items")
        popup.close()

    def test_get_selected_model_method(self):
        """Test _get_selected_model method exists and works."""
        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()
        self.assertTrue(hasattr(popup, '_get_selected_model'))

        provider, model_id = popup._get_selected_model()
        self.assertIsInstance(provider, str)
        self.assertIsInstance(model_id, str)
        self.assertIn(provider, ["Gemini", "Claude", "OpenAI"])
        print(f"  [PASS] _get_selected_model returns ({provider}, {model_id})")
        popup.close()

    def test_model_selection_changes(self):
        """Test that changing model selection updates internal state."""
        from icharlotte_core.word_hotkey import WordLLMPopup, AVAILABLE_MODELS

        popup = WordLLMPopup()

        # Select a different model
        new_index = min(1, len(AVAILABLE_MODELS) - 1)
        popup.model_combo.setCurrentIndex(new_index)

        self.assertEqual(popup.selected_model_index, new_index)
        expected_provider = AVAILABLE_MODELS[new_index][1]
        expected_model = AVAILABLE_MODELS[new_index][2]

        provider, model_id = popup._get_selected_model()
        self.assertEqual(provider, expected_provider)
        self.assertEqual(model_id, expected_model)
        print("  [PASS] Model selection changes correctly")
        popup.close()

    def test_default_model_index(self):
        """Test DEFAULT_MODEL_INDEX is valid."""
        from icharlotte_core.word_hotkey import DEFAULT_MODEL_INDEX, AVAILABLE_MODELS

        self.assertIsInstance(DEFAULT_MODEL_INDEX, int)
        self.assertGreaterEqual(DEFAULT_MODEL_INDEX, 0)
        self.assertLess(DEFAULT_MODEL_INDEX, len(AVAILABLE_MODELS))
        print(f"  [PASS] DEFAULT_MODEL_INDEX ({DEFAULT_MODEL_INDEX}) is valid")


def run_tests():
    """Run all tests."""
    # Create test suite
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()

    # Add test classes
    suite.addTests(loader.loadTestsFromTestCase(TestFormatConstants))
    suite.addTests(loader.loadTestsFromTestCase(TestCustomFormatDialog))
    suite.addTests(loader.loadTestsFromTestCase(TestMarkdownParsing))
    suite.addTests(loader.loadTestsFromTestCase(TestWordLLMPopupUI))
    suite.addTests(loader.loadTestsFromTestCase(TestGetAllDocumentText))
    suite.addTests(loader.loadTestsFromTestCase(TestInsertWithFormat))
    suite.addTests(loader.loadTestsFromTestCase(TestCaptureSelectionFormat))
    suite.addTests(loader.loadTestsFromTestCase(TestModelSelection))

    # Run tests
    runner = unittest.TextTestRunner(verbosity=0)
    result = runner.run(suite)

    # Summary
    print("\n" + "=" * 60)
    print("Test Summary")
    print("=" * 60)
    print(f"  Tests run: {result.testsRun}")
    print(f"  Failures: {len(result.failures)}")
    print(f"  Errors: {len(result.errors)}")

    if result.failures:
        print("\nFailures:")
        for test, traceback in result.failures:
            print(f"  - {test}: {traceback}")

    if result.errors:
        print("\nErrors:")
        for test, traceback in result.errors:
            print(f"  - {test}: {traceback}")

    if result.wasSuccessful():
        print("\nAll tests passed!")
    else:
        print("\nSome tests failed.")

    return result.wasSuccessful()


if __name__ == "__main__":
    success = run_tests()
    sys.exit(0 if success else 1)
