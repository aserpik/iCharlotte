"""
Integration tests for Word formatting functionality.
These tests require Microsoft Word to be running with a document open.

Run with: python tests/test_word_formatting_integration.py
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    import win32com.client
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False
    print("win32com not available - skipping integration tests")
    sys.exit(0)


def test_word_connection():
    """Test connection to Word."""
    print("\n1. Testing Word connection...")
    try:
        pythoncom.CoInitialize()
        word = win32com.client.GetActiveObject("Word.Application")
        doc_count = word.Documents.Count
        print(f"   Connected to Word with {doc_count} document(s)")

        if doc_count == 0:
            print("   WARNING: No documents open. Please open a Word document.")
            return None

        return word
    except Exception as e:
        print(f"   ERROR: Could not connect to Word: {e}")
        return None


def test_get_all_document_text(word):
    """Test getting all document text."""
    print("\n2. Testing get all document text...")
    try:
        doc = word.ActiveDocument
        content = doc.Content.Text
        text_len = len(content) if content else 0
        print(f"   Document has {text_len} characters")
        if text_len > 0:
            preview = content[:100].replace('\r', '\\r').replace('\n', '\\n')
            print(f"   Preview: '{preview}...'")
        return True
    except Exception as e:
        print(f"   ERROR: {e}")
        return False


def test_capture_selection_format(word):
    """Test capturing selection format."""
    print("\n3. Testing capture selection format...")
    try:
        selection = word.Selection
        font = selection.Font
        para = selection.ParagraphFormat

        format_info = {
            'font_name': font.Name,
            'font_size': font.Size,
            'bold': font.Bold,
            'italic': font.Italic,
            'underline': font.Underline,
            'first_indent': para.FirstLineIndent / 72,
            'left_indent': para.LeftIndent / 72,
        }

        print(f"   Font: {format_info['font_name']}, {format_info['font_size']}pt")
        print(f"   Bold: {format_info['bold']}, Italic: {format_info['italic']}, Underline: {format_info['underline']}")
        print(f"   First Indent: {format_info['first_indent']:.2f}in, Left Indent: {format_info['left_indent']:.2f}in")
        return True
    except Exception as e:
        print(f"   ERROR: {e}")
        return False


def test_apply_format(word):
    """Test applying format to text (non-destructive test)."""
    print("\n4. Testing format application (read-only check)...")
    try:
        selection = word.Selection
        font = selection.Font

        # Just verify we can access font properties without error
        _ = font.Name
        _ = font.Size
        _ = font.Bold
        _ = font.Italic
        _ = font.Underline

        print("   Format properties accessible")
        return True
    except Exception as e:
        print(f"   ERROR: {e}")
        return False


def test_markdown_parsing():
    """Test markdown parsing logic."""
    print("\n5. Testing markdown parsing...")

    from PySide6.QtWidgets import QApplication
    app = QApplication.instance()
    if not app:
        app = QApplication([])

    from icharlotte_core.word_hotkey import WordLLMPopup

    popup = WordLLMPopup()

    test_cases = [
        ("**bold text**", "bold text", "bold"),
        ("*italic text*", "italic text", "italic"),
        ("_underlined text_", "underlined text", "underline"),
        ("`code text`", "code text", "code"),
        ("Normal text", "Normal text", None),
    ]

    all_passed = True
    for markdown, expected_text, expected_style in test_cases:
        segments = popup._parse_markdown_segments(markdown)

        # Find the segment with content
        content_segment = next((s for s in segments if s['text'].strip()), None)

        if content_segment:
            text_match = expected_text in content_segment['text']
            if expected_style:
                style_match = content_segment.get(expected_style, False)
            else:
                # Plain text - should have no formatting
                style_match = not any([
                    content_segment.get('bold'),
                    content_segment.get('italic'),
                    content_segment.get('underline'),
                    content_segment.get('code')
                ])

            if text_match and style_match:
                print(f"   PASS: '{markdown}' -> {expected_style or 'plain'}")
            else:
                print(f"   FAIL: '{markdown}' -> expected {expected_style}, got {content_segment}")
                all_passed = False
        else:
            print(f"   FAIL: '{markdown}' -> no content segment found")
            all_passed = False

    popup.close()
    return all_passed


def test_format_combo_items():
    """Test format combo box has all required items."""
    print("\n6. Testing format combo box items...")

    from PySide6.QtWidgets import QApplication
    app = QApplication.instance()
    if not app:
        app = QApplication([])

    from icharlotte_core.word_hotkey import (
        WordLLMPopup, FORMAT_PLAIN, FORMAT_MATCH,
        FORMAT_MARKDOWN, FORMAT_DEFAULT, FORMAT_CUSTOM
    )

    popup = WordLLMPopup()

    expected_items = [FORMAT_PLAIN, FORMAT_MATCH, FORMAT_MARKDOWN, FORMAT_DEFAULT, FORMAT_CUSTOM]
    actual_items = [popup.format_combo.itemText(i) for i in range(popup.format_combo.count())]

    all_present = True
    for item in expected_items:
        if item in actual_items:
            print(f"   PASS: '{item}' present")
        else:
            print(f"   FAIL: '{item}' missing")
            all_present = False

    popup.close()
    return all_present


def test_use_all_text_checkbox():
    """Test 'Use All Text' checkbox."""
    print("\n7. Testing 'Use All Text' checkbox...")

    from PySide6.QtWidgets import QApplication
    app = QApplication.instance()
    if not app:
        app = QApplication([])

    from icharlotte_core.word_hotkey import WordLLMPopup

    popup = WordLLMPopup()

    # Test checkbox exists
    if not hasattr(popup, 'use_all_text_check'):
        print("   FAIL: Checkbox not found")
        popup.close()
        return False

    print("   PASS: Checkbox exists")

    # Test default state
    if popup.use_all_text_check.isChecked():
        print("   FAIL: Checkbox should be unchecked by default")
        popup.close()
        return False

    print("   PASS: Checkbox unchecked by default")

    # Test checking/unchecking
    popup.use_all_text_check.setChecked(True)
    if not popup.use_all_text_check.isChecked():
        print("   FAIL: Checkbox could not be checked")
        popup.close()
        return False

    print("   PASS: Checkbox can be toggled")

    popup.close()
    return True


def test_custom_format_dialog():
    """Test CustomFormatDialog."""
    print("\n8. Testing CustomFormatDialog...")

    from PySide6.QtWidgets import QApplication
    app = QApplication.instance()
    if not app:
        app = QApplication([])

    from icharlotte_core.word_hotkey import CustomFormatDialog

    # Test creation
    dialog = CustomFormatDialog()
    print("   PASS: Dialog created")

    # Test default settings
    settings = dialog.get_settings()
    required_keys = ['font_name', 'font_size', 'bold', 'italic', 'underline',
                     'first_indent', 'left_indent', 'line_spacing']

    missing_keys = [k for k in required_keys if k not in settings]
    if missing_keys:
        print(f"   FAIL: Missing keys: {missing_keys}")
        dialog.close()
        return False

    print("   PASS: All settings keys present")

    # Test with custom settings
    custom_settings = {
        'font_name': 'Arial',
        'font_size': 16,
        'bold': True,
        'italic': True,
        'underline': False,
        'first_indent': 0.5,
        'left_indent': 1.0,
        'line_spacing': '1.5 Lines'
    }

    dialog2 = CustomFormatDialog(current_settings=custom_settings)
    retrieved = dialog2.get_settings()

    if retrieved['font_size'] != 16:
        print(f"   FAIL: font_size not loaded (got {retrieved['font_size']})")
        dialog.close()
        dialog2.close()
        return False

    if retrieved['bold'] != True:
        print(f"   FAIL: bold not loaded (got {retrieved['bold']})")
        dialog.close()
        dialog2.close()
        return False

    print("   PASS: Settings loaded correctly")

    dialog.close()
    dialog2.close()
    return True


def run_all_tests():
    """Run all integration tests."""
    print("=" * 60)
    print("Word Hotkey Formatting Integration Tests")
    print("=" * 60)

    results = {}

    # Word connection test
    word = test_word_connection()
    results['Word Connection'] = word is not None

    if word:
        results['Get All Document Text'] = test_get_all_document_text(word)
        results['Capture Selection Format'] = test_capture_selection_format(word)
        results['Apply Format'] = test_apply_format(word)
    else:
        print("\n   Skipping Word-dependent tests (no Word connection)")
        results['Get All Document Text'] = None
        results['Capture Selection Format'] = None
        results['Apply Format'] = None

    # UI tests (don't require Word)
    results['Markdown Parsing'] = test_markdown_parsing()
    results['Format Combo Items'] = test_format_combo_items()
    results['Use All Text Checkbox'] = test_use_all_text_checkbox()
    results['Custom Format Dialog'] = test_custom_format_dialog()

    # Summary
    print("\n" + "=" * 60)
    print("Test Results")
    print("=" * 60)

    for name, result in results.items():
        if result is None:
            status = "SKIP"
        elif result:
            status = "PASS"
        else:
            status = "FAIL"
        print(f"  {name}: {status}")

    passed = sum(1 for r in results.values() if r is True)
    failed = sum(1 for r in results.values() if r is False)
    skipped = sum(1 for r in results.values() if r is None)

    print(f"\n  Passed: {passed}, Failed: {failed}, Skipped: {skipped}")

    return failed == 0


if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1)
