# Word Redline Integration Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add redline mode to Word AI Assistant so AI suggestions appear as Track Changes instead of replacing text completely.

**Architecture:** Extend WordLLMPopup with a checkbox that switches between replace mode (existing) and redline mode (new). In redline mode, use adeu's RedlineEngine to inject Track Changes XML into the document instead of using selection.TypeText().

**Tech Stack:** Python, PySide6/PyQt6, win32com, adeu RedlineEngine, pytest

---

## Task 1: Add adeu Dependency

**Files:**
- Modify: `requirements.txt`

**Step 1: Add adeu to requirements**

Add at the end of `requirements.txt`:
```
adeu>=0.1.0  # Agentic DOCX Redlining Engine
```

**Step 2: Install adeu**

Run: `pip install adeu`
Expected: Package installs successfully

**Step 3: Verify adeu imports**

Run: `python -c "from adeu import RedlineEngine; print('OK')"`
Expected: Prints "OK"

**Step 4: Commit**

```bash
git add requirements.txt
git commit -m "feat: add adeu dependency for Word redlining"
```

---

## Task 2: Add Redline Configuration Settings

**Files:**
- Read: `icharlotte_core/word_hotkey.py:522-529` (to understand settings paths)
- Create: `icharlotte_core/redline_config.py`

**Step 1: Create redline configuration module**

Create `icharlotte_core/redline_config.py`:
```python
"""Configuration for Word redline functionality."""

import os
import json
from typing import Dict, Any

# Default redline settings
DEFAULT_REDLINE_SETTINGS = {
    "redline_mode_default": False,
    "auto_enable_track_changes": True,
    "redline_fallback_notify": True,
    "max_redline_text_length": 50000
}

def load_redline_settings(config_dir: str) -> Dict[str, Any]:
    """Load redline settings from JSON file.

    Args:
        config_dir: Directory containing configuration files

    Returns:
        Dictionary of redline settings
    """
    settings_path = os.path.join(config_dir, "redline_settings.json")

    if os.path.exists(settings_path):
        try:
            with open(settings_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading redline settings: {e}")

    return DEFAULT_REDLINE_SETTINGS.copy()

def save_redline_settings(config_dir: str, settings: Dict[str, Any]) -> bool:
    """Save redline settings to JSON file.

    Args:
        config_dir: Directory containing configuration files
        settings: Dictionary of redline settings to save

    Returns:
        True if saved successfully, False otherwise
    """
    settings_path = os.path.join(config_dir, "redline_settings.json")

    try:
        os.makedirs(config_dir, exist_ok=True)
        with open(settings_path, 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=2)
        return True
    except Exception as e:
        print(f"Error saving redline settings: {e}")
        return False
```

**Step 2: Write test for configuration loading**

Create `tests/test_redline_config.py`:
```python
"""Tests for redline configuration."""

import os
import json
import tempfile
import unittest
from icharlotte_core.redline_config import (
    load_redline_settings,
    save_redline_settings,
    DEFAULT_REDLINE_SETTINGS
)

class TestRedlineConfig(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def test_load_default_settings(self):
        """Test loading default settings when no file exists."""
        settings = load_redline_settings(self.temp_dir)
        self.assertEqual(settings, DEFAULT_REDLINE_SETTINGS)

    def test_save_and_load_settings(self):
        """Test saving and loading custom settings."""
        custom_settings = {
            "redline_mode_default": True,
            "auto_enable_track_changes": False,
            "redline_fallback_notify": True,
            "max_redline_text_length": 100000
        }

        # Save
        result = save_redline_settings(self.temp_dir, custom_settings)
        self.assertTrue(result)

        # Load
        loaded = load_redline_settings(self.temp_dir)
        self.assertEqual(loaded, custom_settings)

if __name__ == '__main__':
    unittest.main()
```

**Step 3: Run test**

Run: `python -m pytest tests/test_redline_config.py -v`
Expected: 2 tests pass

**Step 4: Commit**

```bash
git add icharlotte_core/redline_config.py tests/test_redline_config.py
git commit -m "feat: add redline configuration module with persistence"
```

---

## Task 3: Add Redline Checkbox to WordLLMPopup UI

**Files:**
- Modify: `icharlotte_core/word_hotkey.py:518-600` (WordLLMPopup.__init__)

**Step 1: Import redline config at top of file**

Add after other imports (around line 30):
```python
from .redline_config import load_redline_settings, save_redline_settings
```

**Step 2: Add redline settings loading in __init__**

In `WordLLMPopup.__init__()` around line 528, add:
```python
# Load redline settings
self.redline_settings = load_redline_settings(GEMINI_DATA_DIR)
```

**Step 3: Add checkbox widget creation**

Find where the UI is built (after line 538). Before calling `self._build_ui()` or in the UI building section, add:
```python
# Redline mode checkbox (only visible for Word context)
self.redline_checkbox = QCheckBox("✏️ Use Redline Mode (Track Changes)")
self.redline_checkbox.setToolTip(
    "Instead of replacing text, insert AI suggestions as Track Changes "
    "that you can accept/reject in Word"
)
self.redline_checkbox.setChecked(self.redline_settings.get("redline_mode_default", False))
```

**Step 4: Add checkbox to layout**

Find the layout section where buttons are added (search for "Process" button). Add checkbox before the button row:
```python
# Add redline checkbox (will be shown/hidden based on context)
layout.addWidget(self.redline_checkbox)
```

**Step 5: Add visibility control based on context**

Find where `self.app_context` is set or checked. Add a method to update checkbox visibility:
```python
def _update_redline_checkbox_visibility(self):
    """Show/hide redline checkbox based on app context."""
    is_word = self.app_context == APP_CONTEXT_WORD
    self.redline_checkbox.setVisible(is_word)
```

Call this method after setting app_context (search for `self.app_context =`).

**Step 6: Manual test**

Run app and trigger Word hotkey (Win+V).
Expected: Checkbox appears in dialog with proper label and tooltip.

**Step 7: Commit**

```bash
git add icharlotte_core/word_hotkey.py
git commit -m "feat: add redline mode checkbox to Word AI popup UI"
```

---

## Task 4: Add Checkbox State Persistence

**Files:**
- Modify: `icharlotte_core/word_hotkey.py` (WordLLMPopup class)

**Step 1: Add method to save checkbox state**

Add this method to `WordLLMPopup` class:
```python
def _save_redline_preference(self):
    """Save the current redline checkbox state to settings."""
    self.redline_settings["redline_mode_default"] = self.redline_checkbox.isChecked()
    save_redline_settings(GEMINI_DATA_DIR, self.redline_settings)
```

**Step 2: Connect checkbox to save on change**

In `__init__` after creating the checkbox, add:
```python
self.redline_checkbox.stateChanged.connect(self._save_redline_preference)
```

**Step 3: Manual test persistence**

Run app, check the checkbox, close dialog, reopen dialog.
Expected: Checkbox remains checked across sessions.

**Step 4: Commit**

```bash
git add icharlotte_core/word_hotkey.py
git commit -m "feat: persist redline checkbox state across sessions"
```

---

## Task 5: Capture Original Text for Redlining

**Files:**
- Modify: `icharlotte_core/word_hotkey.py` (WordLLMPopup class)

**Step 1: Add instance variables for redline data**

In `WordLLMPopup.__init__()` around line 531, add:
```python
# Redline state
self._original_text = None  # Original text before LLM processing
self._original_range_start = None  # Range start position
self._original_range_end = None  # Range end position
self._redline_mode_active = False  # Whether current operation uses redline
```

**Step 2: Find where Word text is captured**

Search for `_get_word_text` method. After capturing the text, store it:

In `_get_word_text()` method, after `text = selection.Text`, add:
```python
# Store original text and range for potential redlining
if self.redline_checkbox.isChecked():
    self._original_text = text
    try:
        self._original_range_start = selection.Range.Start
        self._original_range_end = selection.Range.End
        self._redline_mode_active = True
    except Exception as e:
        print(f"Could not capture range coordinates: {e}")
        self._redline_mode_active = False
else:
    self._redline_mode_active = False
```

**Step 3: Pass redline flag to worker thread**

Find where `LLMWorkerThread` is created and started. Modify to pass the flag. Add this to the thread's `__init__`:
```python
self.redline_mode = redline_mode
```

And in the popup's process method, pass it:
```python
self._worker_thread = LLMWorkerThread(
    self.llm_callback,
    full_prompt,
    parent=self
)
# Store redline state on thread for result handler
self._worker_thread.redline_mode = self._redline_mode_active
```

**Step 4: Commit**

```bash
git add icharlotte_core/word_hotkey.py
git commit -m "feat: capture original text and range for redlining"
```

---

## Task 6: Implement Redline Application Method

**Files:**
- Modify: `icharlotte_core/word_hotkey.py` (WordLLMPopup class)

**Step 1: Add redline application method**

Add this method to `WordLLMPopup` class:
```python
def _apply_redlines(self, word_app, selection, original_text: str, revised_text: str) -> bool:
    """Apply redlines using adeu RedlineEngine.

    Args:
        word_app: Word COM application object
        selection: Word Selection object
        original_text: Original text before LLM processing
        revised_text: LLM-revised text

    Returns:
        True if redlines applied successfully, False if fallback needed
    """
    try:
        from adeu import RedlineEngine

        # Get document object
        doc = word_app.ActiveDocument

        # Auto-enable Track Changes if needed
        if self.redline_settings.get("auto_enable_track_changes", True):
            if not doc.TrackRevisions:
                print("Auto-enabling Track Changes for redlining")
                doc.TrackRevisions = True

        # Reconstruct range from stored coordinates
        try:
            range_obj = doc.Range(
                self._original_range_start,
                self._original_range_end
            )
        except Exception as e:
            print(f"Could not reconstruct range: {e}, using current selection")
            range_obj = selection.Range

        # Apply redlines using adeu
        engine = RedlineEngine()
        engine.apply_redlines(
            doc=doc,
            range_obj=range_obj,
            original=original_text,
            revised=revised_text
        )

        print(f"Successfully applied redlines ({len(original_text)} -> {len(revised_text)} chars)")
        return True

    except ImportError as e:
        print(f"adeu not available: {e}")
        return False
    except Exception as e:
        print(f"Redline failed: {e}")
        return False
```

**Step 2: Write unit test with mock**

Create `tests/test_word_redline.py`:
```python
"""Tests for Word redline functionality."""

import unittest
from unittest.mock import Mock, MagicMock, patch
from icharlotte_core.word_hotkey import WordLLMPopup

class TestWordRedline(unittest.TestCase):

    def setUp(self):
        self.popup = WordLLMPopup()
        self.popup.redline_settings = {
            "auto_enable_track_changes": True
        }

    @patch('icharlotte_core.word_hotkey.RedlineEngine')
    def test_apply_redlines_success(self, mock_engine_class):
        """Test successful redline application."""
        # Setup mocks
        mock_word_app = Mock()
        mock_doc = Mock()
        mock_doc.TrackRevisions = False
        mock_doc.Range = Mock(return_value=Mock())
        mock_word_app.ActiveDocument = mock_doc

        mock_selection = Mock()
        mock_selection.Range = Mock()

        mock_engine = mock_engine_class.return_value
        mock_engine.apply_redlines = Mock()

        # Store range coordinates
        self.popup._original_range_start = 0
        self.popup._original_range_end = 100

        # Apply redlines
        result = self.popup._apply_redlines(
            mock_word_app,
            mock_selection,
            "original text",
            "revised text"
        )

        # Verify
        self.assertTrue(result)
        self.assertTrue(mock_doc.TrackRevisions)  # Should be enabled
        mock_engine.apply_redlines.assert_called_once()

if __name__ == '__main__':
    unittest.main()
```

**Step 3: Run test**

Run: `python -m pytest tests/test_word_redline.py -v`
Expected: Test passes

**Step 4: Commit**

```bash
git add icharlotte_core/word_hotkey.py tests/test_word_redline.py
git commit -m "feat: implement redline application with adeu engine"
```

---

## Task 7: Integrate Redline into Result Handler

**Files:**
- Modify: `icharlotte_core/word_hotkey.py` (WordLLMPopup result handler)

**Step 1: Find the result handler method**

Search for `_on_llm_finished` or similar method that handles LLM worker completion.

**Step 2: Add branching logic for redline vs replace**

In the result handler, find where `_set_word_text` or `selection.TypeText` is called. Replace with:
```python
# Branch based on redline mode
if hasattr(self._worker_thread, 'redline_mode') and self._worker_thread.redline_mode:
    # Redline mode - use adeu
    success = self._apply_redlines(
        self._word_app,
        selection,
        self._original_text,
        result_text
    )

    if success:
        self.status_label.setText("✓ Redlines applied!")
        self.status_label.setStyleSheet("color: #a6e3a1; font-style: italic;")
    else:
        # Fallback to replace mode
        if self.redline_settings.get("redline_fallback_notify", True):
            self.status_label.setText("⚠ Applied as replacement (redline unavailable)")
            self.status_label.setStyleSheet("color: #f9e2af; font-style: italic;")
        selection.TypeText(result_text)
else:
    # Replace mode (existing behavior)
    selection.TypeText(result_text)
```

**Step 3: Manual test with Word**

Run app, select text in Word, trigger hotkey, check redline checkbox, process.
Expected: Track Changes appear in Word document.

**Step 4: Commit**

```bash
git add icharlotte_core/word_hotkey.py
git commit -m "feat: integrate redline mode into LLM result handler"
```

---

## Task 8: Handle Edge Case - No Selection

**Files:**
- Modify: `icharlotte_core/word_hotkey.py` (WordLLMPopup class)

**Step 1: Add selection validation method**

Add method to `WordLLMPopup`:
```python
def _validate_redline_prerequisites(self, selection_text: str) -> tuple[bool, str]:
    """Check if redline mode can be used.

    Args:
        selection_text: Currently selected text

    Returns:
        (is_valid, error_message) tuple
    """
    # Check if text is selected
    if not selection_text or len(selection_text.strip()) == 0:
        return False, "Select text first to use redline mode"

    # Check text length limit
    max_length = self.redline_settings.get("max_redline_text_length", 50000)
    if len(selection_text) > max_length:
        return False, f"Selection too large for redline ({len(selection_text)} > {max_length} chars)"

    return True, ""
```

**Step 2: Update checkbox state on text capture**

In `_get_word_text()` method, after capturing text:
```python
# Validate redline prerequisites
if self.redline_checkbox.isChecked():
    is_valid, error_msg = self._validate_redline_prerequisites(text)
    if not is_valid:
        self.redline_checkbox.setEnabled(False)
        self.redline_checkbox.setToolTip(error_msg)
        self._redline_mode_active = False
    else:
        self.redline_checkbox.setEnabled(True)
        # ... existing redline setup code
```

**Step 3: Write test for validation**

Add to `tests/test_word_redline.py`:
```python
def test_validate_no_selection(self):
    """Test validation fails with no text selected."""
    is_valid, msg = self.popup._validate_redline_prerequisites("")
    self.assertFalse(is_valid)
    self.assertIn("Select text first", msg)

def test_validate_text_too_large(self):
    """Test validation fails with oversized text."""
    large_text = "x" * 100000
    is_valid, msg = self.popup._validate_redline_prerequisites(large_text)
    self.assertFalse(is_valid)
    self.assertIn("too large", msg)

def test_validate_success(self):
    """Test validation succeeds with normal text."""
    is_valid, msg = self.popup._validate_redline_prerequisites("normal text")
    self.assertTrue(is_valid)
    self.assertEqual(msg, "")
```

**Step 4: Run test**

Run: `python -m pytest tests/test_word_redline.py::TestWordRedline::test_validate_no_selection -v`
Expected: Test passes

**Step 5: Commit**

```bash
git add icharlotte_core/word_hotkey.py tests/test_word_redline.py
git commit -m "feat: validate selection before enabling redline mode"
```

---

## Task 9: Add Comprehensive Error Handling

**Files:**
- Modify: `icharlotte_core/word_hotkey.py` (_apply_redlines method)

**Step 1: Enhance error handling in _apply_redlines**

Update the method with more specific error handling:
```python
def _apply_redlines(self, word_app, selection, original_text: str, revised_text: str) -> bool:
    """Apply redlines using adeu RedlineEngine.

    Args:
        word_app: Word COM application object
        selection: Word Selection object
        original_text: Original text before LLM processing
        revised_text: LLM-revised text

    Returns:
        True if redlines applied successfully, False if fallback needed
    """
    try:
        from adeu import RedlineEngine

        # Get document object
        try:
            doc = word_app.ActiveDocument
        except Exception as e:
            print(f"Could not access Word document: {e}")
            return False

        # Auto-enable Track Changes if needed
        if self.redline_settings.get("auto_enable_track_changes", True):
            try:
                if not doc.TrackRevisions:
                    print("Auto-enabling Track Changes for redlining")
                    doc.TrackRevisions = True
            except Exception as e:
                print(f"Could not enable Track Changes: {e}")
                # Continue anyway - may still work

        # Reconstruct range from stored coordinates
        try:
            range_obj = doc.Range(
                self._original_range_start,
                self._original_range_end
            )
        except Exception as e:
            print(f"Could not reconstruct range: {e}, using current selection")
            try:
                range_obj = selection.Range
            except Exception as e2:
                print(f"Could not get selection range: {e2}")
                return False

        # Apply redlines using adeu
        try:
            engine = RedlineEngine()
            engine.apply_redlines(
                doc=doc,
                range_obj=range_obj,
                original=original_text,
                revised=revised_text
            )
        except Exception as e:
            print(f"RedlineEngine.apply_redlines failed: {e}")
            return False

        print(f"Successfully applied redlines ({len(original_text)} -> {len(revised_text)} chars)")
        return True

    except ImportError as e:
        print(f"adeu not available: {e}")
        return False
    except Exception as e:
        print(f"Unexpected error in _apply_redlines: {e}")
        import traceback
        traceback.print_exc()
        return False
```

**Step 2: Add error logging test**

Add to `tests/test_word_redline.py`:
```python
@patch('icharlotte_core.word_hotkey.RedlineEngine')
def test_apply_redlines_engine_failure(self, mock_engine_class):
    """Test graceful failure when engine throws exception."""
    # Setup mocks
    mock_word_app = Mock()
    mock_doc = Mock()
    mock_doc.TrackRevisions = True
    mock_doc.Range = Mock(return_value=Mock())
    mock_word_app.ActiveDocument = mock_doc

    mock_selection = Mock()

    # Make engine fail
    mock_engine = mock_engine_class.return_value
    mock_engine.apply_redlines = Mock(side_effect=Exception("Engine error"))

    self.popup._original_range_start = 0
    self.popup._original_range_end = 100

    # Apply redlines
    result = self.popup._apply_redlines(
        mock_word_app,
        mock_selection,
        "original",
        "revised"
    )

    # Should return False for fallback
    self.assertFalse(result)
```

**Step 3: Run tests**

Run: `python -m pytest tests/test_word_redline.py -v`
Expected: All tests pass

**Step 4: Commit**

```bash
git add icharlotte_core/word_hotkey.py tests/test_word_redline.py
git commit -m "feat: add comprehensive error handling for redline operations"
```

---

## Task 10: Integration Testing and Manual Verification

**Files:**
- Create: `tests/test_word_redline_integration.py`
- Read: `tests/test_word_hotkey_integration.py` (for patterns)

**Step 1: Create integration test file**

Create `tests/test_word_redline_integration.py`:
```python
"""Integration tests for Word redline functionality.

These tests require Word to be installed and may create temporary documents.
Run with: pytest tests/test_word_redline_integration.py --integration
"""

import unittest
import tempfile
import os

try:
    import win32com.client
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

@unittest.skipUnless(HAS_WIN32, "Requires win32com (Word) to be available")
class TestWordRedlineIntegration(unittest.TestCase):

    def setUp(self):
        """Create Word instance and temporary document."""
        self.word = win32com.client.Dispatch("Word.Application")
        self.word.Visible = False
        self.doc = self.word.Documents.Add()

    def tearDown(self):
        """Close document and Word."""
        try:
            self.doc.Close(SaveChanges=False)
            self.word.Quit()
        except:
            pass

    def test_track_changes_enabled(self):
        """Test that Track Changes can be enabled programmatically."""
        self.assertFalse(self.doc.TrackRevisions)
        self.doc.TrackRevisions = True
        self.assertTrue(self.doc.TrackRevisions)

    def test_range_reconstruction(self):
        """Test reconstructing a range from start/end coordinates."""
        # Insert some text
        self.doc.Content.Text = "Hello World"

        # Get range
        range_obj = self.doc.Range(0, 5)
        start = range_obj.Start
        end = range_obj.End

        # Reconstruct
        new_range = self.doc.Range(start, end)
        self.assertEqual(new_range.Text, "Hello")

if __name__ == '__main__':
    unittest.main()
```

**Step 2: Run integration tests**

Run: `python -m pytest tests/test_word_redline_integration.py -v --tb=short`
Expected: Tests pass (or skip if Word not available)

**Step 3: Manual testing checklist**

Create `docs/plans/manual-test-checklist.md`:
```markdown
# Word Redline Manual Testing Checklist

## Basic Functionality
- [ ] Open Word document
- [ ] Select text (e.g., a paragraph)
- [ ] Press Win+V to show popup
- [ ] Verify redline checkbox appears
- [ ] Check redline checkbox
- [ ] Enter prompt (e.g., "Make this more professional")
- [ ] Click Process
- [ ] Verify Track Changes appear in Word
- [ ] Accept/reject changes in Word UI

## Edge Cases
- [ ] No text selected -> checkbox disabled with tooltip
- [ ] Very small selection (1 word) -> works
- [ ] Large selection (1000+ words) -> works or shows warning
- [ ] Selection with formatting (bold, italic) -> formatting preserved
- [ ] Selection in table -> works or gracefully falls back

## Persistence
- [ ] Check redline checkbox -> close dialog -> reopen -> still checked
- [ ] Uncheck redline checkbox -> close dialog -> reopen -> still unchecked

## Error Handling
- [ ] Document with Track Changes OFF -> auto-enables
- [ ] adeu fails (simulate by renaming package) -> falls back to replace mode
- [ ] Shows appropriate error message on fallback

## Outlook Context
- [ ] Open Outlook compose window
- [ ] Press Win+V
- [ ] Verify redline checkbox is HIDDEN (not applicable to email)
```

**Step 4: Run manual tests**

Perform manual testing using the checklist.
Expected: All items pass.

**Step 5: Commit**

```bash
git add tests/test_word_redline_integration.py docs/plans/manual-test-checklist.md
git commit -m "test: add integration tests and manual test checklist"
```

---

## Task 11: Update Documentation

**Files:**
- Modify: `CLAUDE.md`
- Create: `docs/features/word-redline.md`

**Step 1: Create feature documentation**

Create `docs/features/word-redline.md`:
```markdown
# Word Redline Mode

## Overview

The Word AI Assistant now supports **Redline Mode**, which inserts AI suggestions as Track Changes instead of replacing text completely. This provides surgical editing capability crucial for legal document review.

## Usage

1. Select text in Word document
2. Press **Win+V** to open AI Assistant
3. Check **✏️ Use Redline Mode (Track Changes)**
4. Enter your prompt (e.g., "Make this more aggressive")
5. Click **Process**
6. AI suggestions appear as Track Changes in Word
7. Accept or reject changes using Word's review tools

## Configuration

Settings are stored in `~/.gemini/redline_settings.json`:

```json
{
  "redline_mode_default": false,
  "auto_enable_track_changes": true,
  "redline_fallback_notify": true,
  "max_redline_text_length": 50000
}
```

- `redline_mode_default`: Checkbox starts checked if true
- `auto_enable_track_changes`: Auto-enable Track Changes if document has it disabled
- `redline_fallback_notify`: Show notification when falling back to replace mode
- `max_redline_text_length`: Maximum characters for redline mode

## Technical Details

Uses [adeu](https://github.com/dealfluence/adeu) RedlineEngine to inject native Word Track Changes XML (`w:ins`, `w:del`) into documents while preserving formatting.

## Limitations

- Only available for Word (not Outlook emails)
- Requires text selection (cannot redline empty document)
- Very large selections (>50,000 chars) may be slow
- Complex document structures may fall back to replace mode
```

**Step 2: Update CLAUDE.md**

Add to "Recent Features" section in `CLAUDE.md`:
```markdown
### Word Redline Mode (2026-02-07)
- AI Assistant now supports Track Changes mode for surgical editing
- Uses adeu RedlineEngine for native Word redlining
- Configuration in `icharlotte_core/redline_config.py`
- Checkbox state persists across sessions
```

**Step 3: Commit**

```bash
git add docs/features/word-redline.md CLAUDE.md
git commit -m "docs: add Word redline mode documentation"
```

---

## Task 12: Final Testing and Cleanup

**Files:**
- All modified files

**Step 1: Run full test suite**

Run: `python -m pytest tests/ -v --tb=short`
Expected: All tests pass

**Step 2: Check for linting issues**

Run: `python -m flake8 icharlotte_core/word_hotkey.py icharlotte_core/redline_config.py --max-line-length=120`
Expected: No errors (or fix any found)

**Step 3: Verify imports and dependencies**

Run: `python -c "from icharlotte_core.word_hotkey import WordLLMPopup; from icharlotte_core.redline_config import load_redline_settings; print('OK')"`
Expected: Prints "OK"

**Step 4: Review git status**

Run: `git status`
Expected: Working tree clean (all changes committed)

**Step 5: Create summary commit**

```bash
git log --oneline feature/word-redline-integration --not master > /tmp/commits.txt
git commit --allow-empty -m "feat: Word redline integration complete

Summary of changes:
- Added adeu dependency for Word Track Changes integration
- Implemented redline configuration with persistence
- Added redline checkbox to Word AI popup
- Integrated RedlineEngine into LLM result handler
- Comprehensive error handling and fallback logic
- Unit and integration tests
- Documentation

$(cat /tmp/commits.txt)

Co-Authored-By: Claude Opus 4.6 <noreply@anthropic.com>"
```

**Step 6: Final verification**

Open Word, test the full workflow end-to-end.
Expected: Track Changes appear correctly in Word.

**Step 7: Done!**

Feature complete. Ready for merge or PR.

---

## Success Criteria

- [x] adeu dependency added
- [x] Redline configuration module with tests
- [x] Checkbox in Word AI popup (hidden for Outlook)
- [x] Checkbox state persists across sessions
- [x] Original text and range captured for redlining
- [x] RedlineEngine integrated into result handler
- [x] Auto-enable Track Changes if disabled
- [x] Validation for no selection / oversized text
- [x] Graceful fallback to replace mode on errors
- [x] Unit tests for all components
- [x] Integration tests for Word COM operations
- [x] Manual testing checklist completed
- [x] Documentation updated

---

## Notes

- All redline logic is isolated in `_apply_redlines()` method for maintainability
- Checkbox visibility controlled by `app_context` (Word vs Outlook)
- Settings persist to `redline_settings.json` separate from other configs
- Error handling prioritizes graceful degradation over failures
- Tests use mocking to avoid requiring Word installation in CI
