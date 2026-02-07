# Word Redline Integration Design

**Date:** 2026-02-07
**Feature:** Integrate adeu RedlineEngine into Word AI Assistant
**Status:** Design Approved

---

## Overview

Enhance the existing Word AI Assistant (`word_hotkey.py`) to support **redline mode** where AI suggestions are inserted as Track Changes instead of replacing text completely. This provides surgical editing capability crucial for legal document review.

---

## 1. Architecture Overview

**Integration Point:** Extend the existing `WordLLMPopup` class in `word_hotkey.py` with redline capability via the adeu library.

**Key Components:**

1. **adeu RedlineEngine** - Python SDK that will be added as a dependency. The engine takes original text, revised text, and the Word document COM object to inject Track Changes.

2. **Modified LLM Flow** - The `LLMWorkerThread` remains unchanged (still generates improved text), but the result handling splits into two paths:
   - **Replace Mode** (existing): `selection.TypeText(revised_text)`
   - **Redline Mode** (new): `RedlineEngine.apply_changes(doc, selection.Range, original_text, revised_text)`

3. **Document Reference** - The adeu engine needs access to the actual Word document object (not just the selection). We already have this via `word.ActiveDocument` in the COM automation code.

4. **State Management** - Add a boolean flag `self.redline_mode` to track checkbox state, passed to the worker thread and result handler.

The design is **minimally invasive** - the LLM generation logic, prompt system, format handling, and UI framework all remain unchanged. We're only modifying the final step where results are written back to Word.

---

## 2. UI Changes

**Checkbox Placement:** Add a "Use Redline Mode" checkbox in the `WordLLMPopup` dialog, positioned between the prompt selection area and the Process/Cancel buttons.

**Visual Design:**
- Label: "✏️ Use Redline Mode (Track Changes)"
- Tooltip: "Instead of replacing text, insert AI suggestions as Track Changes that you can accept/reject in Word"
- Default state: **Unchecked** (preserves existing behavior)
- Only visible when `APP_CONTEXT_WORD` is active (hidden for Outlook emails)

**Implementation:**
```python
# In WordLLMPopup.__init__()
self.redline_checkbox = QCheckBox("✏️ Use Redline Mode (Track Changes)")
self.redline_checkbox.setToolTip("Insert AI suggestions as Track Changes")
self.redline_checkbox.setChecked(False)

# Add to layout between prompts and buttons
layout.addWidget(self.redline_checkbox)
```

**State Persistence:** Store the checkbox state in the user's settings JSON (same location as custom format settings) so it persists across sessions.

**Context Awareness:** Disable the checkbox with a warning tooltip if the active Word document has Track Changes already disabled in its settings.

---

## 3. Redline Processing Flow

**Flow Bifurcation:** When the Process button is clicked, the handler checks `self.redline_checkbox.isChecked()` and sets a flag on the `LLMWorkerThread`. After LLM generation completes, the result handler branches:

**Replace Mode Path (existing):**
```python
if not redline_mode:
    selection.TypeText(revised_text)
```

**Redline Mode Path (new):**
```python
if redline_mode:
    # 1. Capture original text and range
    original_text = selection.Text
    original_range = selection.Range

    # 2. Get document object
    doc = word.ActiveDocument

    # 3. Use adeu RedlineEngine
    from adeu import RedlineEngine
    engine = RedlineEngine()
    engine.apply_redlines(
        doc=doc,
        range_obj=original_range,
        original=original_text,
        revised=revised_text
    )
```

**Key Technical Details:**
- Preserve the `selection.Range` object before LLM processing (since COM references can become stale)
- Store `original_text` before sending to LLM (needed for adeu's diff calculation)
- adeu handles the complex XML manipulation - we just provide the COM objects and text
- If adeu fails (e.g., document structure too complex), fall back to replace mode with a warning notification

The LLM worker thread remains unchanged - it still just generates text. All redline logic happens in the result handler on the main thread where COM automation is safe.

---

## 4. Error Handling & Edge Cases

**Graceful Degradation:** If adeu's RedlineEngine encounters issues, automatically fall back to replace mode with user notification:

```python
try:
    engine.apply_redlines(doc, range_obj, original, revised)
    self.status_label.setText("✓ Redlines applied!")
except Exception as e:
    # Log the error
    print(f"Redline failed: {e}, falling back to replace mode")
    # Fall back to regular replacement
    selection.TypeText(revised_text)
    self.status_label.setText("⚠ Applied as replacement (redline unavailable)")
```

**Edge Cases:**

1. **No Selection:** If user hasn't selected text, disable the checkbox with tooltip: "Select text first to use redline mode"

2. **Track Changes Disabled:** If Word document has Track Changes globally disabled, **auto-enable** it before applying redlines via `doc.TrackRevisions = True`

3. **Complex Formatting:** If selection contains tables, images, or nested structures that adeu can't handle, gracefully fall back to replace mode

4. **Large Selections:** For very large text blocks (>10,000 words), show a progress indicator since diff calculation may take a few seconds

5. **COM Stale References:** Store original range coordinates (`range.Start`, `range.End`) in addition to the Range object, allowing reconstruction if the reference becomes invalid

---

## 5. Configuration & Dependencies

**Adding adeu Dependency:**

Add to `requirements.txt`:
```
adeu>=0.1.0  # Agentic DOCX Redlining Engine
```

Install via: `pip install adeu`

**Configuration Settings:**

Add to the existing settings JSON:

```json
{
  "word_hotkey": {
    "redline_mode_default": false,
    "auto_enable_track_changes": true,
    "redline_fallback_notify": true,
    "max_redline_text_length": 50000
  }
}
```

**Settings Meanings:**
- `redline_mode_default`: Whether checkbox starts checked
- `auto_enable_track_changes`: Auto-enable Track Changes if disabled (confirmed requirement)
- `redline_fallback_notify`: Show notification when falling back to replace mode
- `max_redline_text_length`: Character limit before showing performance warning

**Testing Approach:**

1. **Unit tests** in `tests/test_word_hotkey.py`:
   - Mock adeu RedlineEngine to verify correct parameters passed
   - Test checkbox state persistence
   - Test fallback logic when engine fails

2. **Integration tests** in `tests/test_word_hotkey_integration.py`:
   - Create test .docx file
   - Apply redlines via API
   - Verify Track Changes XML exists in document

3. **Manual testing checklist:**
   - Test with Track Changes on/off
   - Test with formatted text (bold, bullets)
   - Test with large selections
   - Verify redlines are accept/reject-able in Word

---

## Implementation Files to Modify

1. **`icharlotte_core/word_hotkey.py`**
   - Add redline checkbox to `WordLLMPopup.__init__()`
   - Add `self.redline_mode` flag
   - Modify result handler to branch on redline mode
   - Add `_apply_redlines()` method for adeu integration
   - Add Track Changes auto-enable logic

2. **`requirements.txt`**
   - Add `adeu>=0.1.0`

3. **Settings/Config file** (wherever word_hotkey settings are stored)
   - Add redline configuration section

4. **`tests/test_word_hotkey.py`**
   - Add unit tests for redline mode

5. **`tests/test_word_hotkey_integration.py`**
   - Add integration tests for Track Changes

---

## Success Criteria

- [ ] Checkbox appears in Word AI Assistant popup (not in Outlook mode)
- [ ] Checkbox state persists across sessions
- [ ] When checked, AI suggestions appear as Track Changes in Word
- [ ] When unchecked, existing replace behavior works unchanged
- [ ] Track Changes auto-enables if disabled
- [ ] Graceful fallback to replace mode on errors
- [ ] All tests pass
- [ ] Manual testing validates accept/reject workflow in Word

---

## Future Enhancements (Out of Scope)

- CriticMarkup preview before applying redlines
- Batch redline multiple sections
- Redline statistics (insertions/deletions count)
- Integration with other document formats (Google Docs)
