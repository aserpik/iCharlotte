# .MSG File Support Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Enable .msg (Outlook email) files to be summarized by the summarize agent and loaded as context in the Chat, Email Update, and Liability/Exposure tabs.

**Architecture:** Add a single `extract_text_from_msg()` helper using the existing `win32com` + `pythoncom` Outlook COM approach (already proven in `Scripts/extract_msg_content.py`). Wire it into three extraction points: `DocumentProcessor.extract_text()`, `utils.extract_text_from_file()`, and `ChatTab.read_files_content()`. Update file extension filters across all UI tabs.

**Tech Stack:** `pywin32` (already installed), `pythoncom` for COM threading, existing `DocumentProcessor` and `utils` patterns.

---

### Task 1: Add `_extract_from_msg()` to DocumentProcessor

This is the core extraction method. All agent scripts (summarize, etc.) use `DocumentProcessor` for text extraction.

**Files:**
- Modify: `icharlotte_core/document_processor.py:125-163` (add .msg branch in `extract_text()`)
- Modify: `icharlotte_core/document_processor.py:392-404` (add page count support)
- Test: `tests/test_msg_extraction.py`

**Step 1: Write the failing test**

```python
# tests/test_msg_extraction.py
"""Tests for .msg file extraction via DocumentProcessor."""
import os
import unittest
from unittest.mock import patch, MagicMock
from icharlotte_core.document_processor import DocumentProcessor, ExtractResult, ExtractionMethod


class TestMsgExtraction(unittest.TestCase):
    """Test .msg file extraction in DocumentProcessor."""

    @patch('icharlotte_core.document_processor.pythoncom', create=True)
    @patch('icharlotte_core.document_processor.win32com_client', create=True)
    def test_extract_msg_returns_subject_and_body(self, mock_win32, mock_pythoncom):
        """Extract text from .msg should return subject + body."""
        # Mock the Outlook COM chain
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

    @patch('icharlotte_core.document_processor.pythoncom', create=True)
    @patch('icharlotte_core.document_processor.win32com_client', create=True)
    def test_extract_msg_com_error_returns_failed(self, mock_win32, mock_pythoncom):
        """If Outlook COM fails, return a failed result, not an exception."""
        mock_win32.Dispatch.side_effect = Exception("Outlook not available")

        processor = DocumentProcessor()
        result = processor.extract_text("C:\\fake\\test.msg")

        assert not result.success
        assert "Outlook" in result.error or "error" in result.error.lower()

    @patch('icharlotte_core.document_processor.pythoncom', create=True)
    @patch('icharlotte_core.document_processor.win32com_client', create=True)
    def test_extract_msg_html_fallback(self, mock_win32, mock_pythoncom):
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
```

**Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_msg_extraction.py -v`
Expected: FAIL — `_extract_from_msg` doesn't exist yet, `.msg` falls through to `_extract_from_text` which tries to read raw binary.

**Step 3: Implement `_extract_from_msg()` and wire into `extract_text()`**

In `icharlotte_core/document_processor.py`:

1. Add imports at the top (after existing imports, ~line 40):
```python
# Outlook .msg extraction
try:
    import pythoncom
    import win32com.client as win32com_client
    MSG_AVAILABLE = True
except ImportError:
    MSG_AVAILABLE = False
    pythoncom = None
    win32com_client = None
```

2. Add the `.msg` branch in `extract_text()` (line 146-153, before the `else` fallthrough):
```python
        elif ext == ".msg":
            return self._extract_from_msg(file_path)
```

3. Add the method (after `_extract_from_text`, ~line 390):
```python
    def _extract_from_msg(self, file_path: str) -> ExtractResult:
        """Extract text from an Outlook .msg file using COM automation."""
        if not MSG_AVAILABLE:
            return ExtractResult(
                text="", page_count=0,
                extraction_method=ExtractionMethod.FAILED,
                file_path=file_path,
                error="pywin32 not installed — cannot read .msg files"
            )

        try:
            pythoncom.CoInitialize()
            outlook = win32com_client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            abs_path = os.path.abspath(file_path)
            item = namespace.OpenSharedItem(abs_path)

            subject = item.Subject or ""
            sender = item.SenderName or ""
            sent_on = str(item.SentOn) if item.SentOn else ""
            body = item.Body or ""

            # Fallback: strip HTML tags if plain body is empty
            if not body.strip() and item.HTMLBody:
                import re
                body = re.sub(r'<[^>]+>', '', item.HTMLBody)
                body = body.strip()

            item.Close(0)  # 0 = olDiscard

            full_text = f"From: {sender}\nDate: {sent_on}\nSubject: {subject}\n\n{body}"

            return ExtractResult(
                text=full_text,
                page_count=1,
                extraction_method=ExtractionMethod.NATIVE,
                char_count=len(full_text),
                char_density=len(full_text),
                file_path=file_path
            )
        except Exception as e:
            self._log(f"Error extracting .msg file: {e}", "error")
            return ExtractResult(
                text="", page_count=0,
                extraction_method=ExtractionMethod.FAILED,
                file_path=file_path,
                error=str(e)
            )
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
```

**Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_msg_extraction.py -v`
Expected: All 3 tests PASS.

**Step 5: Commit**

```bash
git add icharlotte_core/document_processor.py tests/test_msg_extraction.py
git commit -m "feat: add .msg extraction to DocumentProcessor via Outlook COM"
```

---

### Task 2: Add .msg support to `utils.extract_text_from_file()`

The Email Update tab uses `utils.extract_text_from_file()` — a separate, simpler extraction function. Currently it treats `.msg` as plain text (line 400), producing garbage output from binary .msg data.

**Files:**
- Modify: `icharlotte_core/utils.py:373-412` (fix .msg branch)
- Test: `tests/test_msg_extraction.py` (add test)

**Step 1: Write the failing test**

Add to `tests/test_msg_extraction.py`:

```python
class TestUtilsMsgExtraction(unittest.TestCase):
    """Test .msg extraction in utils.extract_text_from_file."""

    @patch('icharlotte_core.utils.pythoncom', create=True)
    @patch('icharlotte_core.utils.win32com_client', create=True)
    def test_utils_extract_msg(self, mock_win32, mock_pythoncom):
        """utils.extract_text_from_file should use COM for .msg files."""
        mock_item = MagicMock()
        mock_item.Subject = "Utility Test"
        mock_item.Body = "Body from utils extraction."
        mock_item.SenderName = "Sender"
        mock_item.SentOn = "2026-02-01"

        mock_namespace = MagicMock()
        mock_namespace.OpenSharedItem.return_value = mock_item
        mock_outlook = MagicMock()
        mock_outlook.GetNamespace.return_value = mock_namespace
        mock_win32.Dispatch.return_value = mock_outlook

        # Need to patch os.path.exists to return True for fake path
        with patch('icharlotte_core.utils.os.path.exists', return_value=True):
            from icharlotte_core.utils import extract_text_from_file
            result = extract_text_from_file("C:\\fake\\test.msg")

        assert result is not None
        assert "Utility Test" in result
        assert "Body from utils extraction." in result
```

**Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_msg_extraction.py::TestUtilsMsgExtraction -v`
Expected: FAIL — current code reads .msg as raw text, won't contain structured subject/body.

**Step 3: Fix `utils.extract_text_from_file()` for .msg**

In `icharlotte_core/utils.py`:

1. Add imports near the top (after existing imports):
```python
try:
    import pythoncom
    import win32com.client as win32com_client
    MSG_AVAILABLE = True
except ImportError:
    MSG_AVAILABLE = False
    pythoncom = None
    win32com_client = None
```

2. Replace the `.msg` handling in `extract_text_from_file()`. Change line 400 from:
```python
        elif ext in ['.txt', '.md', '.py', '.json', '.xml', '.html', '.htm', '.msg']:
```
to separate `.msg` into its own branch BEFORE the text fallback:
```python
        elif ext == '.msg':
            if not MSG_AVAILABLE:
                return "[Error: pywin32 not installed — cannot read .msg files]"
            try:
                pythoncom.CoInitialize()
                outlook = win32com_client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
                item = namespace.OpenSharedItem(os.path.abspath(file_path))
                subject = item.Subject or ""
                sender = item.SenderName or ""
                sent_on = str(item.SentOn) if item.SentOn else ""
                body = item.Body or ""
                if not body.strip() and item.HTMLBody:
                    import re
                    body = re.sub(r'<[^>]+>', '', item.HTMLBody).strip()
                item.Close(0)
                return f"From: {sender}\nDate: {sent_on}\nSubject: {subject}\n\n{body}"
            except Exception as e:
                log_event(f"Error reading .msg file: {e}", "error")
                return f"[Error reading .msg file: {e}]"
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

        elif ext in ['.txt', '.md', '.py', '.json', '.xml', '.html', '.htm']:
```

**Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_msg_extraction.py -v`
Expected: All tests PASS.

**Step 5: Commit**

```bash
git add icharlotte_core/utils.py tests/test_msg_extraction.py
git commit -m "feat: add .msg COM extraction to utils.extract_text_from_file"
```

---

### Task 3: Add .msg support to ChatTab (read_files_content + UI filters)

The ChatTab handles drag-and-drop and file dialog. Its `read_files_content()` method has its own inline extraction (doesn't use DocumentProcessor or utils). Need to add `.msg` to both the extension filter and the content reader.

**Files:**
- Modify: `icharlotte_core/ui/tabs.py:720` (file dialog filter)
- Modify: `icharlotte_core/ui/tabs.py:930` (drop event supported_extensions)
- Modify: `icharlotte_core/ui/tabs.py:987-992` (read_files_content — add .msg elif)

**Step 1: Update the file dialog filter (line 720)**

Change:
```python
"All Supported (*.pdf *.docx *.txt *.png *.jpg *.jpeg *.gif *.webp);;Documents (*.pdf *.docx *.txt);;Images (*.png *.jpg *.jpeg *.gif *.webp)"
```
To:
```python
"All Supported (*.pdf *.docx *.txt *.msg *.png *.jpg *.jpeg *.gif *.webp);;Documents (*.pdf *.docx *.txt *.msg);;Images (*.png *.jpg *.jpeg *.gif *.webp)"
```

**Step 2: Update `supported_extensions` in `dropEvent` (line 930)**

Change:
```python
supported_extensions = (".pdf", ".docx", ".txt", ".png", ".jpg", ".jpeg", ".gif", ".webp")
```
To:
```python
supported_extensions = (".pdf", ".docx", ".txt", ".msg", ".png", ".jpg", ".jpeg", ".gif", ".webp")
```

**Step 3: Add .msg extraction in `read_files_content()` (after line 992)**

Add a new `elif` branch after the `.pdf` block and before the `except`:

```python
                    elif ext == ".msg":
                        try:
                            import pythoncom
                            import win32com.client
                            pythoncom.CoInitialize()
                            outlook = win32com.client.Dispatch("Outlook.Application")
                            namespace = outlook.GetNamespace("MAPI")
                            item = namespace.OpenSharedItem(os.path.abspath(path))
                            subject = item.Subject or ""
                            sender = item.SenderName or ""
                            body = item.Body or ""
                            if not body.strip() and item.HTMLBody:
                                import re
                                body = re.sub(r'<[^>]+>', '', item.HTMLBody).strip()
                            item.Close(0)
                            pythoncom.CoUninitialize()
                            content += f"From: {sender}\nSubject: {subject}\n\n{body}\n"
                        except Exception as msg_err:
                            content += f"[Error reading .msg file: {msg_err}]\n"
```

**Step 4: Verify manually**

1. Launch iCharlotte: `python iCharlotte.py`
2. Open the Chat tab
3. Drag-and-drop a `.msg` file — should appear in file list
4. Click "Select Files" — `.msg` should appear in file dialog filter
5. Send a message with the `.msg` file checked — email content should appear in context

**Step 5: Commit**

```bash
git add icharlotte_core/ui/tabs.py
git commit -m "feat: add .msg drag-drop and context loading to ChatTab"
```

---

### Task 4: Add .msg to the Summarize Agent file filter

The summarize agent scans directories for files to process. Currently it only picks up `.pdf` and `.docx`.

**Files:**
- Modify: `Scripts/summarize.py:515` (add .msg to extension tuple)

**Step 1: Update the extension filter (line 515)**

Change:
```python
if file.lower().endswith(('.pdf', '.docx')):
```
To:
```python
if file.lower().endswith(('.pdf', '.docx', '.msg')):
```

**Step 2: Verify manually**

Place a `.msg` file in a case folder and run:
```bash
python -m Scripts.summarize "path/to/test_email.msg"
```
Expected: The summarize agent extracts text via DocumentProcessor (Task 1), sends to LLM, produces summary.

**Step 3: Commit**

```bash
git add Scripts/summarize.py
git commit -m "feat: add .msg to summarize agent file filter"
```

---

### Task 5: Add .msg to config RESOURCE_EXTENSIONS

This ensures `.msg` files show up in file tree browsing and are recognized as valid resources across the app.

**Files:**
- Modify: `icharlotte_core/config.py:14`

**Step 1: Update RESOURCE_EXTENSIONS**

Change:
```python
RESOURCE_EXTENSIONS = ['.pdf', '.docx', '.doc', '.txt', '.html', '.png', '.jpg', '.jpeg']
```
To:
```python
RESOURCE_EXTENSIONS = ['.pdf', '.docx', '.doc', '.txt', '.msg', '.html', '.png', '.jpg', '.jpeg']
```

**Step 2: Commit**

```bash
git add icharlotte_core/config.py
git commit -m "feat: add .msg to RESOURCE_EXTENSIONS config"
```

---

### Task 6: Verify LiabilityExposureTab inherits .msg support

`LiabilityExposureTab` extends `ChatTab` (line 352 of `liability_tab.py`). After Task 3, it should inherit `.msg` support for free — both the drop event and `read_files_content()`. Verify this.

**Step 1: Read `liability_tab.py` to confirm no overrides**

Check that `LiabilityExposureTab` does NOT override `dropEvent`, `read_files_content`, or `select_files`. If it does override any of them, add `.msg` there too.

**Step 2: Verify manually**

1. Launch iCharlotte
2. Open the Liability/Exposure tab
3. Drag a `.msg` file — should be accepted
4. Run an analysis with the `.msg` file checked — content should be included

**Step 3: Commit (only if changes were needed)**

If overrides exist and needed updating:
```bash
git add icharlotte_core/ui/liability_tab.py
git commit -m "feat: add .msg support to LiabilityExposureTab overrides"
```

---

### Task 7: Final integration test

**Step 1: Run all tests**

```bash
python -m pytest tests/test_msg_extraction.py -v
python -m pytest tests/ -v --timeout=30
```

**Step 2: Manual end-to-end verification**

1. **Summarize agent**: Run `python -m Scripts.summarize "path/to/email.msg"` — should produce summary in AI_OUTPUT.docx
2. **Chat tab**: Drag `.msg`, send message, see email content in response
3. **Email Update tab**: Load `.msg` via file dialog, generate update — email content used
4. **Liability tab**: Drag `.msg`, run analysis — email content included

**Step 3: Final commit**

```bash
git add -A
git commit -m "feat: complete .msg file support across summarize agent and all UI tabs"
```

---

## Summary of All Changes

| File | Change |
|------|--------|
| `icharlotte_core/document_processor.py` | Add `_extract_from_msg()` + `.msg` branch in `extract_text()` |
| `icharlotte_core/utils.py` | Replace text fallback with COM extraction for `.msg` |
| `icharlotte_core/ui/tabs.py` | Add `.msg` to file dialog, drop filter, and `read_files_content()` |
| `icharlotte_core/config.py` | Add `.msg` to `RESOURCE_EXTENSIONS` |
| `Scripts/summarize.py` | Add `.msg` to directory scan filter |
| `tests/test_msg_extraction.py` | New: unit tests for both extraction paths |

**Not changed (inherits from ChatTab):**
- `icharlotte_core/ui/liability_tab.py` — `LiabilityExposureTab` extends `ChatTab`
- `icharlotte_core/ui/email_update_tab.py` — already has `.msg` in file dialog; uses `utils.extract_text_from_file()` which is fixed in Task 2

## Key Design Decisions

1. **Win32 COM over `extract-msg` library**: The codebase already uses `pywin32` extensively. COM gives us full fidelity (handles all .msg variants, embedded attachments, HTML). No new dependency.
2. **Three extraction points**: Unfortunately the codebase has 3 independent text extraction codepaths (`DocumentProcessor`, `utils`, `ChatTab inline`). All three need updating. A future refactor could consolidate these.
3. **HTML fallback**: Some emails are HTML-only (`Body` is empty). We strip tags as a fallback — good enough for LLM context. No need for a full HTML parser.
4. **`pythoncom.CoInitialize/CoUninitialize`**: Required for COM in threads. The summarize agent runs in `QThread` workers, so this is essential.
