# Word/Outlook AI Assistant — File Attachments as Context

**Date**: 2026-02-09
**Status**: Approved

## Goal

Add the ability to attach external files (`.doc`, `.docx`, `.pdf`, `.msg`) as additional context in the Word/Outlook AI Assistant popup. Users can drag-and-drop or browse for files, and their extracted text is sent alongside the prompt to the LLM.

## UI Design

### Attachment Area (below prompt text area)

- **Default state**: Dashed-border drop zone, ~40px tall, with text "Drop files here or click to browse" and a paperclip (📎) button.
- **With files**: Each file appears as a removable chip showing `filename (size) ×`. Chips wrap horizontally. Area grows to fit.
- **Drag and drop**: Entire popup accepts file drops. Drop zone highlights on drag-over. Filters to accepted extensions only.
- **Browse button**: Paperclip icon button opens `QFileDialog` filtered to `*.doc;*.docx;*.pdf;*.msg`. Multi-select enabled.
- **Multiple files**: Unlimited files. Total context size label shows "3 files, ~12K chars".
- **Remove**: Click × on any chip to remove it and its extracted text.

## Text Extraction

Each file type routes to a specific extractor, running in a background `QThread`:

| Extension | Extractor | Library |
|-----------|-----------|---------|
| `.docx` | Paragraph text extraction | `python-docx` (existing) |
| `.doc` | COM automation text extraction | `win32com` (existing) |
| `.pdf` | Text extraction with OCR fallback | `pypdf` / Tesseract (existing `DocumentProcessor`) |
| `.msg` | Subject + sender + date + body | `extract-msg` (new dependency) |

- Extraction happens immediately when file is added, in a `QThread`.
- Chip shows spinner during extraction, then character count when done.
- Errors show as red chip text (e.g., "Failed to extract").

## Context Assembly

When the user clicks Execute, attached file texts are appended to the prompt after any document context:

```
{user's prompt}

=== FULL DOCUMENT (for context) ===          ← if "use all text" checked
{active document text}

=== SELECTED TEXT TO PROCESS ===
{selected text from Word/Outlook}

=== ATTACHED FILE: Complaint.pdf ===
{extracted text}

=== ATTACHED FILE: Smith_Deposition.docx ===
{extracted text}
```

Each file gets a clearly labeled block. The system prompt and LLM call remain unchanged.

## Implementation Changes

All changes are in `icharlotte_core/word_hotkey.py`:

### 1. `AttachmentArea` Widget (~50 lines)
- QWidget with QFlowLayout for file chips
- Drop zone styling (dashed border, highlight on drag-over)
- `dragEnterEvent` / `dropEvent` handlers — filter accepted extensions
- Paperclip browse button → `QFileDialog`
- Total context size label

### 2. `FileChipWidget` (~30 lines)
- QFrame showing filename, size/char count, spinner, remove button
- States: extracting (spinner), ready (char count), error (red text)

### 3. `FileExtractorThread(QThread)` (~60 lines)
- Accepts file path, routes to appropriate extractor
- `.docx` → `python-docx` paragraph extraction
- `.doc` → `win32com.client` Word COM
- `.pdf` → `DocumentProcessor` text extraction
- `.msg` → `extract_msg.Message` subject + body
- Emits `finished(path, text)` or `error(path, message)` signal

### 4. Prompt Assembly Modification (~20 lines)
- In the Execute flow (around lines 1588-1621), after building the existing context blocks, iterate `attachment_area.get_attachments()` and append each as `=== ATTACHED FILE: {name} ===\n{text}`

### 5. Popup Integration (~20 lines)
- Insert `AttachmentArea` in the AI Prompt tab layout, below the prompt QTextEdit
- Wire up signals: chip added → start extraction, chip removed → discard text
- Pass attachments dict to prompt assembly

### Not Changed
- `llm.py`, `LLMHandler`, system prompts — no changes needed
- Redline engine — works on selected text only, unaffected
- Format handling — unaffected
- Outlook flow — same attachment area works for both Word and Outlook contexts

## New Dependency

- `extract-msg` — lightweight .msg file parser (pip install extract-msg)

## Estimated Scope

~180 lines of new code, all within `word_hotkey.py`. No new files.
