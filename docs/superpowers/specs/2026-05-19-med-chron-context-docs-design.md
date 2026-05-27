# Med Chron Custom Analyses — Per-Row Context Documents

**Status:** Approved (2026-05-19)
**Scope:** Wizard mode → Med Chron Analysis task → custom analyses

## Problem

A user can create custom Med-Chron analyses (e.g., "Identify treatment providers worth deposing for the defense") whose **instruction text persists globally** via `config/med_chron_custom_analyses.json`. There is no way to attach supplementary context — for example, a case status report whose narrative is needed to evaluate which providers matter to the defense.

The user wants:

1. Per-row context documents (`.pdf`, `.docx`, `.txt`) attached to a single custom analysis.
2. Two ways to attach: drag-and-drop onto the instruction textbox, or an "Add context" button.
3. The **instruction text persists across runs/cases**; the **attached files do NOT persist** (no leakage between cases).
4. On attach, warn if a PDF lacks an extractable text layer (likely needs OCR at run time).

## Approach

**Per-row chip strip.** Each `CustomAnalysisRow` grows a small horizontal strip below the instruction `QPlainTextEdit` containing:

- A chip per attached file: 📎 `filename.pdf` ✕
- A trailing `+ Add context` button → `QFileDialog.getOpenFileNames`
- Below the strip, an inline warning label (hidden when empty) listing any files that failed the on-attach text-layer sniff.

Drag-and-drop is handled by overriding `dragEnterEvent` / `dropEvent` on the instruction `QPlainTextEdit`: file URLs whose extension is `.pdf` / `.docx` / `.txt` are routed to the row's attach logic; any other mime data (plain text, etc.) falls through to the default behavior so normal text editing keeps working.

Extraction of the context documents happens at **Phase 2** (`Scripts/med_chron.py:process_run`) on the worker thread, just before the LLM call. The existing `_extract_full_text` helper covers all three file types. Extracted text is concatenated into a clearly-labeled block and injected into a new `{context_block}` placeholder in `_custom_wrapper.txt`.

## Data Model

### Session JSON

`user_config["custom_analyses"][i]` adds **one optional field**:

```json
{
  "label": "Defense deposition targets",
  "instruction": "Identify treatment providers worth deposing for the defense, given the case status.",
  "context_files": [
    "C:\\cases\\1234.567\\NOTES\\AI OUTPUT\\status_report.pdf",
    "C:\\cases\\1234.567\\NOTES\\IME report.docx"
  ]
}
```

Phase 2 reads it with `entry.get("context_files", [])` — older sessions without the key keep working.

### Global persistence (unchanged shape)

`config/med_chron_custom_analyses.json` continues to store **only `{label, instruction}`** for each entry. `context_files` is never written there — that's how we satisfy "files do not persist between sessions / cases."

### Form-internal state

`CustomAnalysisRow` gains a list `self._context_files: list[str]`. When a saved analysis loads from `custom_analyses_store`, this list starts empty.

## UI Changes

### `icharlotte_core/ui/med_chron_config_form.py`

**`CustomAnalysisRow`:**

1. Replace the plain `QPlainTextEdit` with a small subclass (`ContextDropTextEdit`, defined in the same module) that:
   - Sets `setAcceptDrops(True)`
   - Overrides `dragEnterEvent` / `dragMoveEvent`: if `mimeData.hasUrls()` and **every** URL is a local file with extension in `{.pdf, .docx, .txt}`, `event.acceptProposedAction()`. Else call `super()`.
   - Overrides `dropEvent`: same gate; on file drop emit a `files_dropped(list[str])` signal and `event.acceptProposedAction()`. Else `super().dropEvent(event)`.
2. Below the instruction textbox, add a `QHBoxLayout` chip strip:
   - One `QFrame` per file, styled as a chip: small paperclip icon (using a unicode 📎 or Qt standard icon) + filename + `✕` button.
   - Tooltip on the chip shows the absolute path.
   - Clicking `✕` removes the file from `_context_files` and re-renders the strip.
   - A trailing `QPushButton("+ Add context")` opens `QFileDialog.getOpenFileNames` with filter `Context documents (*.pdf *.docx *.txt)`; the start directory is `case_root` (passed in from the page) when available.
3. Below the chip strip, add a hidden `QLabel` (`self._context_warning_label`) styled in yellow/orange. Populated by `_refresh_context_warning()` after every attach.
4. New methods:
   - `add_context_files(paths: list[str])` — dedup against `_context_files`, append, re-render, run sniff.
   - `_remove_context_file(path: str)` — remove, re-render, re-sniff.
   - `_render_chip_strip()` — clear + re-add all chips.
   - `_refresh_context_warning()` — runs `sniff_text_layer(path)` for each file and populates the warning label.
   - `context_files() -> list[str]` — accessor used by the form's validator.
5. The existing `is_empty()` should still return True iff both label and instruction are empty — **context files alone do not count as content** (the user might attach a file before typing).

**`MedChronConfigForm`:**

- `_validated_custom_rows()` returns two parallel lists:
  - `persisted_rows: list[dict]` with `{label, instruction}` only (for `custom_analyses_store.save`)
  - `run_rows: list[dict]` with `{label, instruction, context_files}` (for the session JSON)
- `commit_user_config()`: pass `persisted_rows` to `custom_analyses_store.save`; pass `run_rows` to `session_manager.update_user_config`.

### Text-layer sniff helper

New module-level helper in `med_chron_config_form.py`:

```python
def sniff_text_layer(path: str) -> tuple[bool, str]:
    """Return (has_text, reason). has_text=False means the user should be warned."""
```

- `.txt`: read first 4 KB; `has_text = bool(content.strip())`.
- `.docx`: open with `python-docx`; sum text length across `doc.paragraphs[:50]`; `has_text = total > 50`.
- `.pdf`: open with `pypdf`; concatenate `page.extract_text()` from pages 0..min(2, n_pages-1); `has_text = len(text.strip()) > 200`.
- Any extraction exception → `(False, "could not read file")`.

Warning label format when one or more files fail the sniff:

```
⚠ Likely needs OCR at run time: status_report.pdf, scanned.pdf
```

This is a **warning, not a block** — the run proceeds.

## Phase 2 Changes (`Scripts/med_chron.py`)

### `_build_run_list`

For each entry in `cfg.get("custom_analyses", [])`:

1. Read `context_files = c.get("context_files", []) or []`.
2. If non-empty, build `context_block` by extracting each file via `_extract_full_text(path)` and concatenating:

   ```
   --- BEGIN CONTEXT DOCUMENT: status_report.pdf ---
   <extracted text, truncated to N chars per file>
   --- END CONTEXT DOCUMENT ---

   --- BEGIN CONTEXT DOCUMENT: IME report.docx ---
   ...
   ```

   Wrap the whole concatenation with a header:

   ```
   ADDITIONAL CONTEXT DOCUMENTS PROVIDED BY THE USER:

   <chunks here>
   ```

3. On extraction failure for any single file (exception, empty result, file missing): log a warning, skip that file, continue with the others. **Do not fail the whole analysis.**
4. If `context_files` was set but all files failed extraction: log a warning and proceed with an empty `context_block`.
5. Replace BOTH placeholders in the wrapper template:
   - `{user_instruction}` → `c["instruction"]`
   - `{context_block}` → the rendered block (empty string if no files / all failed).

### `_custom_wrapper.txt`

Updated template:

```
You will be given the BRIEF SYNOPSIS sections of a medical chronology PLUS
the underlying tables of medical entries.

{context_block}

The user has asked you to perform the following analysis on this
chronology:

{user_instruction}

Ground every finding in the document. Cite specific dates, providers, and
entries where applicable. When context documents are provided above, you
may reference them but the medical chronology is the primary source. Use
Markdown. Be specific.
```

When `context_block` is empty, the resulting prompt has a blank line where the block would be — acceptable; the LLM ignores it.

### Per-file size cap

To avoid runaway prompts, cap each extracted context file to `MAX_CONTEXT_CHARS = 120_000` characters (≈30k tokens). If exceeded, truncate and append:

```
[…context truncated at 120,000 characters…]
```

This is silent (logged but not surfaced in the UI). 120 KB of context per file is generous; the user will rarely hit it with a status report.

## Testing

### Unit tests (new file `tests/test_wizard/test_med_chron_context_docs.py`)

- `test_sniff_text_layer_txt_has_text` — non-empty `.txt` returns `(True, ...)`.
- `test_sniff_text_layer_txt_empty` — empty `.txt` returns `(False, ...)`.
- `test_sniff_text_layer_docx_has_text` — `.docx` with paragraphs returns `(True, ...)`.
- `test_sniff_text_layer_pdf_with_text` — `.pdf` with text layer returns `(True, ...)`.
- `test_sniff_text_layer_pdf_image_only` — `.pdf` with no extractable text returns `(False, ...)`.
- `test_drop_event_accepts_pdf_docx_txt_urls` — `ContextDropTextEdit` accepts file URLs of those types and rejects others.
- `test_drop_event_falls_through_for_plain_text` — non-file mimeData is passed to the base class (so text editing works).
- `test_add_context_files_dedupes` — adding the same file twice yields one chip.
- `test_remove_context_file_clears_chip` — clicking the X removes the file from `context_files()`.
- `test_validated_custom_rows_separates_persisted_and_run_shape` — persisted shape has no `context_files`; run shape includes it.
- `test_commit_user_config_persists_only_label_and_instruction` — after commit, `custom_analyses_store.load()` returns entries with NO `context_files` key.
- `test_loaded_saved_analysis_starts_with_empty_context_files` — re-opening the form after a previous commit yields rows with empty `_context_files`.

### Integration test (new test in `tests/test_med_chron/`)

- `test_phase2_includes_context_block_in_prompt` — set up a session JSON whose `custom_analyses` entry references a small synthetic `.txt` context file; run `process_run` with a mocked `LLMCaller.call`; assert the prompt passed to the LLM contains both `--- BEGIN CONTEXT DOCUMENT:` and the file's contents.
- `test_phase2_skips_missing_context_file_without_failing` — a `context_files` path that doesn't exist on disk is skipped, the analysis still runs against the remaining files (or with an empty context block), and the LLM is still called.
- `test_phase2_truncates_oversized_context_file` — a context file > `MAX_CONTEXT_CHARS` gets truncated and the truncation marker appears in the prompt.

### Manual smoke test

Per the global CLAUDE.md "Always test after developing or changing a feature" rule:

1. Open the iCharlotte app, switch to a case, run the wizard for `med_chron_analysis`.
2. On the settings page, click `+ Add custom analysis`, type a label + instruction.
3. Drag a small `.pdf` and a `.docx` onto the instruction textbox — chips should appear.
4. Click the `+ Add context` button — file dialog opens. Pick a `.txt`.
5. Attempt to drop an unsupported file (e.g., `.png`) — should NOT appear as a chip (falls through to default text-edit drop).
6. Attach an image-only PDF — warning label below the chip strip should mention it.
7. Click ✕ on a chip — it disappears.
8. Click Proceed → confirm Phase 2 runs and the produced `.docx` output reflects awareness of the context (or at least that no error is raised).
9. Reopen the case / restart the app and open the same wizard task — label + instruction should pre-populate from the saved analysis, but `context_files` should be empty.

## Files Touched

**New code paths:**

- `icharlotte_core/ui/med_chron_config_form.py` — `ContextDropTextEdit`, chip strip, sniff helper, dual-shape validator.
- `Scripts/med_chron.py` — context-block rendering in `_build_run_list`, truncation cap.
- `Scripts/MED_CHRON_ANALYSES/prompts/_custom_wrapper.txt` — new `{context_block}` placeholder.

**Tests:**

- `tests/test_wizard/test_med_chron_context_docs.py` (new)
- `tests/test_med_chron/test_context_block_rendering.py` (new) — or add to an existing test file in `tests/test_med_chron/` if one exists.

**Unchanged but verified compatible:**

- `icharlotte_core/med_chron/session_manager.py` — already passes `user_config` through without inspecting its shape.
- `icharlotte_core/med_chron/custom_analyses_store.py` — already discards unknown keys; no change needed.

## Out of Scope

- Persisting context files across sessions/cases — explicitly NOT wanted.
- Applying context to curated (non-custom) analyses — only custom analyses have user-defined prompts where context is meaningful.
- Re-ordering context files within a row — order has no semantic meaning.
- Token-usage display in the UI — not requested; the truncation cap is the only safeguard.
- OCR'ing image-only PDFs at attach time — would block the UI; we only warn.
