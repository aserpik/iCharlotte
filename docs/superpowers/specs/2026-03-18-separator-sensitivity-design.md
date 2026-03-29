# Separator Sensitivity Control — Design Spec

**Date:** 2026-03-18
**Status:** Approved

## Problem

The document separator agent (`Scripts/separate.py`) sometimes groups too aggressively, merging related-but-distinct documents into a single entry. Users need a way to control how fine-grained the separation is.

## Solution

Add a 3-position sensitivity toggle in the IndexTab UI that modifies the LLM prompt used for document boundary detection.

## Sensitivity Levels

| Level | Label | Value | Behavior |
|-------|-------|-------|----------|
| Broad | Broad | 1 | Groups aggressively — related documents merged into fewer entries |
| Default | Default | 2 | Current behavior (unchanged prompt) |
| Fine | Fine | 3 | Splits aggressively — each sub-document, exhibit, attachment gets its own entry |

## UI Changes (IndexTab in `icharlotte_core/ui/tabs.py`)

- 3-position horizontal `QSlider` (min=1, max=3) with labels "Broad" / "Default" / "Fine"
- Placed in a horizontal layout above the document table
- "Re-analyze" button next to the slider
- Visible whenever `self.current_pdf_path` is set (a PDF is selected in the list)
- Slider resets to Default (2) on app startup — no persistence across sessions
- Clicking "Re-analyze" re-runs the separator on the current PDF with the selected sensitivity
- Both the slider and button are disabled while re-analysis is running, re-enabled on completion
- Results replace the existing entries in the document table (no confirmation dialog — the user can always re-analyze again)

### Signal Wiring

IndexTab stores a reference to the main window's `run_separator_path` method (passed during construction or connected via signal). The "Re-analyze" button calls it directly with `(self.current_pdf_path, sensitivity)`.

## Prompt Changes (`Scripts/separate.py`)

### New CLI arg
- `--sensitivity {1,2,3}` via `argparse` with `choices=[1,2,3]`, defaults to 2

### Parameter threading
The sensitivity value flows through the full call chain:
- `main()` parses `--sensitivity` → passes to `run_analysis(pdf_path, headless, sensitivity)`
- `run_analysis()` passes to `analyze_headers(headers, sensitivity)`
- `analyze_headers()` passes to `analyze_headers_chunk(headers_subset, start_page_num, next_id, prev_doc_context, sensitivity)`

### Prompt modifications in `analyze_headers_chunk`

**Broad (1):**
- Keeps Rule 5 (group insurance policy parts)
- Adds: "Group related documents together liberally. For example, a motion and its exhibits should be ONE entry. A letter and its attachments should be ONE entry. Only create separate entries for clearly distinct, unrelated documents."

**Default (2):**
- Current prompt unchanged

**Fine (3):**
- Removes Rule 5 (insurance policy grouping)
- Adds: "Be aggressive about identifying separate documents. Each exhibit, attachment, declaration, addendum, or sub-document should be its own entry. When in doubt, split rather than group."

### Logging
Log the sensitivity level at the start of analysis: `logger.info(f"Sensitivity: {sensitivity}")`

## Wiring (`iCharlotte.py`)

1. `run_separator_path(path, sensitivity=2)` — new optional param
2. Passes `--sensitivity N` as additional CLI arg at line 1520: `args = [script_path, "--headless", "--sensitivity", str(sensitivity), path]`
3. Results return via `JSON_MAP:` stdout, replacing existing `index_data[path]`
4. Default sensitivity from the file tree checkbox flow remains 2

## Files Changed

1. `Scripts/separate.py` — new `--sensitivity` arg, parameter threading through call chain, conditional prompt rules
2. `icharlotte_core/ui/tabs.py` — slider + re-analyze button in IndexTab, disable/enable during run
3. `iCharlotte.py` — `run_separator_path` accepts and passes sensitivity param

## Out of Scope

- No changes to JSON format, index Word doc, or interactive mode
- No changes to the file tree checkbox launch flow (uses default sensitivity)
- No migration of `separate.py` from direct Gemini calls to `LLMCaller`
