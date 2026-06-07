# Med Record Extractor Chronology Viewer - Design

**Date:** 2026-06-06
**Status:** Approved design, pending implementation plan
**Scope:** Replace the pasted-text Med Record Extractor workflow with a structured chronology-document viewer that lets the user select Brief Synopsis paragraphs and/or chronology table rows, then extracts the corresponding source record pages.

## Goal

Make Med Record Extractor work from the medical chronology summary document itself.

The user should select a chronology `.docx`, review its Brief Synopsis paragraphs and medical chronology table in a structured viewer, check the paragraphs and rows that matter, and run extraction from the visible selected row set. The extracted PDFs and index document should continue to be saved in the same output location and format used today.

## Current State

- Wizard task `med_record_extractor` is registered as an in-process task with `script_name=""`.
- `task_routing.py` routes `med_record_extractor` to `build_med_extractor_tab`.
- `build_med_extractor_tab` currently shows `MedExtractorSettingsPage`, which accepts pasted prose text.
- `MedRecordExtractorWorker` currently asks an LLM to parse pasted prose into date/provider entries, finds the most recent chronology automatically, parses the first 5-column table, matches entries to rows, extracts source PDF pages, OCRs sparse output PDFs, and updates `NOTES/AI OUTPUT/Med Record Extracts/Index of Extracted Records.docx`.
- The Hall example chronology contains a heading `BRIEF SYNOPSIS OF POST-INJURY MEDICAL RECORD:` followed by dated prose paragraphs, and a 5-column chronology table with headers `DATE`, `PAGE NO`, `PROVIDER`, `DESCRIPTION`, and `Red Flags/Comments`.

## Approved Decisions

- Launching Med Record Extractor opens a `.docx` file picker first.
- The file picker defaults to `RECORDS/Medical Summary - DO NOT PRODUCE` when that folder exists.
- Use a structured viewer rather than a page-faithful Word overlay.
- Keep a hybrid reference affordance: the structured viewer is the working surface, and the user can open or preview the original `.docx` for orientation.
- Brief Synopsis paragraphs and chronology rows use explicit checkboxes plus selected-state highlighting.
- Selecting a Brief Synopsis paragraph infers matching chronology rows by date and treater/provider name.
- If a synopsis paragraph maps confidently to chronology rows, the matching rows become visibly selected automatically.
- Ask the user for confirmation only when the synopsis-to-row match is ambiguous.
- Keep extraction outputs unchanged: PDFs under `NOTES/AI OUTPUT/Med Record Extracts` plus `Index of Extracted Records.docx`.

## Non-Goals

- Do not build a fully page-faithful Word editor or annotation layer.
- Do not change the Medical Records Review agent (`medical_records` / `Scripts/med_record.py`).
- Do not change where extracted PDFs are saved.
- Do not require the user to paste chronology prose.
- Do not require user confirmation for confident synopsis matches.

## User Flow

1. User opens a case and selects **Med Record Extractor** from Wizard Mode.
2. The app opens a `.docx` picker, defaulting to the case's medical summary folder when available.
3. User selects the medical chronology summary document.
4. The app parses the document and opens a structured viewer.
5. The viewer shows:
   - selectable Brief Synopsis paragraphs
   - selectable medical chronology rows
   - selected count
   - extraction status and warnings
   - **Open Original** or equivalent reference action
   - **Extract Records**
6. User checks one or more Brief Synopsis paragraphs and/or chronology rows.
7. When a synopsis paragraph is checked, the app attempts date/provider matching against table rows.
8. Confident row matches are checked and highlighted automatically.
9. Ambiguous matches open a confirmation dialog with candidate rows.
10. Unmatched synopsis paragraphs stay checked with a warning badge but do not add extraction rows.
11. User clicks **Extract Records**.
12. The app extracts the selected row set and displays the same summary/output-folder page used today.

## Viewer Structure

### MedChronologySelectionPage

The main file-backed selection page.

Responsibilities:

- Store the selected chronology path.
- Own parsed synopsis paragraphs and chronology rows.
- Coordinate selection state.
- Show selected row count and warnings.
- Open the original `.docx` for reference.
- Start extraction with the final selected row set.

### BriefSynopsisPanel

Renders each synopsis paragraph as a selectable item.

Each paragraph item shows:

- checkbox
- paragraph text
- selected highlight
- matched-row indicator when rows were inferred
- warning badge for unmatched or ambiguous state

Selecting a paragraph triggers the matching pipeline. Deselecting a paragraph removes row selections that were contributed only by that paragraph, without removing rows the user directly selected or rows contributed by another still-selected paragraph.

### ChronologyTablePanel

Renders the parsed chronology table in a compact row-selection view.

Each row item shows:

- checkbox
- selected highlight
- date
- parsed page range
- provider
- description preview
- parse warning if `PAGE NO` is malformed or incomplete

Rows selected directly by the user and rows selected by synopsis inference should be visually indistinguishable as selected rows, but row details may show a small source label such as `Selected manually` or `Matched from synopsis`.

### AmbiguousMatchDialog

Shows candidate chronology rows when a selected synopsis paragraph has multiple plausible matches.

The dialog should include:

- the selected synopsis paragraph
- candidate rows with date, provider, page range, and description preview
- checkboxes for one or more candidates
- confirm and cancel actions

Confirmed candidate rows become selected rows. Cancel keeps the paragraph selected with an unresolved warning and does not add rows.

## Parsing

Add a parser layer that turns a chronology `.docx` into viewer data.

The parser should:

- open the selected `.docx` with `python-docx`
- locate the Brief Synopsis section by a heading equivalent to `BRIEF SYNOPSIS OF POST-INJURY MEDICAL RECORD`
- collect non-empty paragraphs after that heading until the chronology table or the next major section boundary
- locate the first 5-column table whose header row matches the chronology table headers
- skip table header rows
- skip merged section-header rows where all cells are identical, such as `POST-INJURY MEDICAL RECORDS`
- parse `PAGE NO` cells using the existing page parser behavior
- assign stable row IDs and paragraph IDs for selection state

The parser should not infer source PDF paths. It should only parse document contents and page references. Source PDF lookup remains part of extraction.

## Data Model

Use small, explicit data objects so parser, matcher, UI, and worker can be tested independently.

Suggested shapes:

- `ChronologyDocument`: source path, synopsis paragraphs, chronology rows, parse warnings
- `SynopsisParagraph`: id, text, order, matched row IDs, warning state
- `SelectableChronologyRow`: id, date, page_no, provider, description, flags, record filename, page start, page end, warnings
- `SelectionState`: selected paragraph IDs, selected row IDs, row selection sources
- `MatchResult`: status `confident`, `ambiguous`, or `none`; row IDs; candidate row IDs; reason

The implementation may reuse or extend the existing `MedChronRow` dataclass as long as UI row identity and warning state stay explicit.

## Synopsis Matching

Brief Synopsis matching should prefer deterministic parsing first.

Pipeline:

1. Extract candidate dates from the selected paragraph.
2. Extract treater/provider name candidates using deterministic patterns.
3. Compare candidates against chronology rows with the same normalized date.
4. Score provider similarity against the chronology row's provider cell.
5. Return a confident match when one row is clearly strongest.
6. Return ambiguous candidates when multiple rows are plausible or the confidence gap is too small.
7. Return no match when there is no plausible same-date/provider row.

LLM fallback is allowed for provider extraction only when deterministic extraction is insufficient. It should not be required for direct row selection or straightforward synopsis paragraphs.

## Extraction

Revise the worker boundary so extraction accepts selected chronology rows directly.

The extraction worker should:

- accept case path, file number, selected chronology path, and selected row payloads
- build the same source file index from `RECORDS` and `DISCOVERY/RESPONSES`
- look up each selected row's record filename
- extract each selected page range to `NOTES/AI OUTPUT/Med Record Extracts`
- OCR sparse extracted PDFs when possible
- continue processing after row-level failures
- update the cumulative index document with matched rows
- report successes, failures, and warnings in the output summary

Existing helper behavior should be reused where possible:

- page reference parsing
- source file index building
- fuzzy source-file lookup
- PDF page extraction
- OCR pass for sparse extracted PDFs
- index document update

Because this path produces or updates a `.docx`, the index update must comply with the project Word-validation rule by validating the updated index document after save.

## Error Handling

The viewer should fail early and visibly when the selected document cannot support extraction.

Cases:

- No Brief Synopsis section: warn, but allow direct table-row selection.
- No usable 5-column chronology table: block extraction and explain that no extractable chronology rows were found.
- Malformed `PAGE NO`: keep the row visible, mark it unextractable, and show the parse problem.
- Source PDF not found: record a row-level failure and continue other selected rows.
- OCR failure: keep the extracted PDF and warn that OCR failed.
- Ambiguous synopsis match: require user confirmation before that paragraph contributes rows.
- Unmatched synopsis paragraph: keep paragraph selected with warning, but do not extract from it.
- No selected extractable rows: disable **Extract Records** and explain why.

## Persistence And Restore

The implementation should preserve the current Wizard completion behavior where practical.

Persist enough state to restore an open extractor tab:

- selected chronology path, preferably case-relative when under the case folder
- selected paragraph IDs
- selected row IDs
- unresolved warning state if cheap to reconstruct

On restore, re-parse the chronology document from disk and reapply selections. If the document is missing or has changed so IDs no longer match, show a clear restore warning and let the user choose a chronology again.

## Testing

Parser tests:

- detects `BRIEF SYNOPSIS OF POST-INJURY MEDICAL RECORD`
- extracts synopsis paragraphs from the Hall-style structure
- finds the 5-column chronology table
- skips merged section-header rows
- parses supported page reference formats
- marks malformed page references as warnings
- produces stable row IDs for unchanged documents

Matching tests:

- selected synopsis paragraph matches row by date and provider
- same-date multiple-row cases use provider similarity
- ambiguous cases return candidate rows instead of auto-selecting
- unmatched paragraphs do not add extraction rows
- inferred rows become selected rows in selection state
- deselecting a paragraph removes only row selections solely contributed by that paragraph

UI and worker tests:

- Med Record Extractor opens the file picker instead of the pasted-text page
- selected chronology document opens the structured viewer
- paragraph and row checkboxes update selected highlights and selected count
- inferred rows are visibly selected automatically
- ambiguous-match dialog receives candidate rows and applies confirmed selection
- extraction worker receives selected row payloads rather than pasted prose
- output folder and summary behavior remain compatible with current output page
- index document validation is called after index save

Focused verification command:

```powershell
python -m pytest tests\test_med_record_extractor.py tests\test_wizard\test_med_record_extractor_viewer.py -q
```

Broader verification if routing or shared wizard containers change:

```powershell
python -m pytest tests\test_wizard -q
```

## Implementation Boundary

This design is ready for a separate implementation plan. The implementation plan should break the work into parser, matcher, UI, worker-boundary, and verification tasks.
