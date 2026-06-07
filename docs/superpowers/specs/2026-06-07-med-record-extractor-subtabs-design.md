# Med Record Extractor Subtabs Design

## Goal

Update the Med Record Extractor chronology viewer so users can review and select chronology content in two focused subtabs:

- Brief Synopsis
- Chronology Rows

Both views must preserve the source chronology text verbatim. The viewer should make selection faster by allowing a single click anywhere on an entry or row to toggle selection, while keeping visible checkboxes as the selected-state indicator.

## Current Context

The existing Med Record Extractor viewer loads a selected medical chronology `.docx`, parses Brief Synopsis paragraphs and chronology table rows, and displays them side-by-side. Selecting a Brief Synopsis paragraph attempts date/provider matching and auto-selects confident chronology rows. Selecting chronology table rows directly extracts those source record pages.

The current parser collapses whitespace for synopsis text and some row display fields. This is acceptable for matching, but no longer acceptable for display. The updated viewer needs display text that is verbatim to the original document while still allowing normalized text internally for matching and identifiers.

## User Experience

The settings page will use a `QTabWidget` with two tabs.

### Brief Synopsis Tab

Each Brief Synopsis entry appears as a full-width, checkable item. The full entry text must be visible and verbatim to the original `.docx` paragraph text. No display-time rewriting, summarizing, punctuation changes, or whitespace normalization is allowed.

The user can select or unselect an entry in either of two ways:

- click the checkbox
- single-click anywhere on the entry text

Clicking a selected entry again unselects it. Selected entries remain visually obvious through the checkbox and selected-row styling. Selecting a synopsis entry continues to run the existing confident date/provider matching logic. Ambiguous synopsis matches continue to require manual chronology row selection.

### Chronology Rows Tab

The chronology table displays all chronology rows and all source columns:

- DATE
- PAGE NO
- PROVIDER
- DESCRIPTION
- Red Flags/Comments

The displayed cell text must be verbatim to the original `.docx` cell text. The table will not use compact display substitutes such as parsed page ranges in place of the original PAGE NO cell. Existing parsed fields remain available internally for extraction and validation.

The user can select or unselect a row in either of two ways:

- click the checkbox
- single-click anywhere in the row

Clicking a selected row again unselects it. Non-extractable rows, if any, remain unselectable and should still show their warning tooltip.

## Table Sizing

Chronology table cell text will wrap. Row heights will auto-fit to show all text in each row based on the current column widths. Row heights are not persisted.

Column widths are user-resizable and persist globally across sessions and across cases. After widths are restored, the table recalculates row heights so all text remains visible. The persisted setting is global for the Med Record Extractor chronology table, not tied to a case or chronology document.

## Data Model

The parser should expose display text separately from normalized matching text where needed. Existing fields used for matching and extraction should keep their current behavior unless a test shows they rely on display-normalized text.

Required display behavior:

- `SynopsisParagraph.text` or a new display field must preserve the original paragraph text as extracted from Word.
- `SelectableChronologyRow` display fields for date, page number, provider, description, and flags must preserve original cell text.
- Parsed page data (`record_filename`, `page_start`, `page_end`) remains derived from the original PAGE NO cell.

Selection persistence remains source-aware:

- selected synopsis paragraph IDs persist
- selected row IDs persist
- row ownership sources persist so mixed manual and synopsis ownership survives restore

Global column width persistence should use an application-level settings store already available in the codebase if one exists for UI preferences. If no local pattern exists, use `QSettings` with a stable key under the Med Record Extractor viewer.

## Error Handling

If a chronology document cannot be parsed, the viewer continues to show the existing blocking error and disables extraction.

If a row has malformed PAGE NO text, the original PAGE NO text remains visible, the row stays unselectable, and the warning explains that record/pages could not be parsed.

If saved global column widths are missing, invalid, or incomplete, the table uses sensible defaults and writes valid widths the next time the user resizes columns.

## Testing

Add focused tests for:

- Brief Synopsis tab exists and displays full verbatim paragraph text.
- Clicking the synopsis item text toggles selection on and off.
- Chronology Rows tab exists and displays verbatim DATE, PAGE NO, PROVIDER, DESCRIPTION, and Red Flags/Comments cell text.
- Clicking anywhere in an extractable chronology row toggles selection on and off.
- Non-extractable rows cannot remain selected through row clicks.
- Chronology table column widths persist globally and are restored by a new viewer instance.
- Row heights auto-fit after loading rows and after restoring column widths.
- Existing restore behavior still preserves mixed manual and synopsis-owned row selection.

Run the existing Med Record Extractor, viewer, in-process tab, and wizard tests after implementation.
