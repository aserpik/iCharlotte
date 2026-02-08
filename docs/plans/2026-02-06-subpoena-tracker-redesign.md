# Subpoena Tracker Redesign

## Overview

Replace the existing `Scripts/subpoena_tracker.py` (LLM-dependent, subprocess-based) with an in-process `QThread` worker that uses deterministic filesystem scanning to track subpoenas, received records, and medical chronology coverage. Generates a `.docx` report with hyperlinked references.

## Architecture

- **New file:** `icharlotte_core/subpoena_tracker.py` — contains `SubpoenaTrackerWorker(QThread)` and all parsing/generation logic
- **Delete:** `Scripts/subpoena_tracker.py` — old implementation removed entirely
- **Modify:** `iCharlotte.py` — rewire the "Subpoena Tracker" button to launch the in-process worker instead of a subprocess
- **Cleanup:** Remove `agent_subpoena` reference from `icharlotte_core/llm_config.py`

### Execution Model

- Runs in a `QThread` (non-blocking, UI stays responsive)
- Button uses existing `EnhancedAgentButton` for visual consistency (spinner, status)
- Signals: `progress(str)`, `finished(bool, str)`, `warning(str)`
- On success: button shows "Last: Just now", status notification with file path
- On failure: button shows "Last: Failed", error message

## Phase 1: Scan Issued Subpoenas

**Source:** `{case_path}/DISCOVERY/Subpoenas/` (case-insensitive folder matching, recursive walk into subfolders)

**For each `.pdf` file:**

1. Extract vendor-subpoena ID via regex: `r'(\d{3,6})[.\-](\d{1,4})'`
   - Handles varying lengths (3-6 digit vendor, 1-4 digit subpoena)
   - Handles `.` or `-` separator
   - Normalizes to canonical form: `XXXXX-YYYY` (dash-separated, subpoena zero-padded to 4)
2. Extract facility name: strip ID and extension from filename, strip leading separators `[ ,_\-]+`, remainder is facility name
3. Store in dict keyed by normalized ID

**Data structure:**
```python
subpoenas = {
    "62122-0001": {
        "facility": "Riverside County Fire Department",
        "subpoena_path": "Z:\\...\\62122-0001, Riverside County Fire Department.pdf"
    }
}
```

**If folder doesn't exist:** finish early with "No Subpoenas folder found".
**If no parseable PDFs:** finish with "No subpoenas found in Subpoenas folder".
**Unparseable filenames:** skip, collect into warnings list.
**Duplicate IDs:** keep first found, log warning.

## Phase 2: Scan Received Records

**Source:** `{case_path}/RECORDS/Subpoenaed/` (case-insensitive folder matching, recursive walk)

**For each file AND subfolder:**

1. Extract vendor-subpoena ID using same parser as Phase 1
2. Classify the descriptor (text after the ID):
   - Contains `CNR` (case-insensitive) → status = `"CNR"`
   - Contains `reply` or `objection` (case-insensitive) → status = `"Other"`
   - Anything else with a matching ID → status = `"Yes"` (actual responsive records)
3. Store path for hyperlinking (prefer PDF over subfolder if both exist)

**If folder doesn't exist:** all statuses = `"No"` (report still generated).
**Subpoena IDs from Phase 1 with no match here:** status = `"No"` (no hyperlink).

**Data structure:**
```python
received = {
    "62122-0001": {
        "status": "Yes",
        "path": "Z:\\...\\62122-0001, Riverside County Fire Dept.pdf"
    }
}
```

## Phase 3: Scan Medical Chronologies

**Source:** `{case_path}/RECORDS/Medical Summary – DO NOT PRODUCE/` (case-insensitive folder matching, recursive walk)

**If folder doesn't exist:** skip phase, Records Summarized column blank for all rows.

**For each `.docx` file:**

1. Extract date from filename via `r'(\d{4})[.\-](\d{2})[.\-](\d{2})'` — handles `2024-05-02` and `2024.05.02`
2. Open with `python-docx`, iterate tables **in reverse**
3. First **2-column table** found = RECORDS SUMMARIZED table (skips the large 4-column chronology tables entirely)
4. Parse data rows: extract vendor-subpoena ID from the "Filename" column using same parser
5. Map each subpoena ID to this chronology's date and path

**Data structure:**
```python
chronologies = {
    "60563-0002": [
        {"date": "2024-05-02", "path": "Z:\\...\\Medical Summary of Keith Martin.docx"}
    ],
    "60563-0012": [
        {"date": "2024-05-02", "path": "Z:\\...\\docx"},
        {"date": "2025-01-15", "path": "Z:\\...\\docx"}
    ]
}
```

**Error handling:**
- `.docx` can't be opened (corrupted, locked) → skip, log warning
- No 2-column table found → skip, log warning
- No date in filename → skip, log warning

## Phase 4: Generate Output .docx

**Output path:** `{case_path}/NOTES/AI OUTPUT/Tracked_Subpoenas YYYY-MM-DD.docx` (create directory if needed)

**Document structure:**
- Title: "Subpoena Tracking Report" (centered, bold, 14pt Times New Roman)
- Generated date line
- 4-column table (Table Grid style):

| Column | Content | Hyperlink |
|--------|---------|-----------|
| Subpoena Number | Normalized ID (e.g., `62122-0001`) | No |
| Facility | Facility name from Phase 1 | No |
| Records Received | "Yes", "No", "CNR", or "Other" | Yes (except "No") → links to file/folder in RECORDS/Subpoenaed/ |
| Records Summarized | Comma-separated dates of chronologies | Yes → each date links to its .docx chronology |

**Styling:** Times New Roman 12pt, bold header row, Table Grid style.
**Sort order:** Rows sorted by subpoena number ascending.
**After generation:** notify via signal (no auto-open).

## Files Changed

| File | Action |
|------|--------|
| `icharlotte_core/subpoena_tracker.py` | **CREATE** — SubpoenaTrackerWorker class, parsing logic, docx generation |
| `iCharlotte.py` | **MODIFY** — rewire Subpoena Tracker button to launch QThread worker |
| `icharlotte_core/llm_config.py` | **MODIFY** — remove `agent_subpoena` entry |
| `Scripts/subpoena_tracker.py` | **DELETE** — old implementation |
