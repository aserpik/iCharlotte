# Import Carrier Reports — Design

**Date:** 2026-04-14
**Component:** ChatTab (`icharlotte_core/ui/tabs.py`)

## Goal

Add an "Import Reports" button to the ChatTab that scans the selected case's `STATUS` subfolder for carrier report documents (`carrier001.doc(x)` through `carrier015.doc(x)`) and attaches any matches to the chat's documents list.

## User Story

While working in the chat tab on a case, the user clicks **Import Reports**. The app looks in `{case_path}/STATUS/` for files matching the carrier report naming convention and adds any matches to the existing attached-files list. A popup reports how many were imported.

## UI Change

In `icharlotte_core/ui/tabs.py`, ChatTab file-button row (around line 343–356), add a fourth button to the existing `file_btn_layout`:

```
[All] [None] [Clear] [Import Reports]
```

- Button: `QPushButton("Import Reports")`
- Tooltip: `"Import carrier reports from the case's STATUS folder"`
- Signal: `clicked` → new method `self.import_carrier_reports`
- Position: appended to `file_btn_layout` after `clear_files_btn`

No other layout changes.

## New Method: `import_carrier_reports`

Location: ChatTab, added near `clear_files` (~line 913).

### Flow

1. **Resolve case path**
   ```python
   main_win = self.window()
   case_path = getattr(main_win, 'case_path', None)
   ```
   If falsy → `QMessageBox.warning` "No case is currently selected." → return.

2. **Resolve STATUS folder**
   ```python
   status_dir = os.path.join(case_path, "STATUS")
   ```
   If not `os.path.isdir(status_dir)` → `QMessageBox.warning` `f"No STATUS folder found at:\n{status_dir}"` → return.

3. **Scan folder** (non-recursive)
   ```python
   try:
       entries = os.listdir(status_dir)
   except (PermissionError, OSError) as e:
       QMessageBox.warning(self, "Import Reports",
                           f"Could not read STATUS folder:\n{e}")
       return
   ```

4. **Match and import**
   - Compile `CARRIER_REPORT_RE` (see Regex section).
   - For each matching filename, build `full_path = os.path.join(status_dir, name)`.
   - Track counts: `imported`, `already_attached`.
   - Before calling `add_file`, check `if full_path in self.attached_files: already_attached += 1; continue`.
   - Otherwise call `self.add_file(full_path)` and `imported += 1`.
   - `add_file` handles list-widget item creation, icon assignment, check-state, and persistence — no duplication needed.

5. **Result popup** (`QMessageBox.information`)
   - Zero matches: `"No carrier reports (carrier001–carrier015) found in STATUS."`
   - Some imported: `f"Imported {imported} carrier report(s) from STATUS."` — if `already_attached > 0`, append `f"\n({already_attached} already attached, skipped.)"`.
   - All already attached (imported=0, already_attached>0): `f"All {already_attached} matching report(s) were already attached."`

### Regex

```python
CARRIER_REPORT_RE = re.compile(
    r'^carrier0(0[1-9]|1[0-5])(?![0-9]).*\.docx?$',
    re.IGNORECASE,
)
```

**Anchors and components:**

| Part | Purpose |
|---|---|
| `^carrier` | Must start at beginning — rejects `[draft]carrier001.docx`. |
| `0(0[1-9]\|1[0-5])` | Matches `001`–`015` only (3-digit, zero-padded). |
| `(?![0-9])` | Negative lookahead — rejects `carrier0016`, `carrier00150`. |
| `.*` | Allows any trailing text: `(FSR)`, `(lit plan)`, `- Final`, or nothing. |
| `\.docx?$` | `.doc` or `.docx` only. |
| `re.IGNORECASE` | Matches `Carrier`, `CARRIER`, `.DOCX`, etc. |

Compile once at module scope (or as a class constant) to avoid recompiling on each click.

### Test Cases

**Must match:**
- `carrier001.docx`
- `carrier015.doc`
- `carrier002 (FSR).docx`
- `carrier003(lit plan).docx`
- `Carrier007.DOCX`
- `carrier010 - Final.docx`
- `carrier001.doc` (lowercase `.doc`)

**Must NOT match:**
- `[draft]carrier001.docx` (prefix)
- `draft_carrier001.docx` (prefix)
- `carrier000.docx` (below range)
- `carrier016.docx` (above range)
- `carrier0011.docx` (4-digit run)
- `carrier001.pdf` (wrong extension)
- `carrier.docx` (no number)
- `carrier01.docx` (2-digit)

## Error Handling Summary

| Condition | Behavior |
|---|---|
| No case selected | Warning popup, return. |
| `STATUS` dir missing | Warning popup with path, return. |
| `os.listdir` fails | Warning popup with error, return. |
| Zero matches in folder | Info popup "No carrier reports found." |
| All matches already attached | Info popup "All N reports already attached." |
| Some/all imported | Info popup with count + skipped count if any. |

## Dependency on Existing Code

Relies on:
- `self.attached_files` — the list tracked by ChatTab.
- `self.add_file(path)` — existing method at ~line 863. Handles PDF OCR check (irrelevant here since `.doc`/`.docx` won't hit the PDF branch), icon/list-item creation, and persistence via `self.persistence`.
- `self.window().case_path` — same pattern used by `_start_mediation_brief_generation` (tabs.py:1912).

No changes required to `add_file`, `clear_files`, persistence layer, or any other file.

## Testing Plan

**Unit test** — new test file or additions to existing UI tests:
- Test the regex against the full fixture list above (must-match + must-not-match).
- Can be a pure-Python test with no Qt dependency: `assert bool(CARRIER_REPORT_RE.match(name)) == expected`.

**Manual test:**
1. Open iCharlotte, load a case that has a `STATUS` subfolder with at least `carrier001.docx` and `carrier002 (FSR).docx`, plus a non-matching file like `[draft]carrier003.docx`.
2. Go to the Chat tab.
3. Verify the row shows `[All] [None] [Clear] [Import Reports]`.
4. Click **Import Reports**.
5. Confirm:
   - Both matching files appear in the documents list, checked.
   - `[draft]carrier003.docx` is NOT in the list.
   - Popup shows "Imported 2 carrier report(s) from STATUS."
6. Click **Import Reports** again.
7. Confirm popup shows "All 2 matching report(s) were already attached." and the list is unchanged.
8. Test on a case with no STATUS folder → confirm warning popup.
9. Test with no case loaded → confirm warning popup.

## Out of Scope

- Recursive scanning of `STATUS` subdirectories.
- Supporting other extensions (`.pdf`, `.txt`, etc.).
- Sorting / grouping imported files in the list.
- Changing `add_file` behavior.
- Any persistence schema changes.
