# Import Carrier Reports Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add an "Import Reports" button to the ChatTab that scans `{case_path}/STATUS/` for carrier report documents (`carrier001.doc(x)` through `carrier015.doc(x)`, with optional trailing text, but no prefix) and attaches them to the chat's documents list.

**Architecture:** Pure UI addition in a single file (`icharlotte_core/ui/tabs.py`). Adds a module-level compiled regex constant, one new button in the existing `file_btn_layout`, and one new method `import_carrier_reports` that reuses the existing `add_file` method for all attachment logic (icons, persistence, dedup). Popups via `QMessageBox` use the same pattern as existing warnings in the file.

**Tech Stack:** Python 3, PySide6 (`QMessageBox`, `QPushButton`), stdlib `re` and `os`.

**Spec:** `docs/superpowers/specs/2026-04-14-import-carrier-reports-design.md`

---

## File Structure

**Modify:**
- `icharlotte_core/ui/tabs.py`
  - Add `import re` near top (line ~1–10, with other stdlib imports).
  - Add module-level constant `CARRIER_REPORT_RE` (after imports, before first class).
  - Add button to `file_btn_layout` (around line 350–356).
  - Add new method `import_carrier_reports` on ChatTab (after `clear_files`, ~line 923).

**Create:**
- `tests/test_import_carrier_reports.py`
  - Unit tests for the regex pattern (no Qt required).

No other files need to change. `add_file`, persistence, list-widget handling, and case-path access are all already in place and unchanged.

---

## Task 1: Add regex constant and unit tests

**Files:**
- Modify: `icharlotte_core/ui/tabs.py` (add `import re` and `CARRIER_REPORT_RE` constant)
- Create: `tests/test_import_carrier_reports.py`

- [ ] **Step 1: Write the failing test**

Create `tests/test_import_carrier_reports.py`:

```python
"""Tests for the Import Reports feature in ChatTab."""
import unittest


class TestCarrierReportRegex(unittest.TestCase):
    """Verify CARRIER_REPORT_RE matches the spec's must/must-not lists."""

    def setUp(self):
        from icharlotte_core.ui.tabs import CARRIER_REPORT_RE
        self.pattern = CARRIER_REPORT_RE

    def _matches(self, name: str) -> bool:
        return self.pattern.match(name) is not None

    # --- Must match -----------------------------------------------------
    def test_basic_carrier001_docx(self):
        self.assertTrue(self._matches("carrier001.docx"))

    def test_carrier015_doc_lowercase(self):
        self.assertTrue(self._matches("carrier015.doc"))

    def test_trailing_space_and_parens(self):
        self.assertTrue(self._matches("carrier002 (FSR).docx"))

    def test_trailing_parens_no_space(self):
        self.assertTrue(self._matches("carrier003(lit plan).docx"))

    def test_uppercase_carrier_and_extension(self):
        self.assertTrue(self._matches("Carrier007.DOCX"))

    def test_trailing_dash_suffix(self):
        self.assertTrue(self._matches("carrier010 - Final.docx"))

    def test_all_caps_carrier(self):
        self.assertTrue(self._matches("CARRIER005.docx"))

    # --- Must NOT match -------------------------------------------------
    def test_bracket_prefix_rejected(self):
        self.assertFalse(self._matches("[draft]carrier001.docx"))

    def test_word_prefix_rejected(self):
        self.assertFalse(self._matches("draft_carrier001.docx"))

    def test_carrier000_below_range(self):
        self.assertFalse(self._matches("carrier000.docx"))

    def test_carrier016_above_range(self):
        self.assertFalse(self._matches("carrier016.docx"))

    def test_four_digit_run_rejected(self):
        self.assertFalse(self._matches("carrier0011.docx"))

    def test_wrong_extension_pdf(self):
        self.assertFalse(self._matches("carrier001.pdf"))

    def test_no_number(self):
        self.assertFalse(self._matches("carrier.docx"))

    def test_two_digit_number(self):
        self.assertFalse(self._matches("carrier01.docx"))

    def test_carrier100_out_of_range(self):
        self.assertFalse(self._matches("carrier100.docx"))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_import_carrier_reports.py -v`
Expected: All tests FAIL with `ImportError: cannot import name 'CARRIER_REPORT_RE' from 'icharlotte_core.ui.tabs'`.

- [ ] **Step 3: Add `import re` to tabs.py**

In `icharlotte_core/ui/tabs.py`, line 1–10, add `re` to the stdlib imports. The current block is:

```python
import os
import sys
import json
import subprocess
import markdown
import shutil
import datetime
import base64
import time
from functools import partial
```

Change to:

```python
import os
import re
import sys
import json
import subprocess
import markdown
import shutil
import datetime
import base64
import time
from functools import partial
```

- [ ] **Step 4: Add the `CARRIER_REPORT_RE` constant**

In `icharlotte_core/ui/tabs.py`, immediately after the `try: import pypdf / except: ...` block (around line 48) and before `class DateTableWidgetItem`, add:

```python
# Matches carrier report filenames: carrier001..carrier015 with optional
# trailing text (e.g. " (FSR)", "(lit plan)", " - Final"), .doc or .docx.
# Anchored at start to reject prefixed variants like "[draft]carrier001.docx".
CARRIER_REPORT_RE = re.compile(
    r'^carrier0(0[1-9]|1[0-5])(?![0-9]).*\.docx?$',
    re.IGNORECASE,
)
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest tests/test_import_carrier_reports.py -v`
Expected: All 16 tests PASS.

- [ ] **Step 6: Commit**

```bash
git add tests/test_import_carrier_reports.py icharlotte_core/ui/tabs.py
git commit -m "feat(chat): add CARRIER_REPORT_RE for carrier report filename matching"
```

---

## Task 2: Add the "Import Reports" button to the file button row

**Files:**
- Modify: `icharlotte_core/ui/tabs.py:342-356` (file_btn_layout)

- [ ] **Step 1: Add the button in `file_btn_layout`**

In `icharlotte_core/ui/tabs.py`, locate the existing block (around lines 342–356):

```python
        # Select All / Deselect All / Clear Files buttons
        file_btn_layout = QHBoxLayout()
        select_all_btn = QPushButton("All")
        select_all_btn.setToolTip("Select all files")
        select_all_btn.clicked.connect(lambda: self._set_all_file_checks(Qt.CheckState.Checked))
        deselect_all_btn = QPushButton("None")
        deselect_all_btn.setToolTip("Deselect all files")
        deselect_all_btn.clicked.connect(lambda: self._set_all_file_checks(Qt.CheckState.Unchecked))
        clear_files_btn = QPushButton("Clear")
        clear_files_btn.setToolTip("Remove all files")
        clear_files_btn.clicked.connect(self.clear_files)
        file_btn_layout.addWidget(select_all_btn)
        file_btn_layout.addWidget(deselect_all_btn)
        file_btn_layout.addWidget(clear_files_btn)
        settings_layout.addLayout(file_btn_layout)
```

Change to:

```python
        # Select All / Deselect All / Clear Files / Import Reports buttons
        file_btn_layout = QHBoxLayout()
        select_all_btn = QPushButton("All")
        select_all_btn.setToolTip("Select all files")
        select_all_btn.clicked.connect(lambda: self._set_all_file_checks(Qt.CheckState.Checked))
        deselect_all_btn = QPushButton("None")
        deselect_all_btn.setToolTip("Deselect all files")
        deselect_all_btn.clicked.connect(lambda: self._set_all_file_checks(Qt.CheckState.Unchecked))
        clear_files_btn = QPushButton("Clear")
        clear_files_btn.setToolTip("Remove all files")
        clear_files_btn.clicked.connect(self.clear_files)
        import_reports_btn = QPushButton("Import Reports")
        import_reports_btn.setToolTip("Import carrier reports from the case's STATUS folder")
        import_reports_btn.clicked.connect(self.import_carrier_reports)
        file_btn_layout.addWidget(select_all_btn)
        file_btn_layout.addWidget(deselect_all_btn)
        file_btn_layout.addWidget(clear_files_btn)
        file_btn_layout.addWidget(import_reports_btn)
        settings_layout.addLayout(file_btn_layout)
```

- [ ] **Step 2: Verify tabs.py imports parse (syntax check)**

Run: `python -c "import ast; ast.parse(open(r'icharlotte_core/ui/tabs.py', encoding='utf-8').read()); print('OK')"`
Expected: `OK`

(Full app startup will fail at this point because `self.import_carrier_reports` does not exist yet — Task 3 adds it. That's fine; do not launch the app until Task 3 is complete.)

- [ ] **Step 3: Do NOT commit yet**

This task is intentionally paired with Task 3 — committing now would leave the app in a broken state where clicking the button raises `AttributeError`. Continue straight to Task 3.

---

## Task 3: Implement the `import_carrier_reports` method

**Files:**
- Modify: `icharlotte_core/ui/tabs.py` (add method after `clear_files`, ~line 923)

- [ ] **Step 1: Add the new method**

In `icharlotte_core/ui/tabs.py`, locate the `clear_files` method (around line 913–922):

```python
    def clear_files(self):
        """Clear all attached files and persist the change."""
        if self.attached_files is None:
            self.attached_files = []
        else:
            self.attached_files.clear()
        self.file_list.clear()
        # Persist the cleared state
        if self.persistence:
            self.persistence.clear_attached_files()
```

Immediately AFTER this method (before `def _clear_files_no_persist`), insert:

```python
    def import_carrier_reports(self):
        """Scan {case_path}/STATUS/ for carrier00X.doc(x) files and attach them.

        Matches carrier001..carrier015 with optional trailing text after the
        number. Rejects filenames with any prefix before "carrier". Adds matches
        to the existing attached-files list (deduped), leaving current files
        intact. Shows a popup summarizing the result.
        """
        main_win = self.window()
        case_path = getattr(main_win, 'case_path', None)
        if not case_path:
            QMessageBox.warning(
                self,
                "Import Reports",
                "No case is currently selected.",
            )
            return

        status_dir = os.path.join(case_path, "STATUS")
        if not os.path.isdir(status_dir):
            QMessageBox.warning(
                self,
                "Import Reports",
                f"No STATUS folder found at:\n{status_dir}",
            )
            return

        try:
            entries = os.listdir(status_dir)
        except (PermissionError, OSError) as e:
            QMessageBox.warning(
                self,
                "Import Reports",
                f"Could not read STATUS folder:\n{e}",
            )
            return

        imported = 0
        already_attached = 0
        for name in entries:
            if not CARRIER_REPORT_RE.match(name):
                continue
            full_path = os.path.join(status_dir, name)
            if full_path in self.attached_files:
                already_attached += 1
                continue
            self.add_file(full_path)
            imported += 1

        if imported == 0 and already_attached == 0:
            QMessageBox.information(
                self,
                "Import Reports",
                "No carrier reports (carrier001\u2013carrier015) found in STATUS.",
            )
        elif imported == 0 and already_attached > 0:
            QMessageBox.information(
                self,
                "Import Reports",
                f"All {already_attached} matching report(s) were already attached.",
            )
        else:
            msg = f"Imported {imported} carrier report(s) from STATUS."
            if already_attached > 0:
                msg += f"\n({already_attached} already attached, skipped.)"
            QMessageBox.information(self, "Import Reports", msg)
```

Note: `\u2013` is the en-dash used in the spec's "carrier001–carrier015" range display.

- [ ] **Step 2: Verify tabs.py parses**

Run: `python -c "import ast; ast.parse(open(r'icharlotte_core/ui/tabs.py', encoding='utf-8').read()); print('OK')"`
Expected: `OK`

- [ ] **Step 3: Run the existing tests (regression check)**

Run: `python -m pytest tests/test_import_carrier_reports.py -v`
Expected: All 16 regex tests still PASS (unchanged from Task 1).

- [ ] **Step 4: Commit the button and method together**

```bash
git add icharlotte_core/ui/tabs.py
git commit -m "feat(chat): add Import Reports button to scan case STATUS folder"
```

---

## Task 4: Manual verification in the running app

Per CLAUDE.md: *"Always test after developing or changing a feature."* This is mandatory.

- [ ] **Step 1: Prepare a test STATUS folder**

Pick an existing case (or make a throwaway one). In a file explorer, inside `{case_path}/STATUS/`, ensure at least these files exist (create empty `.docx` files if needed — the feature only reads filenames):

Must-match fixtures:
- `carrier001.docx`
- `carrier002 (FSR).docx`

Must-not-match fixture:
- `[draft]carrier003.docx`

If STATUS does not already exist for your test case, create it.

- [ ] **Step 2: Launch the app**

Run: `python iCharlotte.py`
Expected: App launches without errors. Load the test case.

- [ ] **Step 3: Verify button placement**

Go to the Chat tab. In the left settings panel under the file list, confirm the button row now reads: `[All] [None] [Clear] [Import Reports]`.

- [ ] **Step 4: Import into an empty list**

Ensure the file list is empty (click **Clear** if needed). Click **Import Reports**.
Expected:
- `carrier001.docx` and `carrier002 (FSR).docx` appear in the documents list, both checked.
- `[draft]carrier003.docx` does NOT appear.
- Popup: "Imported 2 carrier report(s) from STATUS."

- [ ] **Step 5: Import again (dedup path)**

With both files still in the list, click **Import Reports** a second time.
Expected:
- File list is unchanged (still 2 items).
- Popup: "All 2 matching report(s) were already attached."

- [ ] **Step 6: Import when mixing existing and new**

Click **Clear**, then manually attach some unrelated file via **Select File(s)** (e.g., any random `.pdf`). Click **Import Reports**.
Expected:
- The unrelated file is still attached.
- The 2 carrier reports are added alongside it.
- Popup: "Imported 2 carrier report(s) from STATUS."

- [ ] **Step 7: No STATUS folder path**

Switch to a case whose folder does NOT have a STATUS subdirectory (or temporarily rename the STATUS folder on your test case). Click **Import Reports**.
Expected: Warning popup `"No STATUS folder found at:\n{path}"`. File list unchanged.

Restore the STATUS folder name if you renamed it.

- [ ] **Step 8: No case loaded path**

Close the active case (or launch the app fresh with no case selected if that's the normal state). Go to the Chat tab. Click **Import Reports**.
Expected: Warning popup "No case is currently selected." File list unchanged.

- [ ] **Step 9: Zero-match STATUS folder**

On a case whose STATUS folder exists but has no `carrier00X` files, click **Import Reports**.
Expected: Info popup "No carrier reports (carrier001–carrier015) found in STATUS." File list unchanged.

- [ ] **Step 10: Document the result**

If any step fails, stop and debug — do not mark the feature complete. If all steps pass, proceed.

- [ ] **Step 11: Final commit (if any fixes were needed)**

If Steps 1–10 required no changes, skip this step. If a bug was found and fixed during manual testing, commit the fix:

```bash
git add icharlotte_core/ui/tabs.py
git commit -m "fix(chat): <specific fix description>"
```

---

## Completion Checklist

- [ ] Regex unit tests pass (`pytest tests/test_import_carrier_reports.py`)
- [ ] `tabs.py` parses via `ast.parse`
- [ ] Button row shows `[All] [None] [Clear] [Import Reports]`
- [ ] All 10 manual verification steps pass
- [ ] Commits pushed (or left for user to push):
  1. `feat(chat): add CARRIER_REPORT_RE for carrier report filename matching`
  2. `feat(chat): add Import Reports button to scan case STATUS folder`
  3. (optional) any bugfix from manual testing

---

## Self-Review Notes

- **Spec coverage:**
  - UI change (button in `file_btn_layout`) → Task 2
  - `import_carrier_reports` method with all 6 flow steps → Task 3
  - Regex with exact pattern from spec → Task 1
  - All 8 must-match + 8 must-not-match test cases from spec → Task 1 (16 tests)
  - Error handling matrix (no case / no dir / listdir fail / zero match / all attached / mixed) → Task 3 method + Task 4 manual steps 4–9
  - Dedup via `self.attached_files` membership check → Task 3 Step 1
  - Reuses `add_file` (no changes to it) → Task 3 Step 1
- **Placeholder scan:** No TBDs, no "similar to above", no "add error handling" — all code inline.
- **Type consistency:** `CARRIER_REPORT_RE`, `import_carrier_reports`, `case_path`, `status_dir`, `self.attached_files`, `self.add_file` — same identifiers in every task where they appear.
