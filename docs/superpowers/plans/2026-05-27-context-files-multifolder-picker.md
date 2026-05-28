# Multi-folder Context File Picker Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let users pick context files from different folders/subfolders in the Oppose Motion and Respond Discovery wizard tasks via a reusable accumulate-style dialog.

**Architecture:** Add one new modal dialog, `ContextFilesDialog`, that wraps `QFileDialog.getOpenFileNames` in an additive list (each "Add files…" appends, de-duped). Drop it into the two wizard call sites that currently use a single-folder `getOpenFileNames`. The generic `SettingsPage`, Depo Prep, and Med Chron are unchanged (the first two already accumulate; Med Chron is single-file by design).

**Tech Stack:** Python, PySide6 (Qt), pytest + pytest-qt.

---

## File Structure

- `icharlotte_core/ui/context_files_dialog.py` (new) — the reusable `ContextFilesDialog`. Single responsibility: accumulate file paths across folders and return them.
- `icharlotte_core/ui/wizard/pages/oppose_motion_page.py` (modify, 1 call site in `build_oppose_motion_tab`).
- `icharlotte_core/ui/wizard/pages/respond_discovery_page.py` (modify, 1 call site in `_on_select_context_files`).
- `tests/test_context_files_dialog.py` (new) — unit tests for the dialog.
- `tests/test_wizard/test_oppose_motion_page.py` (modify — update 1 existing test, add 1).
- `tests/test_wizard/test_respond_discovery_page.py` (modify — update 2 existing tests, add 1).

Test commands assume the repo root `C:\geminiterminal2` is the working directory and `python` resolves to the project interpreter. pytest-qt must be installed; if it is not, the UI tests `skip` (via `pytest.importorskip("pytestqt")`) rather than fail.

---

### Task 1: Create `ContextFilesDialog`

**Files:**
- Create: `icharlotte_core/ui/context_files_dialog.py`
- Test: `tests/test_context_files_dialog.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_context_files_dialog.py` with this exact content:

```python
import pytest

pytest.importorskip("pytestqt")

from PySide6.QtWidgets import QDialog

from icharlotte_core.ui.context_files_dialog import ContextFilesDialog


def test_add_files_accumulates_across_multiple_folders(qtbot, monkeypatch):
    dlg = ContextFilesDialog(start_dir="/start")
    qtbot.addWidget(dlg)

    calls = [
        ([r"C:\case\DISCOVERY\smith.pdf"], ""),
        ([r"C:\case\RECORDS\med.pdf", r"C:\case\PLEADINGS\complaint.docx"], ""),
    ]

    def fake_get_open(*_args, **_kwargs):
        return calls.pop(0)

    monkeypatch.setattr(
        "icharlotte_core.ui.context_files_dialog.QFileDialog.getOpenFileNames",
        fake_get_open,
    )

    dlg._on_add_files()
    dlg._on_add_files()

    assert dlg.selected_files() == [
        r"C:\case\DISCOVERY\smith.pdf",
        r"C:\case\RECORDS\med.pdf",
        r"C:\case\PLEADINGS\complaint.docx",
    ]


def test_add_files_dedupes_case_insensitively(qtbot, monkeypatch):
    dlg = ContextFilesDialog()
    qtbot.addWidget(dlg)

    calls = [
        ([r"C:\case\smith.pdf"], ""),
        ([r"c:\CASE\Smith.pdf"], ""),  # same file, different case
    ]
    monkeypatch.setattr(
        "icharlotte_core.ui.context_files_dialog.QFileDialog.getOpenFileNames",
        lambda *_a, **_k: calls.pop(0),
    )

    dlg._on_add_files()
    dlg._on_add_files()

    assert len(dlg.selected_files()) == 1


def test_remove_selected_drops_rows(qtbot):
    dlg = ContextFilesDialog(initial=[r"C:\a\one.pdf", r"C:\b\two.pdf", r"C:\c\three.pdf"])
    qtbot.addWidget(dlg)

    dlg.list_widget.item(1).setSelected(True)
    dlg._on_remove_selected()

    assert dlg.selected_files() == [r"C:\a\one.pdf", r"C:\c\three.pdf"]


def test_initial_paths_are_listed(qtbot):
    dlg = ContextFilesDialog(initial=[r"C:\a\one.pdf", r"C:\b\two.pdf"])
    qtbot.addWidget(dlg)

    assert dlg.selected_files() == [r"C:\a\one.pdf", r"C:\b\two.pdf"]
    assert dlg.list_widget.count() == 2


def test_get_files_returns_none_on_cancel(qtbot, monkeypatch):
    monkeypatch.setattr(
        ContextFilesDialog, "exec", lambda self: QDialog.DialogCode.Rejected
    )
    result = ContextFilesDialog.get_files(None, start_dir="/x")
    assert result is None


def test_get_files_returns_list_on_accept(qtbot, monkeypatch):
    def fake_exec(self):
        self._add_paths([r"C:\a\one.pdf"])
        return QDialog.DialogCode.Accepted

    monkeypatch.setattr(ContextFilesDialog, "exec", fake_exec)
    result = ContextFilesDialog.get_files(None)
    assert result == [r"C:\a\one.pdf"]
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_context_files_dialog.py -v`
Expected: FAIL/ERROR with `ModuleNotFoundError: No module named 'icharlotte_core.ui.context_files_dialog'`

- [ ] **Step 3: Write the implementation**

Create `icharlotte_core/ui/context_files_dialog.py` with this exact content:

```python
"""Reusable modal dialog for picking context files across multiple folders.

``QFileDialog.getOpenFileNames`` only allows selecting multiple files within a
single folder. This dialog lets the user accumulate files from different
folders/subfolders: each "Add files…" click opens a normal file browser and
appends the chosen files to a running list (de-duped by normalized path).
"""
from __future__ import annotations

import os
from typing import List, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class ContextFilesDialog(QDialog):
    """Accumulate-style file picker that works across multiple folders."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        title: str = "Select context files",
        start_dir: str = "",
        file_filter: str = "All files (*.*)",
        initial: Optional[List[str]] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(560, 360)
        self._file_filter = file_filter
        self._next_dir = start_dir or ""
        self._paths: List[str] = []
        self._seen: set[str] = set()

        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 16)
        outer.setSpacing(10)

        btn_row = QHBoxLayout()
        self.add_btn = QPushButton("Add files…")
        self.add_btn.clicked.connect(self._on_add_files)
        btn_row.addWidget(self.add_btn)
        self.remove_btn = QPushButton("Remove selected")
        self.remove_btn.clicked.connect(self._on_remove_selected)
        btn_row.addWidget(self.remove_btn)
        btn_row.addStretch()
        outer.addLayout(btn_row)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(
            QListWidget.SelectionMode.ExtendedSelection
        )
        outer.addWidget(self.list_widget, 1)

        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        outer.addWidget(button_box)

        if initial:
            self._add_paths(list(initial))

    @staticmethod
    def _key(path: str) -> str:
        return os.path.normcase(os.path.abspath(path))

    @staticmethod
    def _folder_hint(path: str) -> str:
        parent = os.path.dirname(path)
        return os.path.basename(parent) or parent

    def _add_paths(self, paths: List[str]) -> None:
        for path in paths:
            if not path:
                continue
            key = self._key(path)
            if key in self._seen:
                continue
            self._seen.add(key)
            self._paths.append(path)
            display = os.path.basename(path)
            hint = self._folder_hint(path)
            if hint:
                display = f"{display}  —  {hint}"
            item = QListWidgetItem(display)
            item.setToolTip(path)
            item.setData(Qt.ItemDataRole.UserRole, path)
            self.list_widget.addItem(item)
            self._next_dir = os.path.dirname(path) or self._next_dir

    def _on_add_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Add files", self._next_dir, self._file_filter
        )
        if paths:
            self._add_paths(list(paths))

    def _on_remove_selected(self) -> None:
        rows = sorted(
            {i.row() for i in self.list_widget.selectedIndexes()}, reverse=True
        )
        for r in rows:
            if 0 <= r < len(self._paths):
                removed = self._paths.pop(r)
                self._seen.discard(self._key(removed))
                self.list_widget.takeItem(r)

    def selected_files(self) -> List[str]:
        return list(self._paths)

    @classmethod
    def get_files(
        cls,
        parent: QWidget | None = None,
        *,
        title: str = "Select context files",
        start_dir: str = "",
        file_filter: str = "All files (*.*)",
        initial: Optional[List[str]] = None,
    ) -> Optional[List[str]]:
        """Show modally. Return accumulated list on OK, ``None`` on Cancel."""
        dlg = cls(
            parent,
            title=title,
            start_dir=start_dir,
            file_filter=file_filter,
            initial=initial,
        )
        if dlg.exec() == QDialog.DialogCode.Accepted:
            return dlg.selected_files()
        return None
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_context_files_dialog.py -v`
Expected: PASS (6 passed), or SKIPPED if pytest-qt is not installed.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/context_files_dialog.py tests/test_context_files_dialog.py
git commit -m "feat(ui): add ContextFilesDialog for multi-folder context file picking"
```

---

### Task 2: Integrate into Oppose Motion

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/oppose_motion_page.py` (imports + `build_oppose_motion_tab`)
- Test: `tests/test_wizard/test_oppose_motion_page.py` (update `test_builder_filters_unsupported_context_files`, add `test_builder_accumulates_multifolder_context_files`)

- [ ] **Step 1: Update the existing test and add a new one (write failing tests first)**

In `tests/test_wizard/test_oppose_motion_page.py`, REPLACE the existing
`test_builder_filters_unsupported_context_files` function (currently around
lines 804-824) with the following two functions:

```python
def test_builder_filters_unsupported_context_files(qtbot, tmp_path, monkeypatch):
    spec = type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})()
    motion = tmp_path / "motion.pdf"
    motion.write_bytes(b"")
    good_context = tmp_path / "facts.txt"
    bad_context = tmp_path / "notes.xlsx"
    good_context.write_text("facts")
    bad_context.write_text("spreadsheet")
    monkeypatch.setattr(OpposeMotionTaskTab, "_start_analysis", lambda self: None)

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getOpenFileName",
        return_value=(str(motion), ""),
    ), patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.ContextFilesDialog.get_files",
        return_value=[str(good_context), str(bad_context)],
    ):
        tab = build_oppose_motion_tab(spec, str(tmp_path), "0000.000", None)

    qtbot.addWidget(tab)
    assert tab.settings_page.context_files == [str(good_context)]


def test_builder_accumulates_multifolder_context_files(qtbot, tmp_path, monkeypatch):
    spec = type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})()
    motion = tmp_path / "motion.pdf"
    motion.write_bytes(b"")
    folder_a = tmp_path / "DISCOVERY"
    folder_b = tmp_path / "RECORDS"
    folder_a.mkdir()
    folder_b.mkdir()
    ctx_a = folder_a / "depo.pdf"
    ctx_b = folder_b / "records.pdf"
    ctx_a.write_text("a")
    ctx_b.write_text("b")
    monkeypatch.setattr(OpposeMotionTaskTab, "_start_analysis", lambda self: None)

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getOpenFileName",
        return_value=(str(motion), ""),
    ), patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.ContextFilesDialog.get_files",
        return_value=[str(ctx_a), str(ctx_b)],
    ):
        tab = build_oppose_motion_tab(spec, str(tmp_path), "0000.000", None)

    qtbot.addWidget(tab)
    assert tab.settings_page.context_files == [str(ctx_a), str(ctx_b)]


def test_builder_handles_cancelled_context_picker(qtbot, tmp_path, monkeypatch):
    spec = type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})()
    motion = tmp_path / "motion.pdf"
    motion.write_bytes(b"")
    monkeypatch.setattr(OpposeMotionTaskTab, "_start_analysis", lambda self: None)

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getOpenFileName",
        return_value=(str(motion), ""),
    ), patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.ContextFilesDialog.get_files",
        return_value=None,
    ):
        tab = build_oppose_motion_tab(spec, str(tmp_path), "0000.000", None)

    qtbot.addWidget(tab)
    assert tab.settings_page.context_files == []
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_wizard/test_oppose_motion_page.py -v -k "builder_filters or builder_accumulates or builder_handles_cancelled"`
Expected: FAIL — `AttributeError` / patch target `ContextFilesDialog` does not exist on the module yet (the name isn't imported).

- [ ] **Step 3: Add the import**

In `icharlotte_core/ui/wizard/pages/oppose_motion_page.py`, add this import next to the other `icharlotte_core.ui` / opposition imports near the top of the file (after the existing `from icharlotte_core.ui.wizard.pages.status_page import StatusPage` line):

```python
from icharlotte_core.ui.context_files_dialog import ContextFilesDialog
```

- [ ] **Step 4: Replace the context-file selection in `build_oppose_motion_tab`**

In `icharlotte_core/ui/wizard/pages/oppose_motion_page.py`, find this block in `build_oppose_motion_tab`:

```python
    context_files, _ = QFileDialog.getOpenFileNames(
        parent,
        "Select context document(s)",
        os.path.dirname(motion_file) or case_path,
        "Context files (*.pdf *.docx *.txt *.msg);;All files (*.*)",
    )
    context_files = [
        path for path in (context_files or []) if is_supported_context_file(path)
    ]
```

Replace it with:

```python
    context_files = ContextFilesDialog.get_files(
        parent,
        title="Select context document(s)",
        start_dir=os.path.dirname(motion_file) or case_path,
        file_filter="Context files (*.pdf *.docx *.txt *.msg);;All files (*.*)",
    )
    context_files = [
        path for path in (context_files or []) if is_supported_context_file(path)
    ]
```

(`context_files or []` already handles the `None` cancel case, leaving an empty
context list — matching the prior behavior where a cancelled dialog returned `[]`.)

- [ ] **Step 5: Run the tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_oppose_motion_page.py -v`
Expected: PASS (all existing oppose-motion tests plus the 3 builder tests), or SKIPPED without pytest-qt.

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/oppose_motion_page.py tests/test_wizard/test_oppose_motion_page.py
git commit -m "feat(oppose-motion): pick context files across folders via ContextFilesDialog"
```

---

### Task 3: Integrate into Respond Discovery

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/respond_discovery_page.py` (imports + `_on_select_context_files`)
- Test: `tests/test_wizard/test_respond_discovery_page.py` (update 2 existing tests, add 1)

- [ ] **Step 1: Update the two existing tests and add a cancel test (write failing tests first)**

In `tests/test_wizard/test_respond_discovery_page.py`, REPLACE the two existing
functions `test_context_file_picker_starts_in_status_folder_next_to_discovery_file`
and `test_context_file_picker_prefers_case_status_folder` (currently around lines
406-455) with the following three functions. Note: the start folder is now passed
as the keyword argument `start_dir`, and `get_files` returns a list (OK) or `None`
(cancel) rather than the `(paths, filter)` tuple from `getOpenFileNames`.

```python
    def test_context_file_picker_starts_in_status_folder_next_to_discovery_file(self):
        with tempfile.TemporaryDirectory(dir="C:\\geminiterminal2") as tmp:
            propounded = Path(tmp) / "DISCOVERY" / "PROPOUNDED"
            status = propounded / "STATUS"
            status.mkdir(parents=True)
            discovery_file = propounded / "srogg.pdf"
            discovery_file.write_text("SPECIAL INTERROGATORIES")

            page = RespondDiscoverySettingsPage(
                case_root=tmp,
                file_number="1234.001",
                discovery_file=str(discovery_file),
                detected_type="SI",
            )

            with patch.object(page, "_generate_proposals") as mock_generate:
                with patch(
                    "icharlotte_core.ui.wizard.pages.respond_discovery_page.ContextFilesDialog.get_files",
                    return_value=[],
                ) as mock_dialog:
                    page._on_select_context_files()

            self.assertEqual(mock_dialog.call_args.kwargs["start_dir"], str(status))
            mock_generate.assert_called_once()

    def test_context_file_picker_prefers_case_status_folder(self):
        with tempfile.TemporaryDirectory(dir="C:\\geminiterminal2") as tmp:
            case_status = Path(tmp) / "STATUS"
            case_status.mkdir()
            propounded = Path(tmp) / "DISCOVERY" / "PROPOUNDED"
            propounded.mkdir(parents=True)
            discovery_file = propounded / "srogg.pdf"
            discovery_file.write_text("SPECIAL INTERROGATORIES")

            page = RespondDiscoverySettingsPage(
                case_root=tmp,
                file_number="3000.075",
                discovery_file=str(discovery_file),
                detected_type="SI",
            )

            with patch.object(page, "_generate_proposals") as mock_generate:
                with patch(
                    "icharlotte_core.ui.wizard.pages.respond_discovery_page.ContextFilesDialog.get_files",
                    return_value=[],
                ) as mock_dialog:
                    page._on_select_context_files()

            self.assertEqual(mock_dialog.call_args.kwargs["start_dir"], str(case_status))
            mock_generate.assert_called_once()

    def test_context_file_picker_cancel_aborts_generation(self):
        with tempfile.TemporaryDirectory(dir="C:\\geminiterminal2") as tmp:
            propounded = Path(tmp) / "DISCOVERY" / "PROPOUNDED"
            propounded.mkdir(parents=True)
            discovery_file = propounded / "srogg.pdf"
            discovery_file.write_text("SPECIAL INTERROGATORIES")

            page = RespondDiscoverySettingsPage(
                case_root=tmp,
                file_number="1234.001",
                discovery_file=str(discovery_file),
                detected_type="SI",
            )

            with patch.object(page, "_generate_proposals") as mock_generate:
                with patch(
                    "icharlotte_core.ui.wizard.pages.respond_discovery_page.ContextFilesDialog.get_files",
                    return_value=None,
                ):
                    page._on_select_context_files()

            mock_generate.assert_not_called()
            self.assertEqual(page.context_files, [])
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_wizard/test_respond_discovery_page.py -v -k "context_file_picker"`
Expected: FAIL — patch target `ContextFilesDialog` does not exist on the module yet.

- [ ] **Step 3: Add the import**

In `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`, add this import
after the existing `from icharlotte_core.discovery._io import ( read_document_text )`
import block near the top:

```python
from icharlotte_core.ui.context_files_dialog import ContextFilesDialog
```

- [ ] **Step 4: Replace `_on_select_context_files`**

In `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`, find:

```python
    def _on_select_context_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select context file(s)",
            context_file_start_dir(self.discovery_file, self.case_root),
            "Context files (*.pdf *.docx *.txt);;All files (*.*)",
        )
        self.context_files = list(paths or [])
        self._generate_proposals()
```

Replace it with:

```python
    def _on_select_context_files(self) -> None:
        paths = ContextFilesDialog.get_files(
            self,
            title="Select context file(s)",
            start_dir=context_file_start_dir(self.discovery_file, self.case_root),
            file_filter="Context files (*.pdf *.docx *.txt);;All files (*.*)",
        )
        if paths is None:
            return  # user cancelled — stay on the rules screen, do not generate
        self.context_files = list(paths)
        self._generate_proposals()
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_respond_discovery_page.py -v`
Expected: PASS (all existing respond-discovery tests plus the 3 context-picker tests), or SKIPPED without pytest-qt.

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/respond_discovery_page.py tests/test_wizard/test_respond_discovery_page.py
git commit -m "feat(respond-discovery): pick context files across folders via ContextFilesDialog"
```

---

### Task 4: Full regression check

**Files:** none (verification only)

- [ ] **Step 1: Run the full wizard + dialog test suites**

Run: `python -m pytest tests/test_context_files_dialog.py tests/test_wizard/ -v`
Expected: All PASS or SKIPPED (no failures, no errors).

- [ ] **Step 2: Manual smoke test (per project rule "always test after changing a feature")**

Launch the app (`python iCharlotte.py`), open a case, start the **Oppose a Motion**
wizard task, pick a motion, then in the context dialog: click "Add files…", choose
a file from one folder; click "Add files…" again, navigate to a *different*
subfolder and choose another file; confirm both appear in the list; remove one;
click OK. Repeat for **Respond to Discovery** (the "Next: Context Files" button).
Confirm Cancel on the dialog backs out without starting generation.

If the UI cannot be exercised in this environment, state that explicitly rather
than claiming success.

- [ ] **Step 3: No commit** (verification only; nothing changed).
