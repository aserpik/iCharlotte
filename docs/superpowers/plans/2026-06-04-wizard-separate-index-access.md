# Wizard Separate → Persistent, Viewable Document Index — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Persist Wizard-mode Separate runs to the same per-case index store Advanced mode reads, and add a corner button on the Separate launcher card that reveals the Index tab (re-hideable via its "x").

**Architecture:** A new UI-free `case_index_store` module writes the existing `{file_number}_index.json` format. The wizard `SeparateTaskTab` calls it on analyze + process. `TaskSpec` gains optional card-action fields; `TaskCard` renders a corner `QToolButton` that emits `action_requested`; `WizardTab` re-emits `card_action_requested`; the main window reveals the hidden `IndexTab` singleton and lets its "x" re-hide it in wizard mode.

**Tech Stack:** Python 3, PySide6, pypdf, `unittest` (offscreen Qt via `QT_QPA_PLATFORM=offscreen`).

---

## Before you start (environment)

- **Isolated worktree.** `C:\geminiterminal2` is a live, shared checkout currently on the unrelated branch `feature/generate-motion-detailed-outline` with other sessions' uncommitted changes. Per the Separate→Wizard precedent, execute this plan in an isolated **git worktree** on a new branch (e.g. `feature/wizard-separate-index`) created via `superpowers:using-git-worktrees`. All task commits target that branch.
- **Live verification uses the MAIN checkout.** Worktree files are invisible to the running app (it runs from `C:\geminiterminal2`). The automated tests (Tasks 1–5) run fine in the worktree. The **live** manual check in Task 6 must be done from the main checkout after the branch is merged/applied, then iCharlotte restarted.
- **Reference spec:** `docs/superpowers/specs/2026-06-04-wizard-separate-index-access-design.md`.

## File Structure

- **Create** `icharlotte_core/case_index_store.py` — shared per-case index reader/writer (UI-free).
- **Modify** `icharlotte_core/ui/wizard/pages/separate_page.py` — persist on analyze + process; add `_docs_from_workbench()` + `_persist_to_index()`.
- **Modify** `icharlotte_core/ui/wizard/registry.py` — add optional `card_action_id/glyph/tooltip` to `TaskSpec`; set them on the `separate` spec.
- **Modify** `icharlotte_core/ui/wizard/task_card.py` — corner `QToolButton` + `action_requested` signal.
- **Modify** `icharlotte_core/ui/wizard/wizard_tab.py` — `card_action_requested` signal + re-emit.
- **Modify** `iCharlotte.py` — connect signal; `_on_card_action`; `_reveal_index_tab`; re-hide logic in `_hide_fixed_close_buttons` / `_on_tab_close_requested`; `_hide_fixed_close_buttons()` call in `_apply_mode_visibility`.
- **Create** `tests/test_case_index_store.py`, `tests/test_separate_index_persistence.py`, `tests/test_task_card_action.py`.

---

## Task 1: Shared per-case index store

**Files:**
- Create: `icharlotte_core/case_index_store.py`
- Test: `tests/test_case_index_store.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_case_index_store.py`:

```python
"""Shared per-case document-index store: round-trip + format compatibility."""
import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import sys
import json
import shutil
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from icharlotte_core import config
from icharlotte_core import case_index_store as store


class TestCaseIndexStore(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self._orig = config.GEMINI_DATA_DIR
        config.GEMINI_DATA_DIR = self.tmp

    def tearDown(self):
        config.GEMINI_DATA_DIR = self._orig
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_load_missing_returns_empty(self):
        self.assertEqual(store.load_index("1234"), {})

    def test_upsert_roundtrip(self):
        docs = [{"id": "1", "title": "A", "date": "01/02/2020", "start": 1, "end": 3}]
        store.upsert_pdf("1234", r"C:\x\a.pdf", docs)
        self.assertEqual(store.load_index("1234"), {r"C:\x\a.pdf": docs})

    def test_upsert_second_pdf_preserves_first(self):
        store.upsert_pdf("1234", "p1", [{"id": "1"}])
        store.upsert_pdf("1234", "p2", [{"id": "2"}])
        idx = store.load_index("1234")
        self.assertEqual(set(idx), {"p1", "p2"})

    def test_upsert_same_pdf_overwrites(self):
        store.upsert_pdf("1234", "p1", [{"id": "1"}])
        store.upsert_pdf("1234", "p1", [{"id": "9"}])
        self.assertEqual(store.load_index("1234"), {"p1": [{"id": "9"}]})

    def test_corrupt_returns_empty(self):
        os.makedirs(self.tmp, exist_ok=True)
        with open(store.index_path("1234"), "w", encoding="utf-8") as f:
            f.write("{not valid json")
        self.assertEqual(store.load_index("1234"), {})

    def test_on_disk_shape_is_dict_of_lists(self):
        store.upsert_pdf("1234", "p1", [{"id": "1", "title": "A"}])
        with open(store.index_path("1234"), "r", encoding="utf-8") as f:
            raw = json.load(f)
        self.assertIsInstance(raw, dict)
        self.assertIn("p1", raw)
        self.assertIsInstance(raw["p1"], list)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_case_index_store.py -q`
Expected: FAIL — `ModuleNotFoundError: No module named 'icharlotte_core.case_index_store'`.

- [ ] **Step 3: Write the minimal implementation**

Create `icharlotte_core/case_index_store.py`:

```python
"""Shared per-case document-index store.

Both Advanced Mode's IndexTab and Wizard Mode's Separate task read/write one
JSON per case so a run in either mode is visible in both.

File:  <GEMINI_DATA_DIR>/<file_number>_index.json
Shape: {pdf_path: [ {id, title, date, start, end}, ... ]}

GEMINI_DATA_DIR is read from icharlotte_core.config at call time (attribute
access, not a frozen import binding) so tests can monkeypatch it.
"""
import json
import os

from icharlotte_core import config


def index_path(file_number: str) -> str:
    """Absolute path of the per-case index JSON."""
    return os.path.join(config.GEMINI_DATA_DIR, f"{file_number}_index.json")


def load_index(file_number: str) -> dict:
    """Return the {pdf_path: [docs]} map, or {} if missing/corrupt."""
    path = index_path(file_number)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def upsert_pdf(file_number: str, pdf_path: str, docs: list) -> None:
    """Set data[pdf_path] = docs and persist (matches IndexTab's format)."""
    data = load_index(file_number)
    data[pdf_path] = docs
    os.makedirs(config.GEMINI_DATA_DIR, exist_ok=True)
    with open(index_path(file_number), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest tests/test_case_index_store.py -q`
Expected: PASS (6 tests).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/case_index_store.py tests/test_case_index_store.py
git commit -m "feat(separate): shared per-case document index store"
```

---

## Task 2: Wizard Separate task persists runs to the store

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/separate_page.py`
- Test: `tests/test_separate_index_persistence.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_separate_index_persistence.py`:

```python
"""Wizard Separate task persists its document map to the shared index store."""
import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import sys
import shutil
import tempfile
import unittest
from types import SimpleNamespace

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from icharlotte_core import config
from icharlotte_core import case_index_store as store
from icharlotte_core.ui.wizard.pages import separate_page


class TestDocsFromWorkbench(unittest.TestCase):
    def test_reads_each_row_in_order(self):
        rows = [{"id": "1", "title": "A"}, {"id": "2", "title": "B"}]
        wb = SimpleNamespace(
            doc_table=SimpleNamespace(rowCount=lambda: len(rows)),
            _get_doc_from_row=lambda r: rows[r],
        )
        self.assertEqual(separate_page._docs_from_workbench(wb), rows)

    def test_empty_table(self):
        wb = SimpleNamespace(
            doc_table=SimpleNamespace(rowCount=lambda: 0),
            _get_doc_from_row=lambda r: None,
        )
        self.assertEqual(separate_page._docs_from_workbench(wb), [])


class TestPersistToIndex(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self._orig = config.GEMINI_DATA_DIR
        config.GEMINI_DATA_DIR = self.tmp

    def tearDown(self):
        config.GEMINI_DATA_DIR = self._orig
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_writes_store(self):
        fake = SimpleNamespace(_file_number="777", _pdf_path=r"C:\x\a.pdf")
        docs = [{"id": "1", "title": "A"}]
        separate_page.SeparateTaskTab._persist_to_index(fake, docs)
        self.assertEqual(store.load_index("777"), {r"C:\x\a.pdf": docs})

    def test_no_file_number_writes_nothing(self):
        fake = SimpleNamespace(_file_number="", _pdf_path=r"C:\x\a.pdf")
        separate_page.SeparateTaskTab._persist_to_index(fake, [{"id": "1"}])
        self.assertEqual(os.listdir(self.tmp), [])


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_separate_index_persistence.py -q`
Expected: FAIL — `AttributeError: module ... has no attribute '_docs_from_workbench'` and `type object 'SeparateTaskTab' has no attribute '_persist_to_index'`.

- [ ] **Step 3: Add the import and the two helpers**

In `icharlotte_core/ui/wizard/pages/separate_page.py`, add to the imports block (after `from icharlotte_core.config import SCRIPTS_DIR`):

```python
from icharlotte_core import case_index_store
```

Add a module-level function (place it just above `PAGE_SETTINGS = 0`):

```python
def _docs_from_workbench(workbench) -> list:
    """Read the current (possibly edited) document rows from a SeparatorWorkbench."""
    table = workbench.doc_table
    return [workbench._get_doc_from_row(row) for row in range(table.rowCount())]
```

Add a method to `SeparateTaskTab` (place it just above `closeEvent`):

```python
    def _persist_to_index(self, docs: list) -> None:
        """Persist this run's document map to the shared per-case index store so
        it is viewable later (Advanced Index tab / wizard reveal). Never fatal."""
        if not self._file_number:
            return
        try:
            case_index_store.upsert_pdf(self._file_number, self._pdf_path, docs)
        except Exception as e:  # persistence must never break the task flow
            try:
                from icharlotte_core.utils import log_event
                log_event(
                    f"[separate] failed to persist index for {self._pdf_path}: {e}",
                    "error",
                )
            except Exception:
                pass
```

- [ ] **Step 4: Wire the helpers into the two completion points**

In `_on_analysis_finished`, the success path currently ends:

```python
        docs = payload if isinstance(payload, list) else []
        self.workbench.set_busy(False)
        self.workbench.load_docs(self._pdf_path, docs)
        self.setCurrentIndex(PAGE_WORKBENCH)
```

Append one line so it becomes:

```python
        docs = payload if isinstance(payload, list) else []
        self.workbench.set_busy(False)
        self.workbench.load_docs(self._pdf_path, docs)
        self.setCurrentIndex(PAGE_WORKBENCH)
        self._persist_to_index(docs)
```

In `_on_processing_complete`, add persistence of the edited table as the first action (so the store reflects edits made before splitting):

```python
    def _on_processing_complete(self, summary: dict):
        from datetime import datetime
        self._persist_to_index(_docs_from_workbench(self.workbench))
        self.task_completed.emit({
            "task_id": self._spec.task_id,
            "title": self._spec.title,
            "files": [self._pdf_path],
            "settings": self.settings_page.to_dict(),
            "output_path": summary.get("output_folder", ""),
            "completed_at": datetime.now().isoformat(timespec="seconds"),
        })
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `python -m pytest tests/test_separate_index_persistence.py -q`
Expected: PASS (4 tests).

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/separate_page.py tests/test_separate_index_persistence.py
git commit -m "feat(separate): wizard runs persist to the shared index store"
```

---

## Task 3: TaskSpec optional card-action fields

**Files:**
- Modify: `icharlotte_core/ui/wizard/registry.py:58-87` (dataclass), `:174-182` (separate spec)
- Test: `tests/test_task_card_action.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_task_card_action.py`:

```python
"""Launcher card corner-action: spec fields, TaskCard button, WizardTab re-emit."""
import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from icharlotte_core.ui.wizard.registry import get_task


class TestCardActionSpec(unittest.TestCase):
    def test_separate_spec_declares_index_action(self):
        spec = get_task("separate")
        self.assertEqual(spec.card_action_id, "open_separate_index")
        self.assertTrue(spec.card_action_glyph)
        self.assertTrue(spec.card_action_tooltip)

    def test_other_spec_has_no_action(self):
        self.assertIsNone(get_task("chat").card_action_id)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_task_card_action.py -q`
Expected: FAIL — `AttributeError: 'TaskSpec' object has no attribute 'card_action_id'`.

- [ ] **Step 3: Add the dataclass fields**

In `icharlotte_core/ui/wizard/registry.py`, in `TaskSpec`, add three fields immediately after `phase2_flag: str = "--phase=summary"` (line 69) and before the `_settings_page_cls_factory` line:

```python
    # Optional launcher-card corner button (e.g. Separate → open the Index).
    # When card_action_id is set, TaskCard renders a small QToolButton that
    # emits action_requested(card_action_id). None = no button (default).
    card_action_id: Optional[str] = None
    card_action_glyph: Optional[str] = None
    card_action_tooltip: Optional[str] = None
```

- [ ] **Step 4: Set the fields on the separate spec**

Replace the `separate` spec (lines 174-182) with:

```python
    "separate": TaskSpec(
        task_id="separate",
        title="Separate Documents",
        description="Split a combined PDF into individually-named documents using AI.",
        icon_glyph="\U0001F4D1",  # 📑
        script_name="separate.py",
        default_folders=[],
        category="General",
        card_action_id="open_separate_index",
        card_action_glyph="\U0001F5C2",  # 🗂 card index dividers
        card_action_tooltip="Open the document Index for this case",
    ),
```

- [ ] **Step 5: Run the tests to verify they pass (and no registry regression)**

Run: `python -m pytest tests/test_task_card_action.py tests/test_registry.py -q`
Expected: PASS (the 2 new tests; `test_registry.py::test_initial_tasks_registered` still passes — no task was added, only fields).

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/ui/wizard/registry.py tests/test_task_card_action.py
git commit -m "feat(wizard): optional card-action fields on TaskSpec; set on separate"
```

---

## Task 4: TaskCard corner button

**Files:**
- Modify: `icharlotte_core/ui/wizard/task_card.py`
- Test: `tests/test_task_card_action.py` (extend)

- [ ] **Step 1: Write the failing tests (append to the file)**

Append to `tests/test_task_card_action.py` (before the `if __name__` line):

```python
from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication

from icharlotte_core.ui.wizard.task_card import TaskCard


def _app():
    return QApplication.instance() or QApplication([])


class TestTaskCardButton(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = _app()

    def test_separate_card_has_action_button(self):
        card = TaskCard(get_task("separate"))
        self.assertIsNotNone(card.action_btn)

    def test_other_card_has_no_action_button(self):
        card = TaskCard(get_task("chat"))
        self.assertIsNone(card.action_btn)

    def test_button_emits_action_requested(self):
        card = TaskCard(get_task("separate"))
        seen = []
        card.action_requested.connect(seen.append)
        card.action_btn.click()
        self.assertEqual(seen, ["open_separate_index"])

    def test_button_click_does_not_launch_task(self):
        card = TaskCard(get_task("separate"))
        card.resize(280, 140)
        card.show()
        launched = []
        card.clicked.connect(launched.append)
        self.app.processEvents()
        QTest.mouseClick(card.action_btn, Qt.MouseButton.LeftButton)
        self.app.processEvents()
        self.assertEqual(launched, [], "Corner button must not trigger the card's launch click")
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_task_card_action.py::TestTaskCardButton -q`
Expected: FAIL — `AttributeError: 'TaskCard' object has no attribute 'action_btn'`.

- [ ] **Step 3: Implement the corner button**

In `icharlotte_core/ui/wizard/task_card.py`, update the import line:

```python
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QToolButton, QVBoxLayout
```

Add the signal to `TaskCard` (next to `clicked`):

```python
    clicked = Signal(str)           # task_id
    action_requested = Signal(str)  # card_action_id (corner button)
```

At the end of `__init__` (after `outer.addWidget(self.description_label, 1)`), add:

```python
        self.action_btn = None
        if spec.card_action_id:
            footer = QHBoxLayout()
            footer.setContentsMargins(0, 0, 0, 0)
            footer.addStretch()
            self.action_btn = QToolButton()
            self.action_btn.setText(spec.card_action_glyph or "⋯")  # ⋯ fallback
            self.action_btn.setToolTip(spec.card_action_tooltip or "")
            self.action_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            self.action_btn.setAutoRaise(True)
            self.action_btn.setStyleSheet(
                "QToolButton { border: none; font-size: 16px; padding: 2px; }"
                f" QToolButton:hover {{ background-color: {theme.BG_SUBTLE};"
                f" border-radius: {theme.RADIUS_MD}px; }}"
            )
            self.action_btn.clicked.connect(
                lambda _=False: self.action_requested.emit(self._spec.card_action_id)
            )
            footer.addWidget(self.action_btn)
            outer.addLayout(footer)
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest tests/test_task_card_action.py -q`
Expected: PASS (all card-action tests).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/task_card.py tests/test_task_card_action.py
git commit -m "feat(wizard): TaskCard renders a corner action button"
```

---

## Task 5: WizardTab re-emits the card action

**Files:**
- Modify: `icharlotte_core/ui/wizard/wizard_tab.py:25-27` (signals), `:92-95` (card wiring)
- Test: `tests/test_task_card_action.py` (extend)

- [ ] **Step 1: Write the failing test (append to the file)**

Append to `tests/test_task_card_action.py` (before the `if __name__` line):

```python
from icharlotte_core.ui.wizard.wizard_tab import WizardTab


class TestWizardTabReemit(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = _app()

    def test_reemits_card_action(self):
        w = WizardTab()
        seen = []
        w.card_action_requested.connect(seen.append)
        w.cards["separate"].action_requested.emit("open_separate_index")
        self.assertEqual(seen, ["open_separate_index"])
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `python -m pytest tests/test_task_card_action.py::TestWizardTabReemit -q`
Expected: FAIL — `AttributeError: 'WizardTab' object has no attribute 'card_action_requested'`.

- [ ] **Step 3: Add the signal and wire each card**

In `icharlotte_core/ui/wizard/wizard_tab.py`, add the signal to `WizardTab` (next to the existing ones, lines 26-27):

```python
    task_requested = Signal(str)            # task_id
    reopen_requested = Signal(dict)         # recent-tasks entry
    card_action_requested = Signal(str)     # card_action_id (corner button)
```

In `_build_ui`, the card-creation loop currently reads:

```python
        for spec in list_tasks():
            card = TaskCard(spec)
            card.clicked.connect(self.task_requested.emit)
            self.cards[spec.task_id] = card
```

Add the action re-emit:

```python
        for spec in list_tasks():
            card = TaskCard(spec)
            card.clicked.connect(self.task_requested.emit)
            card.action_requested.connect(self.card_action_requested.emit)
            self.cards[spec.task_id] = card
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `python -m pytest tests/test_task_card_action.py -q`
Expected: PASS (all card-action tests, including re-emit).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/wizard_tab.py tests/test_task_card_action.py
git commit -m "feat(wizard): WizardTab re-emits card_action_requested"
```

---

## Task 6: Main window reveal + re-hide wiring

**Files:**
- Modify: `iCharlotte.py` — `:657` (connect), new `_on_card_action`/`_reveal_index_tab` (near `:1774`), `_hide_fixed_close_buttons` (`:1805-1815`), `_on_tab_close_requested` (`:1775-1781`), `_apply_mode_visibility` (`:1034-1045`).

> No unit test: building the full `MainWindow` in-process risks the known startup-monitor hangs (docket/Outlook). The reveal/re-hide *logic* is covered indirectly by Tasks 1–5; this task is verified by an import-smoke check plus a live app run.

- [ ] **Step 1: Connect the launcher signal**

In `iCharlotte.py`, immediately after (line ~657):

```python
        self.wizard_tab.task_requested.connect(self._open_task_tab)
        self.wizard_tab.reopen_requested.connect(self._on_reopen_recent_task)
```

add:

```python
        self.wizard_tab.card_action_requested.connect(self._on_card_action)
```

- [ ] **Step 2: Add the dispatch + reveal methods**

Add these two methods right after `_open_task_tab` ends (before `_on_tab_close_requested`, ~line 1774):

```python
    def _on_card_action(self, action_id: str) -> None:
        """Dispatch a launcher-card corner-button action."""
        if action_id == "open_separate_index":
            self._reveal_index_tab()

    def _reveal_index_tab(self) -> None:
        """Wizard Mode: reveal the hidden Index singleton, reloaded from disk so
        it reflects the latest Separate runs (wizard or advanced)."""
        if not self.case_path:
            QMessageBox.information(
                self, "No case loaded",
                "Open a case from the Master List first.",
            )
            return
        idx = self._index_of_tab("Index")
        if idx < 0:
            return
        if self.file_number and hasattr(self, "index_tab"):
            self.index_tab.load_data(self.file_number)
        self.tabs.setTabVisible(idx, True)
        self.tabs.setCurrentIndex(idx)
        self._hide_fixed_close_buttons()
```

- [ ] **Step 3: Let the Index "x" re-hide it (close handler)**

In `_on_tab_close_requested`, after the `if widget is None: return` guard and before the `wizard_task_id` early-return, insert:

```python
        # Wizard Mode: the Index tab is the shared singleton — its "x" hides it,
        # never destroys it.
        if (
            widget is getattr(self, "index_tab", None)
            and getattr(self, "mode_controller", None) is not None
            and self.mode_controller.is_wizard
        ):
            self.tabs.setTabVisible(index, False)
            wiz = self._index_of_tab("Wizard")
            if wiz >= 0:
                self.tabs.setCurrentIndex(wiz)
            return
```

- [ ] **Step 4: Show the Index "x" only in wizard mode when visible**

Replace the body of `_hide_fixed_close_buttons` with:

```python
    def _hide_fixed_close_buttons(self) -> None:
        """Hide close buttons on non-TaskTabs.

        Exception: in Wizard Mode the revealed Index singleton gets a visible
        "x" that re-hides (not destroys) it — see _on_tab_close_requested.
        """
        from PySide6.QtWidgets import QTabBar
        bar = self.tabs.tabBar()
        index_tab = getattr(self, "index_tab", None)
        mc = getattr(self, "mode_controller", None)
        is_wizard = bool(mc is not None and mc.is_wizard)
        for i in range(self.tabs.count()):
            widget = self.tabs.widget(i)
            is_task_tab = widget is not None and widget.property("wizard_task_id") is not None
            is_rehideable_index = (
                index_tab is not None
                and widget is index_tab
                and is_wizard
                and self.tabs.isTabVisible(i)
            )
            show_close = is_task_tab or is_rehideable_index
            for side in (QTabBar.ButtonPosition.RightSide, QTabBar.ButtonPosition.LeftSide):
                btn = bar.tabButton(i, side)
                if btn is not None:
                    btn.setVisible(show_close)
```

- [ ] **Step 5: Refresh close buttons on mode change**

In `_apply_mode_visibility`, the method currently ends:

```python
        if not self.tabs.isTabVisible(self.tabs.currentIndex()):
            self.tabs.setCurrentIndex(0)
```

Append:

```python
        if not self.tabs.isTabVisible(self.tabs.currentIndex()):
            self.tabs.setCurrentIndex(0)
        self._hide_fixed_close_buttons()
```

- [ ] **Step 6: Import-smoke + regression tests**

Run: `python -c "import iCharlotte"`
Expected: no output, exit 0 (module imports; syntax/wiring OK).

Run: `python -m pytest tests/test_case_index_store.py tests/test_separate_index_persistence.py tests/test_task_card_action.py -q`
Expected: PASS (all feature tests still green).

- [ ] **Step 7: Live verification (from the MAIN checkout)**

After merging/applying this branch to `C:\geminiterminal2` and restarting iCharlotte:
1. Load a case (defaults to Wizard mode).
2. On the **Separate Documents** card, confirm a small 🗂 button sits at the bottom-right; hover shows "Open the document Index for this case".
3. Click the card body → the Separate task tab opens (button did **not** hijack the launch).
4. Run a separation (Analyze → Review). Close the task tab.
5. Click the 🗂 button on the card → the **Index** tab appears, current, listing the just-separated source PDF with its document rows.
6. Click the Index tab's **"x"** → it disappears and the Wizard tab is shown. Click 🗂 again → it reappears (no crash, data intact).
7. Toggle to Advanced mode → the Index tab is present and has **no** "x" (permanent there).

- [ ] **Step 8: Commit**

```bash
git add iCharlotte.py
git commit -m "feat(wizard): reveal/re-hide the Index tab from the Separate card"
```

---

## Self-Review

**Spec coverage:**
- Goal 1 (persist to shared store) → Tasks 1 + 2.
- Goal 2 (corner button opens Index) → Tasks 3 + 4 + 5 + 6 (steps 1-2).
- Goal 3 (re-hide via "x") → Task 6 (steps 3-5).
- Non-goal "don't refactor IndexTab" → honored (only `case_index_store` writes; IndexTab untouched).
- Testing plan (store, card, wizard, persistence, reveal) → Tasks 1-6 tests + Task 6 live steps.
- Edge cases: corrupt/missing store (Task 1), empty file_number (Task 2), no-case reveal guard (Task 6 step 2), button click isolation (Task 4).

**Placeholder scan:** none — every code step shows complete content; no TBD/TODO/"similar to".

**Type/name consistency (checked across tasks):** `case_index_store.{index_path,load_index,upsert_pdf}` (Tasks 1,2,6); `_docs_from_workbench` + `_persist_to_index` (Task 2); `card_action_id/glyph/tooltip` (Tasks 3,4); `action_requested` (Tasks 4,5); `card_action_requested` (Tasks 5,6); action id `"open_separate_index"` (Tasks 3,4,5,6); `_on_card_action`/`_reveal_index_tab` (Task 6). All consistent.
