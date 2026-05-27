# Wizard Mode Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a guided "Wizard Mode" alongside the existing UI (renamed Advanced Mode). Wizard Mode replaces the multi-tab Case View flow with a card-based task picker and per-task tabs that walk the user through file selection → settings → status → an editable output page.

**Architecture:** A new `icharlotte_core/ui/wizard/` package introduces a `ModeController` (global, QSettings-backed), a `WizardTab` (card grid), and `TaskTab` (`QStackedWidget` with Settings/Status/Output pages). Existing tabs and agents are untouched — `MainWindow` orchestrates which tabs are visible per mode, and task runners are thin `QProcess` wrappers around the existing `Scripts/*.py` agents. Per-case state (open task tabs, recent tasks history) is persisted to `<case>/.icharlotte/wizard_state.json` with atomic writes.

**Tech Stack:** PySide6 (QTabWidget, QStackedWidget, QGridLayout, QFileDialog, QProcess), python-docx, **mammoth** (new dependency for .docx → HTML), QSettings.

**Spec:** `docs/superpowers/specs/2026-05-17-wizard-mode-design.md`

---

## Conventions used throughout this plan

- All file paths are absolute under `C:\geminiterminal2\`. Tasks use POSIX-style paths in code (Python handles both).
- **Test command:** `pytest tests/path/test_file.py::TestClass::test_method -v` (run from repo root). When pytest-qt is required, the test should `import pytestqt` or use the `qtbot` fixture; CI infra runs Windows so headless Qt works.
- **App import sanity check** after structural changes: `python -c "import iCharlotte"` — must exit 0.
- **Commit style:** Conventional commits (`feat(scope): …`, `fix(scope): …`, `test(scope): …`, `refactor(scope): …`, `chore(scope): …`).
- **Co-Authored-By:** Append `Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>` on every commit.
- **TDD discipline:** for non-UI modules (controller, persistence, registry, file picker, runners), write the failing test first, then the implementation. UI widgets get smoke tests with `qtbot` but lean on manual verification too.
- **Existing Qt module** in this codebase: `PySide6` (NOT `PyQt6`). All imports must use `PySide6.*`.

---

# Phase 1 — Mode controller + tab orchestration

This phase adds the dual-mode foundation. After Phase 1, the app still works in Advanced mode (default behavior unchanged), but the new toggle lives in the Master List tab and switches a `ModeController` whose only side effect (for now) is logging — tab visibility wires up at the end of the phase.

### Task 1.1: Add `mammoth` to requirements.txt

**Files:**
- Modify: `requirements.txt`

- [ ] **Step 1: Add the dependency**

Open `C:/geminiterminal2/requirements.txt` and append after the `python-docx>=1.0.0` line:

```
mammoth>=1.6.0  # .docx → HTML for Wizard Mode output editor
```

- [ ] **Step 2: Install it**

Run from repo root:

```bash
python -m pip install "mammoth>=1.6.0"
```

Expected: install succeeds (or "Requirement already satisfied"). Verify:

```bash
python -c "import mammoth; print(mammoth.__version__)"
```

Expected: prints a version like `1.6.0`. Non-zero exit = stop and debug pip.

- [ ] **Step 3: Commit**

```bash
git add requirements.txt
git commit -m "$(cat <<'EOF'
chore(deps): add mammoth for wizard output editor

mammoth converts .docx → semantic HTML for the Wizard Mode output page
editor (rich in-app editor with round-trip save back to .docx).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 1.2: Create `ModeController` (TDD)

**Files:**
- Create: `icharlotte_core/ui/wizard/__init__.py`
- Create: `icharlotte_core/ui/wizard/mode_controller.py`
- Create: `tests/test_wizard/__init__.py`
- Create: `tests/test_wizard/test_mode_controller.py`

- [ ] **Step 1: Create empty package files**

```bash
mkdir -p icharlotte_core/ui/wizard
mkdir -p tests/test_wizard
```

Create `icharlotte_core/ui/wizard/__init__.py` with content:

```python
"""Wizard Mode UI package."""
```

Create `tests/test_wizard/__init__.py` with content: *(empty file)*

- [ ] **Step 2: Write failing tests**

Create `tests/test_wizard/test_mode_controller.py`:

```python
"""Tests for ModeController — global Advanced/Wizard mode persistence."""
import pytest

pytest.importorskip("pytestqt")
from PySide6.QtCore import QSettings, QCoreApplication

from icharlotte_core.ui.wizard.mode_controller import ModeController, MODE_ADVANCED, MODE_WIZARD


@pytest.fixture(autouse=True)
def _qsettings_org(monkeypatch, tmp_path):
    """Isolate QSettings to a temp file per test."""
    QCoreApplication.setOrganizationName("iCharlotteTest")
    QCoreApplication.setApplicationName(f"WizardTest-{tmp_path.name}")
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    s = QSettings()
    s.clear()
    s.sync()
    yield
    s.clear()
    s.sync()


def test_first_run_defaults_to_wizard():
    ctrl = ModeController()
    assert ctrl.mode == MODE_WIZARD


def test_set_mode_persists_via_qsettings():
    ctrl1 = ModeController()
    ctrl1.set_mode(MODE_ADVANCED)
    # New instance must read the persisted value.
    ctrl2 = ModeController()
    assert ctrl2.mode == MODE_ADVANCED


def test_set_mode_emits_signal(qtbot):
    ctrl = ModeController()
    with qtbot.waitSignal(ctrl.mode_changed, timeout=500) as blocker:
        ctrl.set_mode(MODE_ADVANCED)
    assert blocker.args == [MODE_ADVANCED]


def test_set_mode_does_not_emit_when_unchanged(qtbot):
    ctrl = ModeController()
    ctrl.set_mode(MODE_ADVANCED)  # change once
    # Now setting to the same value should NOT fire the signal.
    with qtbot.assertNotEmitted(ctrl.mode_changed, wait=200):
        ctrl.set_mode(MODE_ADVANCED)


def test_invalid_mode_raises():
    ctrl = ModeController()
    with pytest.raises(ValueError):
        ctrl.set_mode("nonsense")


def test_is_wizard_helper():
    ctrl = ModeController()
    ctrl.set_mode(MODE_WIZARD)
    assert ctrl.is_wizard is True
    ctrl.set_mode(MODE_ADVANCED)
    assert ctrl.is_wizard is False
```

- [ ] **Step 3: Run the tests to confirm they fail**

```bash
pytest tests/test_wizard/test_mode_controller.py -v
```

Expected: `ImportError: cannot import name 'ModeController' from 'icharlotte_core.ui.wizard.mode_controller'` or the module doesn't exist. **Do not proceed if tests pass accidentally.**

- [ ] **Step 4: Implement `ModeController`**

Create `icharlotte_core/ui/wizard/mode_controller.py`:

```python
"""ModeController — coordinates Advanced vs Wizard mode app-wide.

Mode is stored globally in QSettings (one toggle for the whole app, not
per-case). On change, emits `mode_changed(str)` so MainWindow can update
tab visibility.
"""
from PySide6.QtCore import QObject, QSettings, Signal


MODE_ADVANCED = "advanced"
MODE_WIZARD = "wizard"
_VALID_MODES = {MODE_ADVANCED, MODE_WIZARD}

_SETTINGS_KEY = "app/mode"
_DEFAULT_MODE = MODE_WIZARD


class ModeController(QObject):
    """Global mode coordinator. Backed by QSettings."""

    mode_changed = Signal(str)  # emits new mode value

    def __init__(self, parent: QObject | None = None):
        super().__init__(parent)
        self._settings = QSettings("iCharlotte", "iCharlotte")
        self._mode = self._read_mode()

    def _read_mode(self) -> str:
        value = self._settings.value(_SETTINGS_KEY, _DEFAULT_MODE)
        if value not in _VALID_MODES:
            return _DEFAULT_MODE
        return value

    @property
    def mode(self) -> str:
        return self._mode

    @property
    def is_wizard(self) -> bool:
        return self._mode == MODE_WIZARD

    @property
    def is_advanced(self) -> bool:
        return self._mode == MODE_ADVANCED

    def set_mode(self, mode: str) -> None:
        if mode not in _VALID_MODES:
            raise ValueError(f"Invalid mode: {mode!r}. Must be one of {_VALID_MODES}.")
        if mode == self._mode:
            return
        self._mode = mode
        self._settings.setValue(_SETTINGS_KEY, mode)
        self._settings.sync()
        self.mode_changed.emit(mode)
```

- [ ] **Step 5: Run the tests to confirm they pass**

```bash
pytest tests/test_wizard/test_mode_controller.py -v
```

Expected: all 6 tests pass.

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/ui/wizard/__init__.py icharlotte_core/ui/wizard/mode_controller.py tests/test_wizard/__init__.py tests/test_wizard/test_mode_controller.py
git commit -m "$(cat <<'EOF'
feat(wizard): ModeController for Advanced/Wizard mode

Global QSettings-backed mode controller. Persists across sessions,
defaults to wizard on first run, emits mode_changed only when value
actually changes.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 1.3: Create `ModeToggle` segmented control widget

**Files:**
- Create: `icharlotte_core/ui/wizard/mode_toggle.py`
- Create: `tests/test_wizard/test_mode_toggle.py`

- [ ] **Step 1: Write the smoke test**

Create `tests/test_wizard/test_mode_toggle.py`:

```python
"""Smoke tests for ModeToggle segmented control."""
import pytest

pytest.importorskip("pytestqt")
from PySide6.QtCore import QCoreApplication, QSettings, Qt

from icharlotte_core.ui.wizard.mode_controller import ModeController, MODE_ADVANCED, MODE_WIZARD
from icharlotte_core.ui.wizard.mode_toggle import ModeToggle


@pytest.fixture(autouse=True)
def _qsettings_org(tmp_path):
    QCoreApplication.setOrganizationName("iCharlotteTest")
    QCoreApplication.setApplicationName(f"WizardTest-{tmp_path.name}")
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    s = QSettings()
    s.clear()
    s.sync()
    yield
    s.clear()
    s.sync()


def test_toggle_reflects_initial_mode(qtbot):
    ctrl = ModeController()
    ctrl.set_mode(MODE_WIZARD)
    w = ModeToggle(ctrl)
    qtbot.addWidget(w)
    assert w.wizard_button.isChecked() is True
    assert w.advanced_button.isChecked() is False


def test_clicking_advanced_updates_controller(qtbot):
    ctrl = ModeController()
    ctrl.set_mode(MODE_WIZARD)
    w = ModeToggle(ctrl)
    qtbot.addWidget(w)
    qtbot.mouseClick(w.advanced_button, Qt.MouseButton.LeftButton)
    assert ctrl.mode == MODE_ADVANCED


def test_external_mode_change_updates_buttons(qtbot):
    ctrl = ModeController()
    ctrl.set_mode(MODE_WIZARD)
    w = ModeToggle(ctrl)
    qtbot.addWidget(w)
    ctrl.set_mode(MODE_ADVANCED)
    assert w.advanced_button.isChecked() is True
    assert w.wizard_button.isChecked() is False
```

- [ ] **Step 2: Run the tests — confirm failure**

```bash
pytest tests/test_wizard/test_mode_toggle.py -v
```

Expected: `ModuleNotFoundError: No module named 'icharlotte_core.ui.wizard.mode_toggle'`.

- [ ] **Step 3: Implement `ModeToggle`**

Create `icharlotte_core/ui/wizard/mode_toggle.py`:

```python
"""ModeToggle — segmented control for Advanced/Wizard mode selection."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QButtonGroup, QHBoxLayout, QPushButton, QWidget

from .mode_controller import ModeController, MODE_ADVANCED, MODE_WIZARD


_SEGMENTED_STYLE = """
QPushButton {
    background-color: #f5f5f5;
    color: #555;
    font-size: 12px;
    font-weight: 500;
    padding: 6px 16px;
    border: 1px solid #ccc;
    min-width: 110px;
    height: 28px;
}
QPushButton:checked {
    background-color: #1976D2;
    color: white;
    font-weight: 600;
    border-color: #0D47A1;
}
QPushButton:hover:!checked {
    background-color: #e8e8e8;
}
"""


class ModeToggle(QWidget):
    """Two-button segmented control bound to a ModeController.

    Listens to the controller's mode_changed signal so external mode
    changes (e.g. keyboard shortcut, programmatic) keep the UI in sync.
    """

    def __init__(self, controller: ModeController, parent: QWidget | None = None):
        super().__init__(parent)
        self._controller = controller

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.advanced_button = QPushButton("Advanced Mode")
        self.advanced_button.setCheckable(True)
        self.advanced_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.advanced_button.setStyleSheet(
            _SEGMENTED_STYLE + "QPushButton { border-top-left-radius: 4px; border-bottom-left-radius: 4px; border-right: none; }"
        )

        self.wizard_button = QPushButton("Wizard Mode")
        self.wizard_button.setCheckable(True)
        self.wizard_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.wizard_button.setStyleSheet(
            _SEGMENTED_STYLE + "QPushButton { border-top-right-radius: 4px; border-bottom-right-radius: 4px; }"
        )

        layout.addWidget(self.advanced_button)
        layout.addWidget(self.wizard_button)

        self._group = QButtonGroup(self)
        self._group.setExclusive(True)
        self._group.addButton(self.advanced_button)
        self._group.addButton(self.wizard_button)

        self.advanced_button.clicked.connect(lambda: self._controller.set_mode(MODE_ADVANCED))
        self.wizard_button.clicked.connect(lambda: self._controller.set_mode(MODE_WIZARD))

        self._controller.mode_changed.connect(self._sync_from_controller)
        self._sync_from_controller(self._controller.mode)

    def _sync_from_controller(self, mode: str) -> None:
        self.advanced_button.setChecked(mode == MODE_ADVANCED)
        self.wizard_button.setChecked(mode == MODE_WIZARD)
```

- [ ] **Step 4: Run the tests — confirm pass**

```bash
pytest tests/test_wizard/test_mode_toggle.py -v
```

Expected: 3 tests pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/mode_toggle.py tests/test_wizard/test_mode_toggle.py
git commit -m "$(cat <<'EOF'
feat(wizard): ModeToggle segmented control widget

Two-button group bound to ModeController. Reflects external mode
changes via mode_changed signal. Used inside Master List tab content.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 1.4: Remove the corner-widget "Change File" button

**Files:**
- Modify: `iCharlotte.py:922-929` (button creation) and the Win+C hotkey/registration messages

- [ ] **Step 1: Delete the button**

Open `iCharlotte.py`. Find lines 922-929 (currently `self.btn_change_file = QPushButton("Change File")` block). Replace **only the Change File block** (8 lines) with a comment placeholder:

```python
        # Change File button removed in favor of Master List mode toggle (Wizard).
        # Win+C hotkey is preserved for power users.
```

- [ ] **Step 2: Verify Win+C still works structurally**

Search file:

```bash
grep -n "win+c\|change_file_signal\|_on_change_file_hotkey\|def change_file" iCharlotte.py
```

Expected: the hotkey at `iCharlotte.py:464`, the signal at line 380, the `_on_change_file_hotkey` at line 526, and the `change_file()` method at line 1293 — all still present. The only removed thing is the button.

- [ ] **Step 3: Smoke-test the app imports**

```bash
python -c "import iCharlotte; print('OK')"
```

Expected: prints `OK`. If it fails, undo and inspect the error (likely a stray reference to `self.btn_change_file` elsewhere — find it with `grep -n btn_change_file iCharlotte.py` and remove).

- [ ] **Step 4: Commit**

```bash
git add iCharlotte.py
git commit -m "$(cat <<'EOF'
refactor(ui): remove Change File corner button

Replaced by the mode toggle inside the Master List tab content (next
task). Win+C global hotkey continues to summon the Change File dialog.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 1.5: Add mode toggle to Master List tab content

**Files:**
- Modify: `icharlotte_core/ui/master_case_tab.py` — add toggle to the header row
- Modify: `iCharlotte.py` — instantiate `ModeController`, pass to `MasterCaseTab`

- [ ] **Step 1: Add the toggle to `MasterCaseTab`**

Open `icharlotte_core/ui/master_case_tab.py`. Find the `__init__` method of `MasterCaseTab`. Locate the top-most header layout (where the case-list title or filter row is created — search for `setLayout` or the first `QHBoxLayout` inside `__init__`). Add the mode toggle to that header row, right-aligned.

Add this import at the top of the file (after the existing `from PySide6...` block):

```python
from .wizard.mode_toggle import ModeToggle
from .wizard.mode_controller import ModeController
```

In `MasterCaseTab.__init__`, modify the signature to accept a controller:

```python
def __init__(self, parent=None, mode_controller: ModeController | None = None):
    super().__init__(parent)
    self.main_window = parent
    self.mode_controller = mode_controller
    ...
```

Then, near the top of `__init__` where the header is built (find the line that creates the first child layout under `self.main_layout` or the equivalent), add:

```python
# Mode toggle — right-aligned in the header row.
if self.mode_controller is not None:
    self.mode_toggle = ModeToggle(self.mode_controller, parent=self)
    header_layout.addStretch()           # push toggle to the right
    header_layout.addWidget(self.mode_toggle)
```

If there is no obvious "header_layout", create one above the table:

```python
header_layout = QHBoxLayout()
header_layout.addStretch()
if self.mode_controller is not None:
    self.mode_toggle = ModeToggle(self.mode_controller, parent=self)
    header_layout.addWidget(self.mode_toggle)
# Insert header_layout as the first child of the tab's main layout.
self.layout().insertLayout(0, header_layout)
```

- [ ] **Step 2: Wire up `ModeController` in `MainWindow`**

Open `iCharlotte.py`. In `MainWindow.__init__`, before `self.setup_ui()` (around line 441), add:

```python
        # Wizard mode controller (global Advanced/Wizard toggle).
        from icharlotte_core.ui.wizard.mode_controller import ModeController
        self.mode_controller = ModeController(parent=self)
```

In `setup_ui()`, where `self.master_tab = MasterCaseTab(self)` is created (around line 597), change to:

```python
        self.master_tab = MasterCaseTab(self, mode_controller=self.mode_controller)
```

- [ ] **Step 3: Smoke test — app starts in both modes**

```bash
python -c "import iCharlotte; print('OK')"
```

Expected: `OK`.

Now run the app manually:

```bash
python iCharlotte.py
```

Expected: app launches. Master List tab is visible. A segmented "Advanced Mode | Wizard Mode" control appears in the top-right of the Master List tab. Wizard Mode is selected on first run. Clicking switches the selection (no other visible effect yet — that comes in Task 1.6).

- [ ] **Step 4: Commit**

```bash
git add iCharlotte.py icharlotte_core/ui/master_case_tab.py
git commit -m "$(cat <<'EOF'
feat(wizard): wire ModeController + add toggle to Master List

ModeController instantiated in MainWindow.__init__ and passed to
MasterCaseTab, which renders a ModeToggle in its header row. No other
visible effect yet — tab visibility wiring follows in next task.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 1.6: Drive tab visibility from `ModeController`

**Files:**
- Modify: `iCharlotte.py` — add `_apply_mode_visibility()`, connect to signal, call on startup

- [ ] **Step 1: Add the visibility method**

In `iCharlotte.py`, inside `MainWindow`, add a new method (place it after `setup_ui`):

```python
    # --- Wizard Mode: tab visibility orchestration ---

    # Names of tabs to HIDE when in Wizard Mode.
    # Master List is always visible. The Wizard tab and any task tabs are
    # added/managed separately in Phase 2+.
    _WIZARD_HIDDEN_TABS = {
        "Case View",
        "Status",
        "Index",
        "Chat",
        "Email",
        "Email Update",
        "Depositions",
        "Discovery",
        "Liability & Exposure",
        "Templates / Resources",
        "Logs",
    }

    def _apply_mode_visibility(self, mode: str) -> None:
        """Show/hide tabs based on current mode."""
        is_wizard = (mode == "wizard")
        for i in range(self.tabs.count()):
            tab_text = self.tabs.tabText(i)
            if tab_text in self._WIZARD_HIDDEN_TABS:
                self.tabs.setTabVisible(i, not is_wizard)
            # Master List (and future Wizard tab + task tabs) stay visible
            # in both modes; they're handled by their own logic.
        # If the current tab just got hidden, fall back to Master List.
        if not self.tabs.isTabVisible(self.tabs.currentIndex()):
            self.tabs.setCurrentIndex(0)
```

- [ ] **Step 2: Connect the signal**

In `MainWindow.__init__`, after `self.setup_ui()`, add:

```python
        # Apply current mode and react to future mode changes.
        self.mode_controller.mode_changed.connect(self._apply_mode_visibility)
        self._apply_mode_visibility(self.mode_controller.mode)
```

- [ ] **Step 3: Manual verification**

```bash
python iCharlotte.py
```

Expected sequence:
1. App opens. Default mode is Wizard. Only **Master List** is visible (Case View, Status, Index, Chat, Email, Email Update, Depositions, Discovery, Liability & Exposure, Templates / Resources, Logs are all hidden).
2. Click **Advanced Mode** in the toggle. All hidden tabs reappear instantly.
3. Click **Wizard Mode** again. They disappear.
4. Close and re-open the app. Whatever mode was last selected is the one shown.

If any tab fails to hide/show, check the tab text matches the `_WIZARD_HIDDEN_TABS` set exactly (case + spacing sensitive).

- [ ] **Step 4: Commit**

```bash
git add iCharlotte.py
git commit -m "$(cat <<'EOF'
feat(wizard): hide Advanced-Mode tabs when Wizard Mode is active

MainWindow now applies tab visibility from ModeController on startup
and on every mode change. Master List stays visible in both modes;
the Wizard tab + task tabs are added in subsequent phases.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

# Phase 2 — Wizard tab + task registry + cards

After Phase 2, switching to Wizard Mode shows a "What would you like to do?" tab with four clickable task cards. Clicks log to console for now; tab creation happens in Phase 4.

### Task 2.1: Create `TASK_REGISTRY`

**Files:**
- Create: `icharlotte_core/ui/wizard/registry.py`
- Create: `tests/test_wizard/test_registry.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_wizard/test_registry.py`:

```python
"""Tests for the task registry."""
import pytest

from icharlotte_core.ui.wizard.registry import (
    TASK_REGISTRY,
    TaskSpec,
    get_task,
    list_tasks,
)


def test_four_initial_tasks_registered():
    ids = {t.task_id for t in list_tasks()}
    assert ids == {
        "summarize_documents",
        "summarize_discovery",
        "summarize_depositions",
        "medical_records",
    }


def test_each_task_has_required_metadata():
    for spec in list_tasks():
        assert isinstance(spec, TaskSpec)
        assert spec.task_id
        assert spec.title
        assert spec.description
        assert isinstance(spec.default_folders, list)


def test_default_folders_per_task():
    assert get_task("summarize_documents").default_folders == []
    assert get_task("summarize_discovery").default_folders == ["DISCOVERY/RESPONSES", "DISCOVERY"]
    assert get_task("summarize_depositions").default_folders == ["DISCOVERY/TRANSCRIPTS", "DISCOVERY"]
    assert get_task("medical_records").default_folders == ["RECORDS"]


def test_get_task_unknown_raises():
    with pytest.raises(KeyError):
        get_task("not_a_real_task")
```

- [ ] **Step 2: Run — confirm failure**

```bash
pytest tests/test_wizard/test_registry.py -v
```

Expected: `ModuleNotFoundError`.

- [ ] **Step 3: Implement the registry**

Create `icharlotte_core/ui/wizard/registry.py`:

```python
"""Task registry — single source of truth for Wizard Mode task cards.

Each task contributes:
  - task_id            : stable identifier used in persistence and code paths
  - title              : human-readable card title
  - description        : one-line card description
  - icon_glyph         : single emoji-ish character used as the card icon
  - default_folders    : ordered list of relative subfolders (under case root)
                         tried in order when opening the pre-Settings file
                         dialog; empty list means default to case root.
"""
from dataclasses import dataclass, field
from typing import List


@dataclass(frozen=True)
class TaskSpec:
    task_id: str
    title: str
    description: str
    icon_glyph: str
    default_folders: List[str] = field(default_factory=list)


TASK_REGISTRY: dict[str, TaskSpec] = {
    "summarize_documents": TaskSpec(
        task_id="summarize_documents",
        title="Summarize Documents",
        description="Produce a concise summary of one or more case documents.",
        icon_glyph="\U0001F4C4",  # 📄
        default_folders=[],
    ),
    "summarize_discovery": TaskSpec(
        task_id="summarize_discovery",
        title="Summarize Discovery",
        description="Summarize discovery responses with structure and citations.",
        icon_glyph="\U0001F4CB",  # 📋
        default_folders=["DISCOVERY/RESPONSES", "DISCOVERY"],
    ),
    "summarize_depositions": TaskSpec(
        task_id="summarize_depositions",
        title="Summarize Depositions",
        description="Generate a structured summary of one or more depositions.",
        icon_glyph="\U0001F399",  # 🎙
        default_folders=["DISCOVERY/TRANSCRIPTS", "DISCOVERY"],
    ),
    "medical_records": TaskSpec(
        task_id="medical_records",
        title="Medical Records Review",
        description="Extract and summarize medical records into a chronology.",
        icon_glyph="\U0001F3E5",  # 🏥
        default_folders=["RECORDS"],
    ),
}


def get_task(task_id: str) -> TaskSpec:
    """Return the TaskSpec for `task_id`. Raises KeyError if unknown."""
    return TASK_REGISTRY[task_id]


def list_tasks() -> list[TaskSpec]:
    """Return all registered tasks in registry-insertion order."""
    return list(TASK_REGISTRY.values())
```

- [ ] **Step 4: Run — confirm pass**

```bash
pytest tests/test_wizard/test_registry.py -v
```

Expected: 4 tests pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/registry.py tests/test_wizard/test_registry.py
git commit -m "$(cat <<'EOF'
feat(wizard): TaskSpec registry for the four initial task cards

TASK_REGISTRY is the single source of truth for available wizard
tasks. Each TaskSpec carries title, description, icon, and the
per-task default folder preferences used by the pre-Settings file
dialog.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 2.2: Create `TaskCard` widget

**Files:**
- Create: `icharlotte_core/ui/wizard/task_card.py`
- Create: `tests/test_wizard/test_task_card.py`

- [ ] **Step 1: Write smoke test**

Create `tests/test_wizard/test_task_card.py`:

```python
"""Smoke tests for TaskCard."""
import pytest

pytest.importorskip("pytestqt")
from PySide6.QtCore import Qt

from icharlotte_core.ui.wizard.registry import get_task
from icharlotte_core.ui.wizard.task_card import TaskCard


def test_card_displays_task_metadata(qtbot):
    spec = get_task("summarize_documents")
    card = TaskCard(spec)
    qtbot.addWidget(card)
    assert spec.title in card.title_label.text()
    assert spec.description in card.description_label.text()


def test_clicking_card_emits_signal(qtbot):
    spec = get_task("medical_records")
    card = TaskCard(spec)
    qtbot.addWidget(card)
    with qtbot.waitSignal(card.clicked, timeout=500) as blocker:
        qtbot.mouseClick(card, Qt.MouseButton.LeftButton)
    assert blocker.args == ["medical_records"]
```

- [ ] **Step 2: Run — confirm failure**

```bash
pytest tests/test_wizard/test_task_card.py -v
```

Expected: `ModuleNotFoundError`.

- [ ] **Step 3: Implement `TaskCard`**

Create `icharlotte_core/ui/wizard/task_card.py`:

```python
"""TaskCard — clickable card on the Wizard tab."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QMouseEvent
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QVBoxLayout

from .registry import TaskSpec


_CARD_STYLE = """
TaskCard {
    background-color: #ffffff;
    border: 1px solid #e0e0e0;
    border-radius: 10px;
}
TaskCard:hover {
    border-color: #1976D2;
    background-color: #fafcff;
}
"""

_ICON_TILE_STYLE = """
QLabel#icon_tile {
    background-color: #fff7e6;
    border-radius: 8px;
    font-size: 22px;
    qproperty-alignment: AlignCenter;
}
"""


class TaskCard(QFrame):
    """A single card representing a task. Clicking emits `clicked(task_id)`."""

    clicked = Signal(str)  # task_id

    def __init__(self, spec: TaskSpec, parent=None):
        super().__init__(parent)
        self._spec = spec
        self.setObjectName("TaskCard")
        self.setStyleSheet(_CARD_STYLE + _ICON_TILE_STYLE)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedSize(280, 140)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 14, 16, 14)
        outer.setSpacing(8)

        header = QHBoxLayout()
        header.setSpacing(10)

        self.icon_tile = QLabel(spec.icon_glyph)
        self.icon_tile.setObjectName("icon_tile")
        self.icon_tile.setFixedSize(36, 36)
        header.addWidget(self.icon_tile)

        self.title_label = QLabel(spec.title)
        self.title_label.setStyleSheet("font-size: 14px; font-weight: 600; color: #1a1a1a;")
        self.title_label.setWordWrap(True)
        header.addWidget(self.title_label, 1)

        outer.addLayout(header)

        self.description_label = QLabel(spec.description)
        self.description_label.setWordWrap(True)
        self.description_label.setStyleSheet("font-size: 12px; color: #666;")
        outer.addWidget(self.description_label, 1)

    @property
    def task_id(self) -> str:
        return self._spec.task_id

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit(self._spec.task_id)
        super().mousePressEvent(event)
```

- [ ] **Step 4: Run — confirm pass**

```bash
pytest tests/test_wizard/test_task_card.py -v
```

Expected: 2 tests pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/task_card.py tests/test_wizard/test_task_card.py
git commit -m "$(cat <<'EOF'
feat(wizard): TaskCard clickable card widget

Fixed-size 280x140 card with icon tile, title, and description. Emits
clicked(task_id) on mouse press. Styled to roughly match the UI
sample mockup.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 2.3: Create `WizardTab` shell (cards only, no Recent Tasks yet)

**Files:**
- Create: `icharlotte_core/ui/wizard/wizard_tab.py`
- Create: `tests/test_wizard/test_wizard_tab.py`

- [ ] **Step 1: Write smoke test**

Create `tests/test_wizard/test_wizard_tab.py`:

```python
"""Smoke tests for WizardTab."""
import pytest

pytest.importorskip("pytestqt")
from PySide6.QtCore import Qt

from icharlotte_core.ui.wizard.wizard_tab import WizardTab


def test_renders_four_cards(qtbot):
    tab = WizardTab()
    qtbot.addWidget(tab)
    assert len(tab.cards) == 4


def test_card_click_emits_task_requested(qtbot):
    tab = WizardTab()
    qtbot.addWidget(tab)
    # Pick the first card and click it.
    first_card = tab.cards[0]
    with qtbot.waitSignal(tab.task_requested, timeout=500) as blocker:
        qtbot.mouseClick(first_card, Qt.MouseButton.LeftButton)
    assert blocker.args[0] == first_card.task_id
```

- [ ] **Step 2: Run — confirm failure**

```bash
pytest tests/test_wizard/test_wizard_tab.py -v
```

Expected: `ModuleNotFoundError`.

- [ ] **Step 3: Implement `WizardTab`**

Create `icharlotte_core/ui/wizard/wizard_tab.py`:

```python
"""WizardTab — header + grid of TaskCards. Recent Tasks added in Phase 7."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QGridLayout,
    QLabel,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from .registry import list_tasks
from .task_card import TaskCard


_CARDS_PER_ROW = 3


class WizardTab(QWidget):
    """The 'What would you like to do?' card grid tab."""

    task_requested = Signal(str)  # task_id

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.cards: list[TaskCard] = []
        self._build_ui()

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(40, 32, 40, 32)
        outer.setSpacing(24)

        header = QLabel("What would you like to do?")
        header.setStyleSheet("font-size: 22px; font-weight: 400; color: #1a1a1a;")
        outer.addWidget(header)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)

        container = QWidget()
        grid = QGridLayout(container)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(20)
        grid.setVerticalSpacing(20)
        grid.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)

        for idx, spec in enumerate(list_tasks()):
            card = TaskCard(spec, parent=container)
            card.clicked.connect(self.task_requested.emit)
            row, col = divmod(idx, _CARDS_PER_ROW)
            grid.addWidget(card, row, col)
            self.cards.append(card)

        scroll.setWidget(container)
        outer.addWidget(scroll, 1)
```

- [ ] **Step 4: Run — confirm pass**

```bash
pytest tests/test_wizard/test_wizard_tab.py -v
```

Expected: 2 tests pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/wizard_tab.py tests/test_wizard/test_wizard_tab.py
git commit -m "$(cat <<'EOF'
feat(wizard): WizardTab with four task cards in a 3-column grid

Header label 'What would you like to do?' + scrolling grid of
TaskCards from TASK_REGISTRY. Card clicks bubble up as
task_requested(task_id). Recent Tasks section added in Phase 7.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 2.4: Add `WizardTab` to `MainWindow`

**Files:**
- Modify: `iCharlotte.py` — instantiate `WizardTab`, add to tabs, manage visibility

- [ ] **Step 1: Add the tab**

In `iCharlotte.py`, inside `setup_ui()`, immediately after the Master List tab is added (around line 598, just after `self.tabs.addTab(self.master_tab, "Master List")`), insert:

```python
        # --- Tab 1 (Wizard Mode only): Wizard ---
        from icharlotte_core.ui.wizard.wizard_tab import WizardTab
        self.wizard_tab = WizardTab(self)
        self.tabs.addTab(self.wizard_tab, "Wizard")
        # Temporary log-only handler. Phase 4 replaces this with task-tab creation.
        self.wizard_tab.task_requested.connect(
            lambda task_id: log_event(f"Wizard card clicked: {task_id}")
        )
```

- [ ] **Step 2: Update visibility logic**

In the existing `_apply_mode_visibility()` method (added in Task 1.6), also hide the Wizard tab when in Advanced Mode. Update the method body:

```python
    def _apply_mode_visibility(self, mode: str) -> None:
        """Show/hide tabs based on current mode."""
        is_wizard = (mode == "wizard")
        for i in range(self.tabs.count()):
            tab_text = self.tabs.tabText(i)
            if tab_text in self._WIZARD_HIDDEN_TABS:
                self.tabs.setTabVisible(i, not is_wizard)
            elif tab_text == "Wizard":
                self.tabs.setTabVisible(i, is_wizard)
            # Master List + task tabs (managed separately) stay visible.
        if not self.tabs.isTabVisible(self.tabs.currentIndex()):
            self.tabs.setCurrentIndex(0)
```

- [ ] **Step 3: Manual verification**

```bash
python iCharlotte.py
```

Expected:
1. App opens in Wizard Mode → only **Master List** and **Wizard** tabs visible.
2. Click the **Wizard** tab → header "What would you like to do?" + four cards in a 3-column grid (3 on top row, 1 on second row).
3. Click any card → check the log file or stderr for a "Wizard card clicked: <task_id>" entry. **No new tab is created yet** — that's Phase 4.
4. Switch to Advanced Mode → **Wizard tab hidden**, all old tabs visible.

- [ ] **Step 4: Commit**

```bash
git add iCharlotte.py
git commit -m "$(cat <<'EOF'
feat(wizard): mount WizardTab and toggle its visibility per mode

WizardTab is inserted after Master List. Visible only in Wizard Mode.
Card clicks log to the event log; task-tab creation arrives in Phase 4.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

# Phase 3 — File picker + default folder resolution

After Phase 3, clicking a task card opens the appropriate `QFileDialog` rooted at the task's default folder (still no tab creation yet — Phase 4 wires the rest).

### Task 3.1: `file_picker.resolve_default_folder` helper

**Files:**
- Create: `icharlotte_core/ui/wizard/file_picker.py`
- Create: `tests/test_wizard/test_file_picker.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_wizard/test_file_picker.py`:

```python
"""Tests for default-folder resolution."""
import os
import pytest

from icharlotte_core.ui.wizard.file_picker import resolve_default_folder


def test_first_existing_pref_wins(tmp_path):
    (tmp_path / "DISCOVERY" / "RESPONSES").mkdir(parents=True)
    (tmp_path / "DISCOVERY").mkdir(exist_ok=True)
    result = resolve_default_folder(str(tmp_path), ["DISCOVERY/RESPONSES", "DISCOVERY"])
    assert os.path.normpath(result) == os.path.normpath(str(tmp_path / "DISCOVERY" / "RESPONSES"))


def test_falls_back_to_second_pref(tmp_path):
    (tmp_path / "DISCOVERY").mkdir()
    result = resolve_default_folder(str(tmp_path), ["DISCOVERY/RESPONSES", "DISCOVERY"])
    assert os.path.normpath(result) == os.path.normpath(str(tmp_path / "DISCOVERY"))


def test_case_insensitive_match(tmp_path):
    (tmp_path / "discovery" / "responses").mkdir(parents=True)
    result = resolve_default_folder(str(tmp_path), ["DISCOVERY/RESPONSES"])
    assert os.path.normpath(result).lower() == os.path.normpath(str(tmp_path / "discovery" / "responses")).lower()


def test_no_match_returns_case_root(tmp_path):
    result = resolve_default_folder(str(tmp_path), ["DOESNT/EXIST"])
    assert os.path.normpath(result) == os.path.normpath(str(tmp_path))


def test_empty_prefs_returns_case_root(tmp_path):
    result = resolve_default_folder(str(tmp_path), [])
    assert os.path.normpath(result) == os.path.normpath(str(tmp_path))


def test_missing_case_root_returns_input_unchanged(tmp_path):
    fake_root = str(tmp_path / "nonexistent")
    # Should not raise — returns the input string even though it doesn't exist.
    result = resolve_default_folder(fake_root, ["RECORDS"])
    assert result == fake_root
```

- [ ] **Step 2: Run — confirm failure**

```bash
pytest tests/test_wizard/test_file_picker.py -v
```

Expected: `ModuleNotFoundError`.

- [ ] **Step 3: Implement**

Create `icharlotte_core/ui/wizard/file_picker.py`:

```python
"""file_picker — default-folder resolution + (later) multi-file QFileDialog launch."""
import os
from typing import Iterable


def resolve_default_folder(case_root: str, prefs: Iterable[str]) -> str:
    """Return the first existing folder under `case_root` matching any pref.

    `prefs` is a list of relative paths like "DISCOVERY/RESPONSES". Matching is
    case-insensitive (Windows preserves on-disk case but users vary). Returns
    `case_root` itself if no pref matches or `prefs` is empty. Never raises.
    """
    if not os.path.isdir(case_root):
        return case_root

    for pref in prefs:
        candidate = _resolve_case_insensitive(case_root, pref)
        if candidate is not None and os.path.isdir(candidate):
            return candidate
    return case_root


def _resolve_case_insensitive(root: str, rel_path: str) -> str | None:
    """Walk `rel_path` segments under `root`, matching each segment case-insensitively.

    Returns the actual on-disk path if all segments resolve, else None.
    """
    parts = [p for p in rel_path.replace("\\", "/").split("/") if p]
    current = root
    for part in parts:
        try:
            entries = os.listdir(current)
        except OSError:
            return None
        target_lc = part.lower()
        match = next((e for e in entries if e.lower() == target_lc), None)
        if match is None:
            return None
        current = os.path.join(current, match)
    return current
```

- [ ] **Step 4: Run — confirm pass**

```bash
pytest tests/test_wizard/test_file_picker.py -v
```

Expected: 6 tests pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/file_picker.py tests/test_wizard/test_file_picker.py
git commit -m "$(cat <<'EOF'
feat(wizard): case-insensitive default-folder resolution

resolve_default_folder() walks a preference list of relative subfolders
under a case root and returns the first existing match (case-insensitive
to handle on-disk case variation). Falls back to the case root, never
raises.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 3.2: Wire pre-Settings file dialog into card-click handler

**Files:**
- Modify: `iCharlotte.py` — replace the temporary log handler with a real `_open_task_tab(task_id)` method that pops the file dialog (still no TaskTab — that's Phase 4)

- [ ] **Step 1: Add `_open_task_tab` stub**

In `iCharlotte.py`, inside `MainWindow`, add a method:

```python
    def _open_task_tab(self, task_id: str) -> None:
        """Phase 3 version: pops the file dialog. Phase 4 will create a real TaskTab."""
        from icharlotte_core.ui.wizard.registry import get_task
        from icharlotte_core.ui.wizard.file_picker import resolve_default_folder
        from PySide6.QtWidgets import QFileDialog

        if not self.case_path:
            QMessageBox.information(self, "No case loaded", "Open a case from the Master List first.")
            return

        spec = get_task(task_id)
        start_dir = resolve_default_folder(self.case_path, spec.default_folders)
        files, _ = QFileDialog.getOpenFileNames(
            self,
            f"Select files for {spec.title}",
            start_dir,
            "All files (*.*)",
        )
        if not files:
            return  # user cancelled → no tab created
        log_event(f"[wizard] {task_id}: selected {len(files)} files from {start_dir}")
```

- [ ] **Step 2: Replace the temporary log handler**

Find the line added in Task 2.4 inside `setup_ui()`:

```python
        self.wizard_tab.task_requested.connect(
            lambda task_id: log_event(f"Wizard card clicked: {task_id}")
        )
```

Replace it with:

```python
        self.wizard_tab.task_requested.connect(self._open_task_tab)
```

- [ ] **Step 3: Manual verification**

```bash
python iCharlotte.py
```

Open a case from the Master List by double-clicking it. Switch to the Wizard tab. Click each of the four cards in turn and verify the file dialog opens at the right folder:

| Card | Expected starting folder (if it exists) |
|---|---|
| Summarize Documents   | case root |
| Summarize Discovery   | `<case>/DISCOVERY/RESPONSES`, falling back to `<case>/DISCOVERY`, then case root |
| Summarize Depositions | `<case>/DISCOVERY/TRANSCRIPTS`, falling back to `<case>/DISCOVERY`, then case root |
| Medical Records       | `<case>/RECORDS`, falling back to case root |

Click Cancel → nothing should happen (no tab, no log "selected N files"). Select 1+ files → log line appears.

- [ ] **Step 4: Commit**

```bash
git add iCharlotte.py
git commit -m "$(cat <<'EOF'
feat(wizard): pre-Settings file dialog wired to card clicks

Clicking a task card now opens QFileDialog rooted at the task's
default folder (falling back per-task). Cancel = no-op. Phase 4 will
turn the selection into an actual TaskTab.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

# Phase 4 — TaskTab + Settings/Status/Output scaffolding

After Phase 4, clicking a card with files selected creates a real task tab. The Settings page is a placeholder with a Proceed button that fakes a 2-second run and lands on a stub Output page.

### Task 4.1: Create page base + `SettingsPage` placeholder

**Files:**
- Create: `icharlotte_core/ui/wizard/pages/__init__.py`
- Create: `icharlotte_core/ui/wizard/pages/settings_page.py`

- [ ] **Step 1: Create package file**

Create `icharlotte_core/ui/wizard/pages/__init__.py`:

```python
"""Task tab pages: Settings → Status → Output."""
```

- [ ] **Step 2: Implement placeholder `SettingsPage`**

Create `icharlotte_core/ui/wizard/pages/settings_page.py`:

```python
"""SettingsPage — pre-run configuration for a task tab.

This is a placeholder for the per-task settings UI; real per-task
settings are defined in follow-up specs. For now it shows:
  - The list of selected input files (with a Remove button per row).
  - A 'Settings for <task title> — to be defined' label.
  - A Proceed button bottom-right.
"""
import os
from typing import List

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from ..registry import TaskSpec


class SettingsPage(QWidget):
    """Configurable inputs + Proceed button. Emits proceed_requested(settings_dict)."""

    proceed_requested = Signal(dict)  # settings dict (placeholder)

    def __init__(self, spec: TaskSpec, files: List[str], parent: QWidget | None = None):
        super().__init__(parent)
        self._spec = spec
        self._files: List[str] = list(files)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(16)

        # Files section
        files_label = QLabel(self._format_files_label())
        files_label.setStyleSheet("font-size: 13px; font-weight: 600;")
        outer.addWidget(files_label)
        self.files_label = files_label

        self.files_list = QListWidget()
        self.files_list.setMaximumHeight(150)
        self._refresh_files_list()
        outer.addWidget(self.files_list)

        # Placeholder body
        body = QLabel(f"Settings for {spec.title} — to be defined.")
        body.setStyleSheet("color: #666; font-style: italic; padding: 24px;")
        body.setAlignment(body.alignment())
        outer.addWidget(body, 1)

        # Proceed button bottom-right
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.proceed_btn = QPushButton("Proceed")
        self.proceed_btn.setStyleSheet(
            "background-color: #1976D2; color: white; font-weight: 600; padding: 8px 24px; border-radius: 4px;"
        )
        self.proceed_btn.clicked.connect(self._on_proceed)
        btn_row.addWidget(self.proceed_btn)
        outer.addLayout(btn_row)

        self._update_proceed_enabled()

    def _format_files_label(self) -> str:
        return f"Files ({len(self._files)})"

    def _refresh_files_list(self) -> None:
        self.files_list.clear()
        for path in self._files:
            display = os.path.basename(path)
            item = QListWidgetItem(display)
            item.setToolTip(path)
            if not os.path.exists(path):
                item.setText(f"{display}  (missing)")
                item.setForeground(item.foreground())  # placeholder; greyed via stylesheet if desired
            self.files_list.addItem(item)
        self.files_label.setText(self._format_files_label())
        self._update_proceed_enabled()

    def _update_proceed_enabled(self) -> None:
        self.proceed_btn.setEnabled(len(self._files) > 0)

    def _on_proceed(self) -> None:
        self.proceed_requested.emit(self.to_dict())

    # ---- Persistence-friendly API ----

    @property
    def files(self) -> list[str]:
        return list(self._files)

    def to_dict(self) -> dict:
        """Placeholder settings dict. Real per-task settings will override."""
        return {}

    def from_dict(self, data: dict) -> None:
        """Placeholder — real subclasses will restore form state."""
        return None
```

- [ ] **Step 3: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/__init__.py icharlotte_core/ui/wizard/pages/settings_page.py
git commit -m "$(cat <<'EOF'
feat(wizard): placeholder SettingsPage with files list + Proceed

Generic placeholder used by every task tab until per-task settings are
defined. Shows the selected files with tooltips, a 'to be defined'
body, and a Proceed button. to_dict()/from_dict() stubs in place for
later persistence wiring.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 4.2: `StatusPage`

**Files:**
- Create: `icharlotte_core/ui/wizard/pages/status_page.py`

- [ ] **Step 1: Implement**

Create `icharlotte_core/ui/wizard/pages/status_page.py`:

```python
"""StatusPage — progress bar + log + Cancel button while a task is running."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class StatusPage(QWidget):
    """Shows progress + log lines. Emits cancel_requested when Cancel is clicked."""

    cancel_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(12)

        self.status_label = QLabel("Starting…")
        self.status_label.setStyleSheet("font-size: 13px; font-weight: 600;")
        outer.addWidget(self.status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # indeterminate by default
        outer.addWidget(self.progress_bar)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setStyleSheet("font-family: Consolas, 'Courier New', monospace; font-size: 11px;")
        outer.addWidget(self.log_view, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setStyleSheet(
            "background-color: #d32f2f; color: white; font-weight: 600; padding: 8px 20px; border-radius: 4px;"
        )
        self.cancel_btn.clicked.connect(self._on_cancel)
        btn_row.addWidget(self.cancel_btn)
        outer.addLayout(btn_row)

    def reset(self) -> None:
        self.status_label.setText("Starting…")
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setValue(0)
        self.log_view.clear()
        self.cancel_btn.setEnabled(True)
        self.cancel_btn.setText("Cancel")

    # ---- Slots / public API for the worker connection ----

    def on_status(self, line: str) -> None:
        self.status_label.setText(line)
        self.log_view.appendPlainText(line)
        self.log_view.verticalScrollBar().setValue(self.log_view.verticalScrollBar().maximum())

    def on_progress(self, pct: int) -> None:
        if pct < 0 or pct > 100:
            return
        if self.progress_bar.maximum() == 0:
            self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(pct)

    def _on_cancel(self) -> None:
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.setText("Cancelling…")
        self.cancel_requested.emit()
```

- [ ] **Step 2: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/status_page.py
git commit -m "$(cat <<'EOF'
feat(wizard): StatusPage with progress, log, and Cancel button

Progress bar starts indeterminate; flips to 0-100 the first time a
percent is reported. Log view auto-scrolls. Cancel transitions the
button to a disabled 'Cancelling…' state and emits cancel_requested.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 4.3: `OutputPage` scaffolding (minimal — full editor in Phase 8)

**Files:**
- Create: `icharlotte_core/ui/wizard/pages/output_page.py`

- [ ] **Step 1: Implement minimal scaffold**

Create `icharlotte_core/ui/wizard/pages/output_page.py`:

```python
"""OutputPage scaffold. Phase 8 replaces the placeholder body with the mammoth
.docx → HTML editor + Save/Open in Word actions.
"""
import os

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


class OutputPage(QWidget):
    """Shows the task's output and action buttons.

    Phase 4: minimal scaffold — header with file name + Open in Word + a plain
    text view of the file. Full mammoth-rendered editor + Save round-trip arrive
    in Phase 8.
    """

    rerun_requested = Signal()
    edit_settings_requested = Signal()
    open_in_word_requested = Signal()
    copy_all_requested = Signal()
    save_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._output_path: str | None = None

        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(12)

        header = QHBoxLayout()
        self.file_label = QLabel("File: —")
        self.file_label.setStyleSheet("font-weight: 600;")
        header.addWidget(self.file_label, 1)
        self.open_in_word_btn = QPushButton("Open in Word")
        self.open_in_word_btn.clicked.connect(self.open_in_word_requested.emit)
        header.addWidget(self.open_in_word_btn)
        outer.addLayout(header)

        self.editor = QTextEdit()
        self.editor.setReadOnly(False)
        outer.addWidget(self.editor, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.copy_all_btn = QPushButton("Copy All")
        self.copy_all_btn.clicked.connect(self.copy_all_requested.emit)
        btn_row.addWidget(self.copy_all_btn)
        self.rerun_btn = QPushButton("Re-run")
        self.rerun_btn.clicked.connect(self.rerun_requested.emit)
        btn_row.addWidget(self.rerun_btn)
        self.edit_settings_btn = QPushButton("Edit Settings & Re-run")
        self.edit_settings_btn.clicked.connect(self.edit_settings_requested.emit)
        btn_row.addWidget(self.edit_settings_btn)
        self.save_btn = QPushButton("Save")
        self.save_btn.setStyleSheet("background-color: #1976D2; color: white; font-weight: 600; padding: 6px 18px;")
        self.save_btn.clicked.connect(self.save_requested.emit)
        btn_row.addWidget(self.save_btn)
        outer.addLayout(btn_row)

    # ---- Public API ----

    @property
    def output_path(self) -> str | None:
        return self._output_path

    def load_output(self, output_path: str) -> None:
        """Phase 4 stub: shows the file name and a placeholder body."""
        self._output_path = output_path
        self.file_label.setText(f"File: {os.path.basename(output_path)}")
        self.editor.setPlainText(
            f"(Phase 4 scaffold) Output file at:\n{output_path}\n\n"
            "Full mammoth-rendered editor arrives in Phase 8."
        )
```

- [ ] **Step 2: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/output_page.py
git commit -m "$(cat <<'EOF'
feat(wizard): OutputPage scaffold

Minimal version with action-button row (Copy All, Re-run, Edit
Settings & Re-run, Save) and an editable QTextEdit. load_output()
shows a placeholder; Phase 8 swaps the body for mammoth-rendered
.docx content + a real Save round-trip.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 4.4: `TaskTab` orchestrator + fake worker for end-to-end smoke test

**Files:**
- Create: `icharlotte_core/ui/wizard/task_tab.py`
- Create: `tests/test_wizard/test_task_tab.py`

- [ ] **Step 1: Write smoke test**

Create `tests/test_wizard/test_task_tab.py`:

```python
"""Smoke test for TaskTab state machine."""
import pytest

pytest.importorskip("pytestqt")
from icharlotte_core.ui.wizard.registry import get_task
from icharlotte_core.ui.wizard.task_tab import TaskTab, PAGE_SETTINGS, PAGE_STATUS, PAGE_OUTPUT


def test_initial_state_is_settings(qtbot):
    tab = TaskTab(get_task("summarize_documents"), files=["/tmp/x.pdf"])
    qtbot.addWidget(tab)
    assert tab.current_page == PAGE_SETTINGS


def test_proceed_transitions_to_status(qtbot):
    tab = TaskTab(get_task("summarize_documents"), files=["/tmp/x.pdf"])
    qtbot.addWidget(tab)
    # Disable the fake worker so the test is deterministic.
    tab._fake_worker_delay_ms = 0
    tab.settings_page._on_proceed()
    assert tab.current_page in (PAGE_STATUS, PAGE_OUTPUT)  # 0ms timer may already have fired


def test_show_output_transitions(qtbot):
    tab = TaskTab(get_task("summarize_documents"), files=["/tmp/x.pdf"])
    qtbot.addWidget(tab)
    tab._show_output("/tmp/fake_output.docx")
    assert tab.current_page == PAGE_OUTPUT
    assert tab.output_page.output_path == "/tmp/fake_output.docx"
```

- [ ] **Step 2: Run — confirm failure**

```bash
pytest tests/test_wizard/test_task_tab.py -v
```

Expected: `ModuleNotFoundError`.

- [ ] **Step 3: Implement `TaskTab` with a fake worker (real workers arrive in Phase 5)**

Create `icharlotte_core/ui/wizard/task_tab.py`:

```python
"""TaskTab — QStackedWidget orchestrating Settings → Status → Output for one task.

Phase 4 ships with a 'fake worker' that just sleeps and emits a synthetic
output path. Phase 5 replaces that with real subprocess-based runners.
"""
from typing import List

from PySide6.QtCore import QTimer, Signal
from PySide6.QtWidgets import QStackedWidget, QWidget

from .pages.output_page import OutputPage
from .pages.settings_page import SettingsPage
from .pages.status_page import StatusPage
from .registry import TaskSpec


PAGE_SETTINGS = 0
PAGE_STATUS = 1
PAGE_OUTPUT = 2


class TaskTab(QStackedWidget):
    """Stateful container for one running task. Owns its own worker."""

    closed = Signal()  # emitted when the tab is being removed

    def __init__(
        self,
        spec: TaskSpec,
        files: List[str],
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self._spec = spec
        self._files: List[str] = list(files)
        self._worker = None
        self._fake_worker_delay_ms = 2000  # Phase 4 fake-run duration

        self.settings_page = SettingsPage(spec, files=files)
        self.status_page = StatusPage()
        self.output_page = OutputPage()

        self.addWidget(self.settings_page)  # index 0 = PAGE_SETTINGS
        self.addWidget(self.status_page)    # index 1 = PAGE_STATUS
        self.addWidget(self.output_page)    # index 2 = PAGE_OUTPUT

        self.settings_page.proceed_requested.connect(self._on_proceed)
        self.status_page.cancel_requested.connect(self._on_cancel)
        self.output_page.edit_settings_requested.connect(self._on_edit_settings)
        self.output_page.rerun_requested.connect(self._on_rerun)

    @property
    def spec(self) -> TaskSpec:
        return self._spec

    @property
    def files(self) -> List[str]:
        return list(self._files)

    @property
    def current_page(self) -> int:
        return self.currentIndex()

    # ---- Transitions ----

    def _on_proceed(self, settings_dict: dict) -> None:
        self.status_page.reset()
        self.setCurrentIndex(PAGE_STATUS)
        self._start_run(settings_dict)

    def _on_cancel(self) -> None:
        if self._worker is not None and hasattr(self._worker, "cancel"):
            self._worker.cancel()
        # Phase 4 fake worker has no cancel — just snap back to Settings.
        self._worker = None
        self.setCurrentIndex(PAGE_SETTINGS)

    def _on_edit_settings(self) -> None:
        self.setCurrentIndex(PAGE_SETTINGS)

    def _on_rerun(self) -> None:
        self._on_proceed(self.settings_page.to_dict())

    def _show_output(self, output_path: str) -> None:
        self.output_page.load_output(output_path)
        self.setCurrentIndex(PAGE_OUTPUT)

    # ---- Worker (Phase 4 fake) ----

    def _start_run(self, settings_dict: dict) -> None:
        self.status_page.on_status(f"Running {self._spec.title}…")
        self.status_page.on_status(f"Inputs: {len(self._files)} file(s)")
        # Phase 4 fake: after a short delay, "finish" with a stub path.
        delay = max(0, self._fake_worker_delay_ms)
        if self._files:
            stub_output = self._files[0]  # not a real .docx; replaced in Phase 5
        else:
            stub_output = ""
        QTimer.singleShot(delay, lambda: self._show_output(stub_output))
```

- [ ] **Step 4: Run — confirm pass**

```bash
pytest tests/test_wizard/test_task_tab.py -v
```

Expected: 3 tests pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/task_tab.py tests/test_wizard/test_task_tab.py
git commit -m "$(cat <<'EOF'
feat(wizard): TaskTab QStackedWidget state machine + fake worker

Three pages (Settings/Status/Output) inside a QStackedWidget. Proceed
starts a placeholder fake run that ends after 2s on the Output page.
Real subprocess workers arrive in Phase 5.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 4.5: Multi-instance tab title suffixing

**Files:**
- Create: `icharlotte_core/ui/wizard/instance_naming.py`
- Create: `tests/test_wizard/test_instance_naming.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_wizard/test_instance_naming.py`:

```python
"""Tests for tab-title disambiguation when the same task is opened twice."""
from icharlotte_core.ui.wizard.instance_naming import next_instance_suffix


def test_no_existing_returns_empty_suffix():
    assert next_instance_suffix("Summarize Documents", existing_titles=[]) == ""


def test_existing_base_only_returns_2():
    assert next_instance_suffix("Summarize Documents", existing_titles=["Summarize Documents"]) == "(2)"


def test_fills_gap_with_lowest_unused():
    existing = ["Summarize Documents", "Summarize Documents (3)"]
    assert next_instance_suffix("Summarize Documents", existing_titles=existing) == "(2)"


def test_returns_highest_plus_one_when_no_gap():
    existing = ["Summarize Documents", "Summarize Documents (2)", "Summarize Documents (3)"]
    assert next_instance_suffix("Summarize Documents", existing_titles=existing) == "(4)"


def test_ignores_unrelated_titles():
    existing = ["Medical Records", "Summarize Discovery"]
    assert next_instance_suffix("Summarize Documents", existing_titles=existing) == ""
```

- [ ] **Step 2: Run — confirm failure**

```bash
pytest tests/test_wizard/test_instance_naming.py -v
```

Expected: `ModuleNotFoundError`.

- [ ] **Step 3: Implement**

Create `icharlotte_core/ui/wizard/instance_naming.py`:

```python
"""Compute the next instance suffix for a duplicated task tab title."""
import re
from typing import Iterable


def next_instance_suffix(base_title: str, existing_titles: Iterable[str]) -> str:
    """Return the suffix to append to `base_title` for a new task tab.

    - If no existing tab uses `base_title` (with or without suffix), returns "".
    - Otherwise returns "(N)" where N is the lowest positive integer >= 2 that
      isn't already taken by a tab titled `base_title (N)`.

    Examples:
      base_title="Summarize Documents", existing=[] -> ""
      existing=["Summarize Documents"] -> "(2)"
      existing=["Summarize Documents", "Summarize Documents (3)"] -> "(2)"
    """
    existing = list(existing_titles)
    pattern = re.compile(rf"^{re.escape(base_title)}(?: \((\d+)\))?$")
    used_ns: set[int] = set()
    has_base = False
    for t in existing:
        m = pattern.match(t)
        if not m:
            continue
        num_str = m.group(1)
        if num_str is None:
            has_base = True
            used_ns.add(1)
        else:
            used_ns.add(int(num_str))
    if not has_base:
        return ""
    n = 2
    while n in used_ns:
        n += 1
    return f"({n})"
```

- [ ] **Step 4: Run — confirm pass**

```bash
pytest tests/test_wizard/test_instance_naming.py -v
```

Expected: 5 tests pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/instance_naming.py tests/test_wizard/test_instance_naming.py
git commit -m "$(cat <<'EOF'
feat(wizard): tab-title suffix logic for duplicate task tabs

next_instance_suffix() returns the lowest unused integer suffix for a
new task tab when one or more tabs of the same task type already exist.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 4.6: Wire `_open_task_tab` to create real `TaskTab`s

**Files:**
- Modify: `iCharlotte.py` — replace the Phase 3 logging stub with real TaskTab creation, suffix, and close-button handling

- [ ] **Step 1: Replace `_open_task_tab`**

In `iCharlotte.py`, find the `_open_task_tab` method added in Task 3.2 and **replace** it with:

```python
    def _open_task_tab(self, task_id: str) -> None:
        from icharlotte_core.ui.wizard.registry import get_task
        from icharlotte_core.ui.wizard.file_picker import resolve_default_folder
        from icharlotte_core.ui.wizard.task_tab import TaskTab
        from icharlotte_core.ui.wizard.instance_naming import next_instance_suffix
        from PySide6.QtWidgets import QFileDialog

        if not self.case_path:
            QMessageBox.information(self, "No case loaded", "Open a case from the Master List first.")
            return

        spec = get_task(task_id)
        start_dir = resolve_default_folder(self.case_path, spec.default_folders)
        files, _ = QFileDialog.getOpenFileNames(
            self,
            f"Select files for {spec.title}",
            start_dir,
            "All files (*.*)",
        )
        if not files:
            return  # user cancelled → no tab created

        # Compute title with suffix for multi-instance.
        existing_titles = [self.tabs.tabText(i) for i in range(self.tabs.count())]
        suffix = next_instance_suffix(spec.title, existing_titles)
        title = f"{spec.title} {suffix}".strip()

        task_tab = TaskTab(spec=spec, files=files, parent=self)
        task_tab.setProperty("wizard_task_id", spec.task_id)
        task_tab.setProperty("wizard_instance_suffix", suffix)
        new_index = self.tabs.addTab(task_tab, title)
        self.tabs.setCurrentIndex(new_index)
        log_event(f"[wizard] opened task tab '{title}' with {len(files)} file(s)")
```

- [ ] **Step 2: Enable per-tab close buttons (closeable task tabs only)**

In `setup_ui()`, after `self.tabs = QTabWidget()` is created (around line 558), add:

```python
        self.tabs.setTabsClosable(False)  # default: not closeable
        # We will set the close button only on TaskTabs via _set_close_button_for_task_tabs().
        self.tabs.tabCloseRequested.connect(self._on_tab_close_requested)
```

Then add helper methods to `MainWindow`:

```python
    def _on_tab_close_requested(self, index: int) -> None:
        """Only TaskTabs are closeable (they carry a 'wizard_task_id' property)."""
        widget = self.tabs.widget(index)
        if widget is None:
            return
        if widget.property("wizard_task_id") is None:
            return  # not a task tab; ignore
        # Phase 4: just remove. Cancellation hook arrives in Phase 5.
        self.tabs.removeTab(index)
        widget.deleteLater()
```

To actually show close buttons on task tabs (but not on the fixed tabs), set `tabsClosable` to True globally and hide the close button on non-task tabs after each addition. Update `setup_ui()` `setTabsClosable` line to:

```python
        self.tabs.setTabsClosable(True)
        self.tabs.tabCloseRequested.connect(self._on_tab_close_requested)
        # Hide close buttons on the fixed Master List / Wizard / Advanced tabs.
        # (We'll re-hide after every addTab via _hide_fixed_close_buttons.)
```

Add the helper to `MainWindow`:

```python
    def _hide_fixed_close_buttons(self) -> None:
        """Hide close buttons on tabs that are not TaskTabs."""
        from PySide6.QtWidgets import QTabBar
        bar = self.tabs.tabBar()
        for i in range(self.tabs.count()):
            widget = self.tabs.widget(i)
            is_task_tab = widget is not None and widget.property("wizard_task_id") is not None
            for side in (QTabBar.ButtonPosition.RightSide, QTabBar.ButtonPosition.LeftSide):
                btn = bar.tabButton(i, side)
                if btn is not None:
                    btn.setVisible(is_task_tab)
```

Call `self._hide_fixed_close_buttons()` at the end of `setup_ui()` and also at the end of `_open_task_tab()`.

- [ ] **Step 3: Manual verification**

```bash
python iCharlotte.py
```

1. Open a case → switch to Wizard tab → click Summarize Documents → pick 1-2 files → OK.
2. A new tab named "Summarize Documents" appears to the right of Wizard, automatically focused. The Settings page shows the file list and a Proceed button.
3. Click Proceed → flips to Status page; after ~2s flips to Output page with a placeholder.
4. Click Summarize Documents again → second tab "Summarize Documents (2)".
5. Close the (2) tab via its X. Click Summarize Documents again → suffix should be "(2)" again (lowest unused).
6. Confirm Master List and Wizard tabs **do not** have close buttons.

- [ ] **Step 4: Commit**

```bash
git add iCharlotte.py
git commit -m "$(cat <<'EOF'
feat(wizard): real TaskTab creation + closeable task tabs

_open_task_tab() now creates a TaskTab from registry + selected files,
appends it to the QTabWidget with a unique suffix, and shows a close
button on task tabs only. Master List / Wizard tabs remain pinned.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

# Phase 5 — Runner shims for the four existing agents

After Phase 5, clicking Proceed actually runs the real agent (`Scripts/summarize.py` etc.) as a `QProcess`, streams its stdout into the Status page, and lands on the Output page bound to the agent's actual `.docx` output. The TaskTab fake worker is removed.

> **Important context** — the existing `Scripts/*.py` agents are already invoked elsewhere in the app (see `MainWindow.create_enhanced_agent_button` and the AGENTS list at `iCharlotte.py:742`). They each accept a `--file_number` argument and write outputs into `<case>/NOTES/AI Output/`. We do **not** modify those scripts.

### Task 5.1: `BaseWorker` + signals contract

**Files:**
- Create: `icharlotte_core/ui/wizard/runners/__init__.py`
- Create: `icharlotte_core/ui/wizard/runners/base.py`
- Create: `tests/test_wizard/test_runners_base.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_wizard/test_runners_base.py`:

```python
"""Tests for BaseWorker contract — cancel flag, signal surface."""
import pytest

pytest.importorskip("pytestqt")

from icharlotte_core.ui.wizard.runners.base import BaseWorker


def test_cancel_sets_flag(qtbot):
    w = BaseWorker(case_path="/tmp/case", file_number="0000.000", files=[], settings={})
    assert w.is_cancel_requested is False
    w.cancel()
    assert w.is_cancel_requested is True


def test_signals_present(qtbot):
    w = BaseWorker(case_path="/tmp/case", file_number="0000.000", files=[], settings={})
    # Just touch them to ensure the attributes exist.
    assert w.status is not None
    assert w.progress is not None
    assert w.finished is not None
    assert w.failed is not None
    assert w.cancelled is not None
```

- [ ] **Step 2: Run — confirm failure**

```bash
pytest tests/test_wizard/test_runners_base.py -v
```

Expected: `ModuleNotFoundError`.

- [ ] **Step 3: Implement**

Create `icharlotte_core/ui/wizard/runners/__init__.py`:

```python
"""Wizard task runners — thin shims around existing Scripts agents."""
```

Create `icharlotte_core/ui/wizard/runners/base.py`:

```python
"""BaseWorker — common signal surface + cancellation contract for wizard runners."""
from typing import List

from PySide6.QtCore import QObject, Signal


class BaseWorker(QObject):
    """Abstract worker for wizard tasks.

    Subclasses implement start() (which kicks off a QProcess or QThread) and call
    self._on_status / _on_progress / _on_finished / _on_failed as the work proceeds.
    Cancellation is cooperative: cancel() flips a flag; subclasses decide how to
    honor it (e.g., terminating a QProcess, polling the flag in a loop).
    """

    status = Signal(str)         # one log line
    progress = Signal(int)       # 0-100
    finished = Signal(str)       # output_path (.docx)
    failed = Signal(str)         # error message
    cancelled = Signal()         # emitted after cancel takes effect

    def __init__(
        self,
        case_path: str,
        file_number: str,
        files: List[str],
        settings: dict,
        parent: QObject | None = None,
    ):
        super().__init__(parent)
        self.case_path = case_path
        self.file_number = file_number
        self.files = list(files)
        self.settings = dict(settings)
        self._cancel_requested = False

    @property
    def is_cancel_requested(self) -> bool:
        return self._cancel_requested

    def cancel(self) -> None:
        """Request cancellation. Subclasses may override to take additional action."""
        self._cancel_requested = True

    def start(self) -> None:
        """Subclasses must implement."""
        raise NotImplementedError
```

- [ ] **Step 4: Run — confirm pass**

```bash
pytest tests/test_wizard/test_runners_base.py -v
```

Expected: 2 tests pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/runners/__init__.py icharlotte_core/ui/wizard/runners/base.py tests/test_wizard/test_runners_base.py
git commit -m "$(cat <<'EOF'
feat(wizard): BaseWorker contract for task runners

Defines the common signal surface (status/progress/finished/failed/
cancelled) and the cooperative-cancel flag pattern used by all wizard
task workers.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 5.2: `SubprocessWorker` — generic QProcess wrapper

**Files:**
- Create: `icharlotte_core/ui/wizard/runners/subprocess_worker.py`

- [ ] **Step 1: Implement**

Create `icharlotte_core/ui/wizard/runners/subprocess_worker.py`:

```python
"""SubprocessWorker — runs an existing Scripts/*.py agent as a QProcess.

Stdout is parsed for:
  - lines starting with "PROGRESS:" → progress(int) updates
  - all other lines             → status(str) log emissions

When the process exits with returncode 0, we look for the .docx the agent
wrote into <case>/NOTES/AI Output/ that didn't exist before the run started.
That path is emitted as finished(output_path).

If returncode != 0 or no new .docx appears, we emit failed(...).

Cancellation: cancel() calls QProcess.terminate() (and then kill() after
2 seconds if it hasn't exited), then emits cancelled().
"""
import os
import re
import time
from typing import List

from PySide6.QtCore import QProcess, QTimer

from .base import BaseWorker


_AI_OUTPUT_SUBPATH = os.path.join("NOTES", "AI Output")
_PROGRESS_RE = re.compile(r"^PROGRESS:\s*(\d+)\s*$")


class SubprocessWorker(BaseWorker):
    """Run a Scripts/*.py agent via QProcess."""

    def __init__(
        self,
        script_name: str,           # e.g. "summarize.py"
        case_path: str,
        file_number: str,
        files: List[str],
        settings: dict,
        parent=None,
    ):
        super().__init__(case_path, file_number, files, settings, parent)
        self._script_name = script_name
        self._process: QProcess | None = None
        self._pre_existing_outputs: set[str] = set()
        self._stdout_buf = bytearray()

    # ---- Lifecycle ----

    def start(self) -> None:
        self._pre_existing_outputs = self._scan_outputs()
        self._process = QProcess(self)
        self._process.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)
        self._process.readyReadStandardOutput.connect(self._drain_stdout)
        self._process.finished.connect(self._on_process_finished)
        self._process.errorOccurred.connect(self._on_process_error)

        # Build argv. Existing agents accept --file_number; some accept --files
        # (forward as repeated --file args; the agent ignores unknown flags
        # gracefully because we control which script we call).
        # __file__ lives 5 levels deep:
        #   <root>/icharlotte_core/ui/wizard/runners/subprocess_worker.py
        here = os.path.abspath(__file__)
        repo_root = here
        for _ in range(5):
            repo_root = os.path.dirname(repo_root)
        script_path = os.path.join(repo_root, "Scripts", self._script_name)
        argv = ["-u", script_path, "--file_number", self.file_number]
        for f in self.files:
            argv.extend(["--file", f])

        import sys
        self.status.emit(f"Running: python {' '.join(os.path.basename(a) if a == script_path else a for a in argv)}")
        self._process.start(sys.executable, argv)

    def cancel(self) -> None:
        super().cancel()
        if self._process is None:
            self.cancelled.emit()
            return
        self.status.emit("Cancellation requested…")
        self._process.terminate()
        # Hard-kill backup after 2s.
        QTimer.singleShot(2000, self._hard_kill_if_running)

    def _hard_kill_if_running(self) -> None:
        if self._process is not None and self._process.state() != QProcess.ProcessState.NotRunning:
            self.status.emit("Forcing kill…")
            self._process.kill()

    # ---- Stdout / events ----

    def _drain_stdout(self) -> None:
        if self._process is None:
            return
        data = bytes(self._process.readAllStandardOutput())
        if not data:
            return
        self._stdout_buf.extend(data)
        # Process complete lines.
        while b"\n" in self._stdout_buf:
            line_bytes, _, rest = self._stdout_buf.partition(b"\n")
            self._stdout_buf = bytearray(rest)
            line = line_bytes.decode("utf-8", errors="replace").rstrip("\r")
            self._handle_line(line)

    def _handle_line(self, line: str) -> None:
        m = _PROGRESS_RE.match(line)
        if m:
            try:
                self.progress.emit(int(m.group(1)))
            except ValueError:
                pass
            return
        if line:
            self.status.emit(line)

    def _on_process_finished(self, exit_code: int, exit_status: QProcess.ExitStatus) -> None:
        # Drain any remaining bytes.
        self._drain_stdout()
        if self._stdout_buf:
            self._handle_line(self._stdout_buf.decode("utf-8", errors="replace"))
            self._stdout_buf.clear()

        if self.is_cancel_requested:
            self.cancelled.emit()
            return
        if exit_code != 0:
            self.failed.emit(f"Agent exited with code {exit_code}.")
            return
        new_outputs = self._scan_outputs() - self._pre_existing_outputs
        if not new_outputs:
            self.failed.emit("Agent reported success but no new .docx appeared in NOTES/AI Output/.")
            return
        # Pick the newest of the new outputs.
        newest = max(new_outputs, key=lambda p: os.path.getmtime(p))
        self.finished.emit(newest)

    def _on_process_error(self, err: QProcess.ProcessError) -> None:
        if self.is_cancel_requested:
            return
        self.failed.emit(f"Process error: {err}")

    # ---- Output discovery ----

    def _scan_outputs(self) -> set[str]:
        out_dir = os.path.join(self.case_path, _AI_OUTPUT_SUBPATH)
        if not os.path.isdir(out_dir):
            return set()
        results: set[str] = set()
        for name in os.listdir(out_dir):
            if name.lower().endswith(".docx"):
                results.add(os.path.join(out_dir, name))
        return results
```

- [ ] **Step 2: Commit**

```bash
git add icharlotte_core/ui/wizard/runners/subprocess_worker.py
git commit -m "$(cat <<'EOF'
feat(wizard): SubprocessWorker — QProcess wrapper around existing agents

Runs Scripts/<agent>.py with --file_number + --file args; pipes stdout
into status/progress signals; detects the newly written .docx in
NOTES/AI Output/ by diffing the directory before/after. Cancel calls
terminate() with a 2s hard-kill fallback.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 5.3: Map each task to its agent script

**Files:**
- Modify: `icharlotte_core/ui/wizard/registry.py` — add `script_name` to each TaskSpec
- Update: `tests/test_wizard/test_registry.py` — assert script_name values

- [ ] **Step 1: Add tests for new field**

Append to `tests/test_wizard/test_registry.py`:

```python
def test_each_task_has_script_name():
    assert get_task("summarize_documents").script_name == "summarize.py"
    assert get_task("summarize_discovery").script_name == "summarize_discovery.py"
    assert get_task("summarize_depositions").script_name == "summarize_deposition.py"
    assert get_task("medical_records").script_name == "med_record.py"
```

- [ ] **Step 2: Run — confirm new test fails**

```bash
pytest tests/test_wizard/test_registry.py::test_each_task_has_script_name -v
```

Expected: `AttributeError: 'TaskSpec' object has no attribute 'script_name'`.

- [ ] **Step 3: Add field**

In `icharlotte_core/ui/wizard/registry.py`, modify the `TaskSpec` dataclass:

```python
@dataclass(frozen=True)
class TaskSpec:
    task_id: str
    title: str
    description: str
    icon_glyph: str
    script_name: str  # name of Scripts/<file>.py
    default_folders: List[str] = field(default_factory=list)
```

Update each registry entry — add `script_name`:

```python
    "summarize_documents": TaskSpec(
        task_id="summarize_documents",
        title="Summarize Documents",
        description="Produce a concise summary of one or more case documents.",
        icon_glyph="\U0001F4C4",
        script_name="summarize.py",
        default_folders=[],
    ),
    "summarize_discovery": TaskSpec(
        task_id="summarize_discovery",
        title="Summarize Discovery",
        description="Summarize discovery responses with structure and citations.",
        icon_glyph="\U0001F4CB",
        script_name="summarize_discovery.py",
        default_folders=["DISCOVERY/RESPONSES", "DISCOVERY"],
    ),
    "summarize_depositions": TaskSpec(
        task_id="summarize_depositions",
        title="Summarize Depositions",
        description="Generate a structured summary of one or more depositions.",
        icon_glyph="\U0001F399",
        script_name="summarize_deposition.py",
        default_folders=["DISCOVERY/TRANSCRIPTS", "DISCOVERY"],
    ),
    "medical_records": TaskSpec(
        task_id="medical_records",
        title="Medical Records Review",
        description="Extract and summarize medical records into a chronology.",
        icon_glyph="\U0001F3E5",
        script_name="med_record.py",
        default_folders=["RECORDS"],
    ),
```

- [ ] **Step 4: Run — confirm all pass**

```bash
pytest tests/test_wizard/test_registry.py -v
```

Expected: 5 tests pass (4 prior + 1 new).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/registry.py tests/test_wizard/test_registry.py
git commit -m "$(cat <<'EOF'
feat(wizard): map each TaskSpec to its Scripts/*.py script_name

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 5.4: Replace fake worker in `TaskTab` with real `SubprocessWorker`

**Files:**
- Modify: `icharlotte_core/ui/wizard/task_tab.py`
- Modify: `iCharlotte.py` — `_open_task_tab` now passes case_path + file_number into TaskTab

- [ ] **Step 1: Update `TaskTab` to take case context**

Open `icharlotte_core/ui/wizard/task_tab.py`. Replace the constructor and worker logic:

```python
class TaskTab(QStackedWidget):
    """Stateful container for one running task. Owns its own worker."""

    closed = Signal()

    def __init__(
        self,
        spec: TaskSpec,
        files: List[str],
        case_path: str,
        file_number: str,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self._spec = spec
        self._files: List[str] = list(files)
        self._case_path = case_path
        self._file_number = file_number
        self._worker = None
        self._worker_thread = None  # reserved if we move to QThread later

        self.settings_page = SettingsPage(spec, files=files)
        self.status_page = StatusPage()
        self.output_page = OutputPage()

        self.addWidget(self.settings_page)
        self.addWidget(self.status_page)
        self.addWidget(self.output_page)

        self.settings_page.proceed_requested.connect(self._on_proceed)
        self.status_page.cancel_requested.connect(self._on_cancel)
        self.output_page.edit_settings_requested.connect(self._on_edit_settings)
        self.output_page.rerun_requested.connect(self._on_rerun)
```

Replace `_start_run` with a real subprocess invocation:

```python
    def _start_run(self, settings_dict: dict) -> None:
        from .runners.subprocess_worker import SubprocessWorker

        self.status_page.on_status(f"Starting {self._spec.title}…")
        self._worker = SubprocessWorker(
            script_name=self._spec.script_name,
            case_path=self._case_path,
            file_number=self._file_number,
            files=self._files,
            settings=settings_dict,
            parent=self,
        )
        self._worker.status.connect(self.status_page.on_status)
        self._worker.progress.connect(self.status_page.on_progress)
        self._worker.finished.connect(self._on_worker_finished)
        self._worker.failed.connect(self._on_worker_failed)
        self._worker.cancelled.connect(self._on_worker_cancelled)
        self._worker.start()

    def _on_worker_finished(self, output_path: str) -> None:
        self._worker = None
        self._show_output(output_path)

    def _on_worker_failed(self, err: str) -> None:
        self._worker = None
        self.status_page.on_status(f"FAILED: {err}")
        self.status_page.cancel_btn.setText("Back to Settings")
        self.status_page.cancel_btn.setEnabled(True)
        try:
            self.status_page.cancel_btn.clicked.disconnect()
        except RuntimeError:
            pass
        self.status_page.cancel_btn.clicked.connect(lambda: self.setCurrentIndex(PAGE_SETTINGS))

    def _on_worker_cancelled(self) -> None:
        self._worker = None
        self.setCurrentIndex(PAGE_SETTINGS)
```

Remove the `_fake_worker_delay_ms` attribute and the `QTimer.singleShot` line entirely.

Also update `_on_cancel`:

```python
    def _on_cancel(self) -> None:
        if self._worker is not None:
            self._worker.cancel()
        else:
            self.setCurrentIndex(PAGE_SETTINGS)
```

- [ ] **Step 2: Update the smoke test to pass case context**

In `tests/test_wizard/test_task_tab.py`, update each `TaskTab(...)` constructor call to include `case_path` and `file_number`:

```python
def test_initial_state_is_settings(qtbot):
    tab = TaskTab(get_task("summarize_documents"), files=["/tmp/x.pdf"], case_path="/tmp/case", file_number="0000.000")
    qtbot.addWidget(tab)
    assert tab.current_page == PAGE_SETTINGS
```

Remove the `test_proceed_transitions_to_status` test entirely — it was Phase 4 fake-worker-specific and no longer applies (Proceed now actually spawns a subprocess).

Keep `test_show_output_transitions` (it doesn't touch the worker).

- [ ] **Step 3: Run — confirm pass**

```bash
pytest tests/test_wizard/test_task_tab.py -v
```

Expected: 2 tests pass.

- [ ] **Step 4: Update `_open_task_tab` to pass case context**

In `iCharlotte.py`, update the `TaskTab` construction inside `_open_task_tab`:

```python
        task_tab = TaskTab(
            spec=spec,
            files=files,
            case_path=self.case_path,
            file_number=self.file_number,
            parent=self,
        )
```

- [ ] **Step 5: Manual end-to-end smoke**

```bash
python iCharlotte.py
```

1. Open a case with real PDFs in `<case>/NOTES/AI Output/` already populated (or empty).
2. Switch to Wizard tab → click **Summarize Documents** → pick a small PDF → OK.
3. New task tab appears on Settings page → click Proceed → flips to Status page → real log lines stream in from `summarize.py` stdout.
4. After the agent finishes, the tab flips to Output page; the placeholder body shows the .docx path.
5. Click Cancel during a run → button greys, "Cancelling…", subprocess terminates within 2s, tab returns to Settings.

If the agent crashes, the FAILED line should appear with a "Back to Settings" button.

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/ui/wizard/task_tab.py iCharlotte.py tests/test_wizard/test_task_tab.py
git commit -m "$(cat <<'EOF'
feat(wizard): wire real SubprocessWorker into TaskTab

Proceed now spawns the agent specified in TaskSpec.script_name via
QProcess, streams its stdout into the Status page, and lands on the
Output page bound to the newly-written .docx in NOTES/AI Output/.
Cancel terminates the subprocess.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

# Phase 6 — Persistence + case-switch lifecycle

After Phase 6, open task tabs survive case switches and app restarts. Mid-run tabs reset to Settings on next session per the spec.

### Task 6.1: `WizardStatePersistence` (TDD)

**Files:**
- Create: `icharlotte_core/ui/wizard/persistence.py`
- Create: `tests/test_wizard/test_persistence.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_wizard/test_persistence.py`:

```python
"""Tests for WizardStatePersistence."""
import json
import os
import pytest

from icharlotte_core.ui.wizard.persistence import WizardStatePersistence


def test_load_missing_file_returns_default(tmp_path):
    p = WizardStatePersistence(str(tmp_path))
    data = p.load()
    assert data["version"] == 1
    assert data["open_tabs"] == []
    assert data["recent_tasks"] == []


def test_save_then_load_roundtrips(tmp_path):
    p = WizardStatePersistence(str(tmp_path))
    p.set_open_tabs([
        {"task_id": "summarize_documents", "instance_suffix": "", "files": ["a.pdf"],
         "settings": {}, "page": "settings", "output_path": None},
    ])
    p.add_recent_task({
        "task_id": "summarize_documents", "title": "Summarize Documents",
        "files": ["a.pdf"], "settings": {},
        "output_path": "NOTES/AI Output/x.docx",
        "completed_at": "2026-05-15T10:42:00",
    })
    p.save()

    p2 = WizardStatePersistence(str(tmp_path))
    data = p2.load()
    assert len(data["open_tabs"]) == 1
    assert data["open_tabs"][0]["task_id"] == "summarize_documents"
    assert len(data["recent_tasks"]) == 1


def test_recent_tasks_capped_at_20(tmp_path):
    p = WizardStatePersistence(str(tmp_path))
    for i in range(25):
        p.add_recent_task({
            "task_id": "summarize_documents", "title": f"Run {i}",
            "files": [], "settings": {}, "output_path": "x.docx",
            "completed_at": f"2026-05-15T{i:02d}:00:00",
        })
    p.save()
    data = WizardStatePersistence(str(tmp_path)).load()
    assert len(data["recent_tasks"]) == 20
    # Newest first.
    titles = [t["title"] for t in data["recent_tasks"]]
    assert titles[0] == "Run 24"
    assert titles[-1] == "Run 5"


def test_atomic_write_uses_tmp_then_rename(tmp_path):
    p = WizardStatePersistence(str(tmp_path))
    p.set_open_tabs([])
    p.save()
    state_file = os.path.join(str(tmp_path), ".icharlotte", "wizard_state.json")
    assert os.path.exists(state_file)
    # tmp file should not be lingering.
    tmp_file = state_file + ".tmp"
    assert not os.path.exists(tmp_file)


def test_corrupt_file_falls_back_to_default(tmp_path):
    folder = os.path.join(str(tmp_path), ".icharlotte")
    os.makedirs(folder)
    with open(os.path.join(folder, "wizard_state.json"), "w") as f:
        f.write("{ not json")
    p = WizardStatePersistence(str(tmp_path))
    data = p.load()
    assert data["open_tabs"] == []
    assert data["recent_tasks"] == []


def test_readme_created_on_first_save(tmp_path):
    p = WizardStatePersistence(str(tmp_path))
    p.set_open_tabs([])
    p.save()
    readme = os.path.join(str(tmp_path), ".icharlotte", "README.txt")
    assert os.path.exists(readme)
```

- [ ] **Step 2: Run — confirm failure**

```bash
pytest tests/test_wizard/test_persistence.py -v
```

Expected: `ModuleNotFoundError`.

- [ ] **Step 3: Implement**

Create `icharlotte_core/ui/wizard/persistence.py`:

```python
"""WizardStatePersistence — per-case JSON store for wizard open tabs + history.

Stored at `<case_root>/.icharlotte/wizard_state.json`. Atomic writes via .tmp +
os.replace. Recent tasks capped at 20 (newest first).
"""
import json
import os
from typing import Any


SCHEMA_VERSION = 1
_RECENT_CAP = 20

_README_TEXT = (
    "This folder stores iCharlotte app state for this case.\n"
    "Files here are managed by the application — do not edit manually.\n"
)


class WizardStatePersistence:
    def __init__(self, case_root: str):
        self.case_root = case_root
        self._data: dict[str, Any] | None = None

    # ---- Path helpers ----

    @property
    def folder(self) -> str:
        return os.path.join(self.case_root, ".icharlotte")

    @property
    def state_path(self) -> str:
        return os.path.join(self.folder, "wizard_state.json")

    @property
    def readme_path(self) -> str:
        return os.path.join(self.folder, "README.txt")

    # ---- Load / save ----

    def load(self) -> dict:
        if self._data is not None:
            return self._data
        if not os.path.isfile(self.state_path):
            self._data = self._default()
            return self._data
        try:
            with open(self.state_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
        except (json.JSONDecodeError, OSError):
            self._data = self._default()
            return self._data
        # Tolerate missing keys.
        self._data = self._default()
        if isinstance(raw, dict):
            if isinstance(raw.get("open_tabs"), list):
                self._data["open_tabs"] = raw["open_tabs"]
            if isinstance(raw.get("recent_tasks"), list):
                self._data["recent_tasks"] = raw["recent_tasks"][:_RECENT_CAP]
        return self._data

    def save(self) -> None:
        data = self.load()
        os.makedirs(self.folder, exist_ok=True)
        if not os.path.exists(self.readme_path):
            try:
                with open(self.readme_path, "w", encoding="utf-8") as f:
                    f.write(_README_TEXT)
            except OSError:
                pass
        tmp_path = self.state_path + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        os.replace(tmp_path, self.state_path)

    # ---- Public API ----

    def set_open_tabs(self, tabs: list[dict]) -> None:
        d = self.load()
        d["open_tabs"] = list(tabs)

    def get_open_tabs(self) -> list[dict]:
        return list(self.load().get("open_tabs", []))

    def add_recent_task(self, entry: dict) -> None:
        d = self.load()
        d.setdefault("recent_tasks", []).insert(0, entry)
        if len(d["recent_tasks"]) > _RECENT_CAP:
            d["recent_tasks"] = d["recent_tasks"][:_RECENT_CAP]

    def get_recent_tasks(self) -> list[dict]:
        return list(self.load().get("recent_tasks", []))

    # ---- Internals ----

    def _default(self) -> dict:
        return {"version": SCHEMA_VERSION, "open_tabs": [], "recent_tasks": []}
```

- [ ] **Step 4: Run — confirm pass**

```bash
pytest tests/test_wizard/test_persistence.py -v
```

Expected: 6 tests pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/persistence.py tests/test_wizard/test_persistence.py
git commit -m "$(cat <<'EOF'
feat(wizard): WizardStatePersistence per-case JSON store

Stored under <case>/.icharlotte/wizard_state.json with atomic writes.
Recent tasks capped at 20 newest-first. Tolerates corrupt or missing
files by falling back to defaults. Auto-creates README on first save.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 6.2: Snapshot/restore plumbing in `MainWindow`

**Files:**
- Modify: `iCharlotte.py` — add `_snapshot_open_task_tabs()`, `_remove_all_task_tabs()`, `_restore_task_tabs_for_case()`; wire into `load_case_by_number()` and the close event

- [ ] **Step 1: Add helpers to `MainWindow`**

Add these methods to `MainWindow` (place after `_apply_mode_visibility`):

```python
    # --- Wizard Mode: per-case task-tab snapshot/restore ---

    def _iter_task_tabs(self) -> list[tuple[int, "QWidget"]]:
        """Return (index, widget) for every TaskTab currently in self.tabs."""
        out = []
        for i in range(self.tabs.count()):
            w = self.tabs.widget(i)
            if w is not None and w.property("wizard_task_id") is not None:
                out.append((i, w))
        return out

    def _relpath_under(self, root: str, path: str) -> str:
        try:
            return os.path.relpath(path, root)
        except ValueError:
            return path

    def _snapshot_open_task_tabs(self) -> list[dict]:
        """Build the persistence-ready snapshot of currently-open task tabs."""
        if not self.case_path:
            return []
        snapshots = []
        for _, tab in self._iter_task_tabs():
            # Determine page label.
            page_idx = tab.currentIndex()
            if page_idx == 1:  # PAGE_STATUS — cancel and store as settings
                if getattr(tab, "_worker", None) is not None:
                    try:
                        tab._worker.cancel()
                    except Exception:
                        pass
                page = "settings"
            elif page_idx == 2:  # PAGE_OUTPUT
                page = "output"
            else:
                page = "settings"

            files_rel = [self._relpath_under(self.case_path, f) for f in tab.files]
            output_path = tab.output_page.output_path if page == "output" else None
            output_path_rel = self._relpath_under(self.case_path, output_path) if output_path else None

            snapshots.append({
                "task_id": tab.spec.task_id,
                "instance_suffix": tab.property("wizard_instance_suffix") or "",
                "files": files_rel,
                "settings": tab.settings_page.to_dict(),
                "page": page,
                "output_path": output_path_rel,
            })
        return snapshots

    def _remove_all_task_tabs(self) -> None:
        # Iterate in reverse so indices stay stable.
        for idx, widget in reversed(self._iter_task_tabs()):
            self.tabs.removeTab(idx)
            widget.deleteLater()

    def _save_wizard_state_for_current_case(self) -> None:
        if not self.case_path:
            return
        from icharlotte_core.ui.wizard.persistence import WizardStatePersistence
        p = WizardStatePersistence(self.case_path)
        p.set_open_tabs(self._snapshot_open_task_tabs())
        p.save()

    def _restore_task_tabs_for_case(self) -> None:
        if not self.case_path:
            return
        from icharlotte_core.ui.wizard.persistence import WizardStatePersistence
        from icharlotte_core.ui.wizard.registry import get_task, TASK_REGISTRY
        from icharlotte_core.ui.wizard.task_tab import TaskTab, PAGE_OUTPUT, PAGE_SETTINGS

        p = WizardStatePersistence(self.case_path)
        for entry in p.get_open_tabs():
            task_id = entry.get("task_id")
            if task_id not in TASK_REGISTRY:
                continue
            spec = get_task(task_id)
            files_abs = [
                f if os.path.isabs(f) else os.path.join(self.case_path, f)
                for f in entry.get("files", [])
            ]
            tab = TaskTab(
                spec=spec,
                files=files_abs,
                case_path=self.case_path,
                file_number=self.file_number,
                parent=self,
            )
            suffix = entry.get("instance_suffix", "") or ""
            tab.setProperty("wizard_task_id", spec.task_id)
            tab.setProperty("wizard_instance_suffix", suffix)
            title = f"{spec.title} {suffix}".strip()

            # Restore settings dict if present.
            settings_dict = entry.get("settings") or {}
            try:
                tab.settings_page.from_dict(settings_dict)
            except Exception:
                pass

            self.tabs.addTab(tab, title)

            # Restore page.
            page = entry.get("page", "settings")
            if page == "output":
                out_rel = entry.get("output_path")
                out_abs = os.path.join(self.case_path, out_rel) if out_rel else None
                if out_abs and os.path.exists(out_abs):
                    tab.output_page.load_output(out_abs)
                    tab.setCurrentIndex(PAGE_OUTPUT)
                else:
                    tab.setCurrentIndex(PAGE_SETTINGS)
            else:
                tab.setCurrentIndex(PAGE_SETTINGS)
        self._hide_fixed_close_buttons()
```

- [ ] **Step 2: Wire into `load_case_by_number`**

Modify `MainWindow.load_case_by_number` at `iCharlotte.py:1349`. Replace the existing method body. The new version:

```python
    def load_case_by_number(self, file_number):
        log_debug(f"load_case_by_number: switching to {file_number}")
        new_path = get_case_path(file_number)
        if not new_path:
            QMessageBox.critical(self, "Error", f"Could not find case directory for {file_number}")
            return

        log_debug(f"load_case_by_number: path={new_path}")
        # ---- WIZARD: snapshot current case's task tabs, then remove them ----
        try:
            self._save_wizard_state_for_current_case()
        except Exception as e:
            log_event(f"[wizard] snapshot failed: {e}")
        self._remove_all_task_tabs()

        self.save_status_history()
        if hasattr(self, 'chat_tab') and self.chat_tab:
            self.chat_tab.save_current_state()
        self.file_number = file_number
        self.case_path = new_path
        self._update_window_title()
        self.populate_tree()
        self.clear_all_status()

        for btn in self.agent_buttons.values():
            btn.set_running(False)
        for script, case_num in self.running_agents.items():
            if case_num == file_number and script in self.agent_buttons:
                self.agent_buttons[script].set_running(True)

        self.load_status_history()

        # Activate the appropriate tab for the current mode.
        if self.mode_controller.is_wizard:
            wizard_idx = self._index_of_tab("Wizard")
            if wizard_idx >= 0:
                self.tabs.setCurrentIndex(wizard_idx)
        else:
            case_view_idx = self._index_of_tab("Case View")
            if case_view_idx >= 0:
                self.tabs.setCurrentIndex(case_view_idx)

        if hasattr(self, 'index_tab'):
            self.index_tab.load_data(self.file_number)
        if hasattr(self, 'chat_tab'):
            self.chat_tab.load_case(self.file_number)
        if hasattr(self, 'liability_tab'):
            self.liability_tab.reset_state()
        if hasattr(self, 'email_tab'):
            self.email_tab.search_bar.clear()
            self.email_tab.check_db_init()
            self.email_tab.perform_search()
        if hasattr(self, 'email_update_tab'):
            self.email_update_tab.on_case_changed(file_number)
        if hasattr(self, 'deposition_tab'):
            self.deposition_tab.load_case(file_number)
        if hasattr(self, 'discovery_tab'):
            self.discovery_tab.load_case(file_number)

        # ---- WIZARD: restore the new case's task tabs ----
        try:
            self._restore_task_tabs_for_case()
        except Exception as e:
            log_event(f"[wizard] restore failed: {e}")

        log_event(f"Switched to case {self.file_number}")
```

Add the `_index_of_tab` helper:

```python
    def _index_of_tab(self, tab_text: str) -> int:
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == tab_text:
                return i
        return -1
```

- [ ] **Step 3: Wire into the close event**

Find `MainWindow.closeEvent` (if it doesn't exist, search `def closeEvent` — it may live alongside other event handlers). If absent, add it:

```python
    def closeEvent(self, event):
        try:
            self._save_wizard_state_for_current_case()
        except Exception as e:
            log_event(f"[wizard] close-save failed: {e}")
        super().closeEvent(event)
```

If `closeEvent` already exists, **add** the wizard-save block at the top of its body (before any existing logic).

- [ ] **Step 4: Initial restore on startup**

In `MainWindow.__init__`, after the case is loaded (after `self.load_status_history()` at line 450), add:

```python
        # Restore wizard task tabs for the initial case (if any).
        try:
            self._restore_task_tabs_for_case()
        except Exception as e:
            log_event(f"[wizard] startup restore failed: {e}")
```

- [ ] **Step 5: Manual end-to-end test**

```bash
python iCharlotte.py
```

1. Open a case. Switch to Wizard. Open two task tabs (e.g., Summarize Documents and Medical Records). Leave Summarize Documents on Settings page; let Medical Records finish so it lands on Output page.
2. Switch to a different case via Master List → double-click. The previous case's task tabs disappear immediately. Switch back → both reappear, Settings on its Settings page, Medical Records on its Output page.
3. Close the app. Re-open `python iCharlotte.py`. The same case loads with both tabs still open in the same state.
4. Verify the JSON exists: `cat "<case>/.icharlotte/wizard_state.json"` — should contain both entries.

- [ ] **Step 6: Commit**

```bash
git add iCharlotte.py
git commit -m "$(cat <<'EOF'
feat(wizard): per-case task-tab snapshot/restore across case switches

load_case_by_number now (1) snapshots & removes current case's task
tabs, (2) loads new case, (3) restores the new case's task tabs from
<case>/.icharlotte/wizard_state.json. Status-page tabs are cancelled
on snapshot and restored as Settings. closeEvent also persists state.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

# Phase 7 — Recent Tasks history + Reopen

After Phase 7, completed task runs are recorded into `recent_tasks`, displayed at the bottom of the Wizard tab, and reopenable on the Output page.

### Task 7.1: Emit completion events to persistence

**Files:**
- Modify: `icharlotte_core/ui/wizard/task_tab.py` — emit a `task_completed(entry_dict)` signal alongside the existing `_on_worker_finished`
- Modify: `iCharlotte.py` — listen for `task_completed`, append to persistence

- [ ] **Step 1: Add signal to `TaskTab`**

In `icharlotte_core/ui/wizard/task_tab.py`, add a new signal at the class level (near `closed = Signal()`):

```python
    task_completed = Signal(dict)  # recent-tasks entry dict
```

Update `_on_worker_finished`:

```python
    def _on_worker_finished(self, output_path: str) -> None:
        from datetime import datetime
        self._worker = None
        entry = {
            "task_id": self._spec.task_id,
            "title": self._spec.title,
            "files": list(self._files),
            "settings": self.settings_page.to_dict(),
            "output_path": output_path,
            "completed_at": datetime.now().isoformat(timespec="seconds"),
        }
        self.task_completed.emit(entry)
        self._show_output(output_path)
```

- [ ] **Step 2: Listen from `MainWindow`**

In `iCharlotte.py`, find where `task_tab = TaskTab(...)` is created inside `_open_task_tab` and in `_restore_task_tabs_for_case`. After each construction (and signal-wiring), connect:

```python
        task_tab.task_completed.connect(self._on_task_completed)
```

Add the slot to `MainWindow`:

```python
    def _on_task_completed(self, entry: dict) -> None:
        if not self.case_path:
            return
        from icharlotte_core.ui.wizard.persistence import WizardStatePersistence
        # Store files & output_path as case-relative.
        entry = dict(entry)
        entry["files"] = [self._relpath_under(self.case_path, f) for f in entry.get("files", [])]
        if entry.get("output_path"):
            entry["output_path"] = self._relpath_under(self.case_path, entry["output_path"])
        p = WizardStatePersistence(self.case_path)
        p.add_recent_task(entry)
        p.save()
        # Tell the Wizard tab to refresh its Recent Tasks list.
        if hasattr(self, "wizard_tab") and self.wizard_tab is not None:
            self.wizard_tab.refresh_recent_tasks(p.get_recent_tasks())
```

- [ ] **Step 3: Commit**

```bash
git add icharlotte_core/ui/wizard/task_tab.py iCharlotte.py
git commit -m "$(cat <<'EOF'
feat(wizard): record completed runs into recent_tasks history

TaskTab emits task_completed(entry) on successful finish; MainWindow
appends to <case>/.icharlotte/wizard_state.json recent_tasks (newest
first, capped at 20).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 7.2: Recent Tasks UI on the Wizard tab

**Files:**
- Modify: `icharlotte_core/ui/wizard/wizard_tab.py` — add Recent Tasks section + `refresh_recent_tasks()` slot + `reopen_requested(entry)` signal

- [ ] **Step 1: Update `WizardTab`**

Edit `icharlotte_core/ui/wizard/wizard_tab.py`. Replace the `_build_ui` body to include a Recent Tasks section, and add the new slot + signal:

```python
"""WizardTab — header + grid of TaskCards + Recent Tasks list."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from .registry import list_tasks
from .task_card import TaskCard


_CARDS_PER_ROW = 3


class WizardTab(QWidget):
    task_requested = Signal(str)            # task_id
    reopen_requested = Signal(dict)         # recent-tasks entry

    def __init__(self, parent=None):
        super().__init__(parent)
        self.cards: list[TaskCard] = []
        self._recent_layout: QVBoxLayout | None = None
        self._build_ui()

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(40, 32, 40, 32)
        outer.setSpacing(24)

        header = QLabel("What would you like to do?")
        header.setStyleSheet("font-size: 22px; font-weight: 400; color: #1a1a1a;")
        outer.addWidget(header)

        # Card grid
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        container = QWidget()
        grid = QGridLayout(container)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(20)
        grid.setVerticalSpacing(20)
        grid.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        for idx, spec in enumerate(list_tasks()):
            card = TaskCard(spec, parent=container)
            card.clicked.connect(self.task_requested.emit)
            row, col = divmod(idx, _CARDS_PER_ROW)
            grid.addWidget(card, row, col)
            self.cards.append(card)
        scroll.setWidget(container)
        outer.addWidget(scroll, 1)

        # Divider
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("color: #e0e0e0;")
        outer.addWidget(line)

        # Recent Tasks
        recent_label = QLabel("Recent Tasks")
        recent_label.setStyleSheet("font-size: 14px; font-weight: 600; color: #333;")
        outer.addWidget(recent_label)

        self._recent_layout = QVBoxLayout()
        self._recent_layout.setSpacing(6)
        outer.addLayout(self._recent_layout)
        self._render_recent_empty_state()

    def _render_recent_empty_state(self):
        self._clear_recent_layout()
        empty = QLabel("No completed tasks for this case yet.")
        empty.setStyleSheet("color: #999; font-style: italic;")
        self._recent_layout.addWidget(empty)

    def _clear_recent_layout(self):
        if self._recent_layout is None:
            return
        while self._recent_layout.count():
            item = self._recent_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def refresh_recent_tasks(self, entries: list[dict]):
        """Update the Recent Tasks list."""
        if self._recent_layout is None:
            return
        if not entries:
            self._render_recent_empty_state()
            return
        self._clear_recent_layout()
        for entry in entries:
            row = self._build_recent_row(entry)
            self._recent_layout.addWidget(row)

    def _build_recent_row(self, entry: dict) -> QWidget:
        w = QFrame()
        w.setStyleSheet("QFrame { border-bottom: 1px solid #f0f0f0; padding: 4px 0; }")
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(12)

        title = entry.get("title", entry.get("task_id", "Unknown"))
        ts = entry.get("completed_at", "")
        label = QLabel(f"• {title}  —  {ts}")
        label.setStyleSheet("font-size: 12px; color: #333;")
        h.addWidget(label, 1)

        out_path = entry.get("output_path") or ""
        if out_path:
            label.setToolTip(out_path)

        btn = QPushButton("Reopen")
        btn.setFixedHeight(26)
        btn.setStyleSheet("padding: 0 12px;")
        btn.clicked.connect(lambda _=False, e=entry: self.reopen_requested.emit(e))
        h.addWidget(btn)

        return w
```

- [ ] **Step 2: Wire signal in `MainWindow`**

In `iCharlotte.py`, where `WizardTab` is added (Task 2.4 location), connect the new signal:

```python
        self.wizard_tab.reopen_requested.connect(self._on_reopen_recent_task)
```

And add the slot:

```python
    def _on_reopen_recent_task(self, entry: dict) -> None:
        from icharlotte_core.ui.wizard.registry import get_task, TASK_REGISTRY
        from icharlotte_core.ui.wizard.task_tab import TaskTab, PAGE_OUTPUT, PAGE_SETTINGS
        from icharlotte_core.ui.wizard.instance_naming import next_instance_suffix

        task_id = entry.get("task_id")
        if task_id not in TASK_REGISTRY:
            QMessageBox.warning(self, "Unknown task", f"This case references an unknown task: {task_id!r}")
            return
        spec = get_task(task_id)

        out_rel = entry.get("output_path") or ""
        out_abs = os.path.join(self.case_path, out_rel) if out_rel and self.case_path else out_rel
        files = [
            os.path.join(self.case_path, f) if self.case_path and not os.path.isabs(f) else f
            for f in entry.get("files", [])
        ]

        existing_titles = [self.tabs.tabText(i) for i in range(self.tabs.count())]
        suffix = next_instance_suffix(spec.title, existing_titles)
        title = f"{spec.title} {suffix}".strip()

        task_tab = TaskTab(
            spec=spec,
            files=files,
            case_path=self.case_path,
            file_number=self.file_number,
            parent=self,
        )
        task_tab.setProperty("wizard_task_id", spec.task_id)
        task_tab.setProperty("wizard_instance_suffix", suffix)
        try:
            task_tab.settings_page.from_dict(entry.get("settings") or {})
        except Exception:
            pass
        new_index = self.tabs.addTab(task_tab, title)
        task_tab.task_completed.connect(self._on_task_completed)

        if out_abs and os.path.exists(out_abs):
            task_tab.output_page.load_output(out_abs)
            task_tab.setCurrentIndex(PAGE_OUTPUT)
        else:
            QMessageBox.information(
                self,
                "Output missing",
                f"The saved output file no longer exists.\nYou can re-run with the saved settings.",
            )
            task_tab.setCurrentIndex(PAGE_SETTINGS)

        self.tabs.setCurrentIndex(new_index)
        self._hide_fixed_close_buttons()
```

- [ ] **Step 3: Refresh Recent Tasks on case load**

In `load_case_by_number()` (Task 6.2 version), after `_restore_task_tabs_for_case()`, add:

```python
        # Refresh Wizard tab's Recent Tasks list for the new case.
        if hasattr(self, "wizard_tab") and self.wizard_tab is not None:
            from icharlotte_core.ui.wizard.persistence import WizardStatePersistence
            try:
                p = WizardStatePersistence(self.case_path)
                self.wizard_tab.refresh_recent_tasks(p.get_recent_tasks())
            except Exception as e:
                log_event(f"[wizard] refresh recent_tasks failed: {e}")
```

Also call the same block on initial startup, right after the existing `_restore_task_tabs_for_case` call in `__init__`.

- [ ] **Step 4: Manual verification**

```bash
python iCharlotte.py
```

1. Open a case. Run a task to completion. Switch back to the Wizard tab → Recent Tasks shows the completed run with timestamp + [Reopen].
2. Click [Reopen] → a new task tab opens directly on the Output page bound to the saved .docx.
3. Manually delete the .docx from disk → click [Reopen] again → info popup appears, and the new tab lands on Settings.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/wizard_tab.py iCharlotte.py
git commit -m "$(cat <<'EOF'
feat(wizard): Recent Tasks list + Reopen flow

WizardTab now shows up to 20 most-recent completed runs for the
current case, each with a [Reopen] button that creates a task tab
directly on the Output page. Missing output files gracefully fall
back to Settings with a notice.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

# Phase 8 — Output Page editor (mammoth + python-docx round-trip)

After Phase 8, the Output Page renders the .docx with formatting via mammoth and supports Save (writes a fresh .docx), Open in Word, Copy All, Re-run, and Edit Settings & Re-run.

### Task 8.1: `.docx` → HTML loader (mammoth) + tests

**Files:**
- Create: `icharlotte_core/ui/wizard/docx_io.py`
- Create: `tests/test_wizard/test_docx_io.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_wizard/test_docx_io.py`:

```python
"""Tests for docx <-> HTML helpers used by the Output Page."""
import os
import pytest

from icharlotte_core.ui.wizard.docx_io import load_docx_as_html, save_qtextdocument_as_docx


def test_load_docx_as_html_returns_html(tmp_path):
    # Build a tiny .docx using python-docx for the test.
    from docx import Document
    p = tmp_path / "hello.docx"
    doc = Document()
    doc.add_heading("Title", level=1)
    doc.add_paragraph("Hello, ").add_run("world").bold = True
    doc.save(str(p))

    html = load_docx_as_html(str(p))
    assert "<h1" in html.lower() or "<h1>" in html.lower()
    assert "world" in html
    assert "<strong>" in html.lower() or "<b>" in html.lower()


def test_save_qtextdocument_as_docx_roundtrips_basic_text(tmp_path, qtbot):
    pytest.importorskip("pytestqt")
    from PySide6.QtGui import QTextDocument
    from docx import Document

    qdoc = QTextDocument()
    qdoc.setHtml("<h1>Header</h1><p>Hello <b>bold</b> world.</p>")
    out_path = str(tmp_path / "out.docx")
    save_qtextdocument_as_docx(qdoc, out_path)
    assert os.path.exists(out_path)
    # Re-read with python-docx and check basic content.
    d = Document(out_path)
    all_text = "\n".join(p.text for p in d.paragraphs)
    assert "Header" in all_text
    assert "bold" in all_text
    assert "Hello" in all_text and "world" in all_text
```

- [ ] **Step 2: Run — confirm failure**

```bash
pytest tests/test_wizard/test_docx_io.py -v
```

Expected: `ModuleNotFoundError`.

- [ ] **Step 3: Implement**

Create `icharlotte_core/ui/wizard/docx_io.py`:

```python
"""Convert between .docx and the QTextEdit HTML model used by the Output Page.

Forward (.docx → HTML): mammoth's `convert_to_html`.
Reverse (QTextDocument → .docx): walk QTextBlocks and emit python-docx
paragraphs. We capture: heading levels (Heading 1/2/3 paragraph styles
or HTML <hN>), bold, italic, underline, and bullet/numbered lists. Tables,
images, and other complex structures may render approximately and may
be dropped on save (see spec known limitations).
"""
import re

import mammoth
from docx import Document
from docx.shared import Pt
from PySide6.QtGui import QTextDocument, QTextBlock, QTextCharFormat, QTextBlockFormat


_HEADING_STYLE_RE = re.compile(r"^heading\s+(\d+)$", re.IGNORECASE)


def load_docx_as_html(path: str) -> str:
    """Convert a .docx file to a self-contained HTML string."""
    with open(path, "rb") as f:
        result = mammoth.convert_to_html(f)
    return result.value


def save_qtextdocument_as_docx(qdoc: QTextDocument, out_path: str) -> None:
    """Write a python-docx Document mirroring qdoc's block/inline structure.

    Limitations:
      - Tables, images, embedded objects from the editor are not preserved.
      - Bullet / numbered lists fall back to plain paragraphs (python-docx
        list-style is template-dependent and we don't carry a template here).
    """
    document = Document()
    block: QTextBlock = qdoc.begin()
    while block.isValid():
        text = block.text()
        level = _detect_heading_level(block)
        if level is not None:
            para = document.add_heading(text, level=level)
        else:
            para = document.add_paragraph()
            _write_block_runs(block, para)
        # Spacing — preserve "blank paragraph" feel without touching styles.
        block = block.next()
    document.save(out_path)


def _detect_heading_level(block: QTextBlock) -> int | None:
    """Detect heading level from QTextBlockFormat properties (set by setHtml on <hN>)."""
    fmt: QTextBlockFormat = block.blockFormat()
    style_name = fmt.property(QTextBlockFormat.UserProperty + 1)  # may be None
    if isinstance(style_name, str):
        m = _HEADING_STYLE_RE.match(style_name)
        if m:
            try:
                lvl = int(m.group(1))
                return max(1, min(9, lvl))
            except ValueError:
                return None
    # Fall back to heading-detection by paragraph format heading level
    # (Qt 6 has `headingLevel()` on QTextBlockFormat).
    if hasattr(fmt, "headingLevel"):
        lvl = fmt.headingLevel()
        if lvl and lvl > 0:
            return min(lvl, 9)
    return None


def _write_block_runs(block: QTextBlock, para) -> None:
    """Walk inline fragments of `block` and emit python-docx runs with formatting."""
    it = block.begin()
    while not it.atEnd():
        frag = it.fragment()
        if frag.isValid():
            text = frag.text()
            char_fmt: QTextCharFormat = frag.charFormat()
            run = para.add_run(text)
            if char_fmt.fontWeight() >= 600:
                run.bold = True
            if char_fmt.fontItalic():
                run.italic = True
            if char_fmt.fontUnderline():
                run.underline = True
            # Font size (in points), if set.
            size_pt = char_fmt.fontPointSize()
            if size_pt and size_pt > 0:
                run.font.size = Pt(size_pt)
        it += 1
```

- [ ] **Step 4: Run — confirm pass**

```bash
pytest tests/test_wizard/test_docx_io.py -v
```

Expected: 2 tests pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/docx_io.py tests/test_wizard/test_docx_io.py
git commit -m "$(cat <<'EOF'
feat(wizard): mammoth-based .docx <-> HTML helpers for Output Page

load_docx_as_html() uses mammoth for clean semantic HTML.
save_qtextdocument_as_docx() walks the QTextDocument and emits
python-docx paragraphs/headings/runs (bold/italic/underline/size).
Limitations documented in module docstring.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 8.2: Wire mammoth output into `OutputPage` + action buttons

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/output_page.py` — replace `load_output` with mammoth rendering; wire Save / Open in Word / Copy All; track dirty state
- Modify: `icharlotte_core/ui/wizard/task_tab.py` — connect Output Page action signals to handlers

- [ ] **Step 1: Replace `OutputPage.load_output` + dirty tracking**

Open `icharlotte_core/ui/wizard/pages/output_page.py`. Replace the entire file with the Phase 8 version:

```python
"""OutputPage — mammoth-rendered editor + action buttons."""
import os

from PySide6.QtCore import Signal
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..docx_io import load_docx_as_html, save_qtextdocument_as_docx


class OutputPage(QWidget):
    rerun_requested = Signal()
    edit_settings_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._output_path: str | None = None
        self._dirty = False

        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(12)

        header = QHBoxLayout()
        self.file_label = QLabel("File: —")
        self.file_label.setStyleSheet("font-weight: 600;")
        header.addWidget(self.file_label, 1)
        self.open_in_word_btn = QPushButton("Open in Word")
        self.open_in_word_btn.clicked.connect(self._on_open_in_word)
        header.addWidget(self.open_in_word_btn)
        outer.addLayout(header)

        self.editor = QTextEdit()
        self.editor.setReadOnly(False)
        self.editor.textChanged.connect(self._on_text_changed)
        outer.addWidget(self.editor, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.copy_all_btn = QPushButton("Copy All")
        self.copy_all_btn.clicked.connect(self._on_copy_all)
        btn_row.addWidget(self.copy_all_btn)
        self.rerun_btn = QPushButton("Re-run")
        self.rerun_btn.clicked.connect(self.rerun_requested.emit)
        btn_row.addWidget(self.rerun_btn)
        self.edit_settings_btn = QPushButton("Edit Settings & Re-run")
        self.edit_settings_btn.clicked.connect(self.edit_settings_requested.emit)
        btn_row.addWidget(self.edit_settings_btn)
        self.save_btn = QPushButton("Save")
        self.save_btn.setStyleSheet(
            "background-color: #1976D2; color: white; font-weight: 600; padding: 6px 18px;"
        )
        self.save_btn.clicked.connect(self._on_save)
        btn_row.addWidget(self.save_btn)
        outer.addLayout(btn_row)

        self._refresh_save_enabled()

    # ---- Public API ----

    @property
    def output_path(self) -> str | None:
        return self._output_path

    @property
    def is_dirty(self) -> bool:
        return self._dirty

    def load_output(self, output_path: str) -> None:
        self._output_path = output_path
        self.file_label.setText(f"File: {os.path.basename(output_path)}")
        if os.path.isfile(output_path) and output_path.lower().endswith(".docx"):
            try:
                html = load_docx_as_html(output_path)
                self.editor.blockSignals(True)
                self.editor.setHtml(html)
                self.editor.blockSignals(False)
            except Exception as e:
                self.editor.setPlainText(f"(Failed to render {output_path}:\n{e})")
        else:
            self.editor.setPlainText(f"(File not found or not a .docx: {output_path})")
        self._dirty = False
        self._refresh_save_enabled()

    # ---- Internals ----

    def _on_text_changed(self) -> None:
        self._dirty = True
        self._refresh_save_enabled()

    def _refresh_save_enabled(self) -> None:
        self.save_btn.setEnabled(self._dirty and self._output_path is not None)

    def _on_save(self) -> None:
        if self._output_path is None:
            return
        try:
            save_qtextdocument_as_docx(self.editor.document(), self._output_path)
            self._dirty = False
            self._refresh_save_enabled()
        except Exception as e:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.critical(self, "Save failed", f"Could not save .docx:\n{e}")

    def _on_copy_all(self) -> None:
        QGuiApplication.clipboard().setText(self.editor.toPlainText())

    def _on_open_in_word(self) -> None:
        if self._output_path is None:
            return
        if self._dirty:
            from PySide6.QtWidgets import QMessageBox
            ans = QMessageBox.question(
                self,
                "Save first?",
                "You have unsaved changes. Save before opening in Word?",
            )
            if ans == QMessageBox.StandardButton.Yes:
                self._on_save()
        try:
            os.startfile(self._output_path)  # Windows
        except Exception as e:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.critical(self, "Open failed", f"Could not open in Word:\n{e}")
```

- [ ] **Step 2: Verify TaskTab wiring is still correct**

The existing connections from Task 4.3 (`output_page.rerun_requested.connect(self._on_rerun)` and `output_page.edit_settings_requested.connect(self._on_edit_settings)`) should still work — confirm by reading `task_tab.py`.

Note: the original `OutputPage` exposed `open_in_word_requested`, `copy_all_requested`, and `save_requested` signals — the new version handles those internally instead. Remove any stale connections in `task_tab.py` that reference those signals (search for `open_in_word_requested`, `copy_all_requested`, `save_requested` and delete them if present).

Run:

```bash
grep -n "open_in_word_requested\|copy_all_requested\|save_requested" icharlotte_core/ui/wizard/task_tab.py
```

Expected: no matches. If any appear, delete those lines.

- [ ] **Step 3: Run all wizard tests**

```bash
pytest tests/test_wizard -v
```

Expected: all tests still pass.

- [ ] **Step 4: Manual end-to-end test**

```bash
python iCharlotte.py
```

1. Open a case → run Summarize Documents on a small PDF → wait for output.
2. Output page renders the generated .docx with headings, bullets, formatting visible.
3. Edit a name in the editor. **Save** button becomes enabled. Click Save. Reopen in Word externally to confirm the change persisted.
4. Click **Open in Word** (with no unsaved changes) → Word opens the file.
5. Make another edit → click Open in Word → "Save first?" dialog appears. Choose Yes → file saves, then opens.
6. Click **Copy All** → paste into Notepad → text appears.
7. Click **Re-run** → confirms (no confirm dialog in Phase 8 yet — fine for now) and starts the run.
8. Click **Edit Settings & Re-run** → flips back to Settings page with files intact.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/output_page.py icharlotte_core/ui/wizard/task_tab.py
git commit -m "$(cat <<'EOF'
feat(wizard): mammoth-rendered Output Page editor with round-trip Save

Output page now loads the agent's .docx via mammoth into a QTextEdit
(editable, formatting preserved). Save writes a fresh .docx via
python-docx (overwrite). Open in Word prompts to save dirty edits
first. Copy All copies plain text; Re-run / Edit Settings & Re-run
trigger the existing handlers in TaskTab.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

# Phase 9 — Manual end-to-end verification

This phase has no code changes. It's a structured walkthrough of every requirement in the spec, performed against a real case folder. **Do not skip — this is the project's mandatory testing step per CLAUDE.md.**

### Task 9.1: End-to-end verification checklist

- [ ] **Setup**

Pick a real test case with files in:
  - `<case>/DISCOVERY/RESPONSES/` (at least one .pdf)
  - `<case>/DISCOVERY/TRANSCRIPTS/` (at least one .pdf)
  - `<case>/RECORDS/` (at least one .pdf)
  - At least one .pdf in the case root.

If `<case>/.icharlotte/wizard_state.json` already exists from prior runs, delete it for a clean baseline.

- [ ] **Verify: Mode toggle + tab visibility**

  1. Launch app. Default = Wizard. Only Master List + Wizard visible.
  2. Click **Advanced Mode**. All Advanced tabs reappear; Wizard hidden.
  3. Click **Wizard Mode**. Reverse.
  4. Restart app → last-selected mode persists.

- [ ] **Verify: Master List flow**

  1. Master List shows the case table.
  2. **Change File** button is **gone** from the corner.
  3. Win+C still opens the FileNumberDialog (and switches to the right tab per current mode).
  4. Double-click a case → loads it; lands on **Wizard tab** (in Wizard Mode) or **Case View** (in Advanced).

- [ ] **Verify: Each of four cards opens with correct default folder**

| Card | Expected default folder |
|---|---|
| Summarize Documents | case root |
| Summarize Discovery | `<case>/DISCOVERY/RESPONSES` |
| Summarize Depositions | `<case>/DISCOVERY/TRANSCRIPTS` |
| Medical Records | `<case>/RECORDS` |

  - Cancel the file dialog → no tab appears.

- [ ] **Verify: Task tab lifecycle**

  1. Pick a file → new task tab appears to the right of Wizard, focused.
  2. Settings page shows file list + placeholder body + Proceed.
  3. Click Proceed → Status page; live log lines stream from the agent.
  4. Wait for completion → Output page renders the .docx.
  5. Confirm `<case>/NOTES/AI Output/<name>.docx` exists on disk.

- [ ] **Verify: Multi-instance**

  1. Click Summarize Documents twice in a row. Second tab is `Summarize Documents (2)`.
  2. Close the first one. Click again → suffix is `(2)` again (lowest unused).

- [ ] **Verify: Cancel**

  1. Start a long task. Click Cancel → button shows "Cancelling…" → within 2s, the tab returns to Settings.

- [ ] **Verify: Output page**

  1. Edit a word in the rendered output. Save button enables. Click Save.
  2. Open the .docx externally in Word → edit is present.
  3. Click Open in Word with dirty edits → prompts to save first.
  4. Click Copy All → paste into Notepad → plain text appears.

- [ ] **Verify: Case switch snapshot/restore**

  1. With several task tabs open (mix of Settings and Output state), double-click a different case.
  2. Case A's task tabs disappear. New case loads on Wizard tab.
  3. Double-click Case A again. Both tabs reappear in their saved states (Status → Settings, Output → Output).
  4. Inspect `<case_A>/.icharlotte/wizard_state.json` — both entries present.

- [ ] **Verify: App restart restore**

  1. Close the app while two task tabs are open.
  2. Relaunch. Same case loads with both tabs restored.

- [ ] **Verify: Recent Tasks**

  1. After a successful run, scroll to the bottom of the Wizard tab → "Recent Tasks" shows the completed run with title + timestamp.
  2. Click [Reopen] → a new tab opens directly on the Output page bound to the same .docx.
  3. Manually delete the .docx → click [Reopen] again → notice appears; tab lands on Settings.

- [ ] **Verify: Word safety rule**

  At no point during testing did the wizard close any pre-existing Word window. Open in Word should attach to / create new Word windows but never call `Quit()` on the user's session.

- [ ] **Step: If everything passes**

Create a summary commit (no code changes) that documents the verification result:

```bash
git commit --allow-empty -m "$(cat <<'EOF'
test(wizard): end-to-end manual verification complete

All spec requirements verified on a real case. See plan task 9.1 for
the executed checklist.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

If anything fails, file an issue or fix inline before proceeding.

---

## Spec coverage matrix (self-review)

| Spec section | Implemented in |
|---|---|
| Goals: dual modes, card grid, per-task tabs, persistence | Phases 1–7 |
| `ModeController` (QSettings, default `wizard`) | Task 1.2 |
| Mode toggle in Master List header | Tasks 1.3, 1.5 |
| Tab visibility per mode | Task 1.6 |
| Remove Change File button (keep Win+C) | Task 1.4 |
| `WizardTab`, "What would you like to do?", card grid | Tasks 2.3, 7.2 |
| `TASK_REGISTRY` with 4 tasks + `default_folders` + `script_name` | Tasks 2.1, 5.3 |
| Pre-Settings file dialog at task-specific default folder | Tasks 3.1, 3.2, 4.6 |
| `TaskTab` Settings → Status → Output state machine | Phase 4 |
| Multi-instance auto-numbered tabs | Task 4.5 |
| Closeable task tabs only (Master List/Wizard pinned) | Task 4.6 |
| Subprocess worker contract + soft cancel + agent shims | Phase 5 |
| `<case>/.icharlotte/wizard_state.json` (atomic, README) | Task 6.1 |
| Snapshot/restore on case switch + app close | Task 6.2 |
| Status-page tabs reset to Settings across sessions | Task 6.2 |
| Recent Tasks list + [Reopen] + missing-file fallback | Tasks 7.1, 7.2 |
| mammoth-rendered Output Page, round-trip Save (overwrite) | Phase 8 |
| Open in Word, Copy All, Re-run, Edit Settings & Re-run | Phase 8 |
| Known limitations (single-folder file picker, placeholder Settings, soft cancel) | Acknowledged in spec; respected by plan |
| Manual end-to-end verification | Phase 9 |
