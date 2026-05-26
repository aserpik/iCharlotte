# Workbench Model Defaults Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move default LLM model settings into the Prompt Engineering Workbench and remove the separate main-window LLM Settings menu item.

**Architecture:** Extract the existing LLM settings form into a reusable `LLMSettingsWidget` hosted inside `PromptsDialog` as a new **Model Defaults** tab. Keep `LLMConfig` and `config/llm_preferences.json` as the only settings source, and refresh the Workbench's selected-agent model panel after defaults are saved or reset.

**Tech Stack:** Python, PySide6, unittest, existing `LLMConfig`, existing Prompt Engineering Workbench classes in `icharlotte_core/ui/dialogs.py`.

---

## Execution Note

The active workspace already has unrelated edits in `iCharlotte.py` and `icharlotte_core/ui/dialogs.py`. If using the current shared workspace, do not commit implementation changes unless you can stage only the new hunks. If using an isolated worktree, run the commit steps exactly as written.

---

## File Structure

- Modify `icharlotte_core/ui/dialogs.py`
  - Add `LLMSettingsWidget(QWidget)`.
  - Move the current `LLMSettingsDialog` settings UI and persistence methods into the widget.
  - Keep `LLMSettingsDialog(QDialog)` as a thin wrapper around the widget for compatibility.
  - Add a **Model Defaults** tab to `PromptsDialog`.
  - Refresh the selected-agent model panel when embedded defaults are saved or reset.

- Modify `iCharlotte.py`
  - Remove `LLMSettingsDialog` from the dialog import.
  - Remove the `LLM Settings` action from the Settings dropdown.
  - Remove `open_settings_dialog()`.

- Create `tests/test_prompts_dialog_model_defaults.py`
  - Test that Workbench exposes **Model Defaults**.
  - Test that the embedded settings widget loads task profile rows.
  - Test that saving the widget uses `LLMConfig.update_task_config()` and emits a refresh signal.
  - Test that `iCharlotte.py` no longer exposes the removed main-window dialog entry point.

---

### Task 1: Add Failing Tests For Workbench Model Defaults

**Files:**
- Create: `tests/test_prompts_dialog_model_defaults.py`

- [ ] **Step 1: Write the failing test file**

Create `tests/test_prompts_dialog_model_defaults.py` with this content:

```python
"""Tests for moving LLM model defaults into the Prompt Engineering Workbench."""

import os
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from PySide6.QtWidgets import QApplication

app = QApplication.instance() or QApplication(sys.argv)

from icharlotte_core.llm_config import AgentConfig, ModelSpec, TaskConfig
from icharlotte_core.ui import dialogs as dialogs_module
from icharlotte_core.ui.dialogs import LLMSettingsWidget, PromptsDialog


class FakeLLMConfig:
    """Small in-memory stand-in for LLMConfig used by the settings widget tests."""

    _instance = None
    latest_instance = None

    def __init__(self):
        FakeLLMConfig.latest_instance = self
        self.updated_agents = []
        self.updated_tasks = []
        self.task_configs = {
            "general": TaskConfig(
                name="general",
                model_sequence=[
                    ModelSpec(provider="Gemini", model="gemini-2.5-flash", max_tokens=4096)
                ],
                max_retries=2,
                timeout_seconds=90,
            )
        }

    def is_provider_available(self, provider):
        return provider == "Gemini"

    def get_all_task_types(self):
        return list(self.task_configs.keys())

    def get_task_config(self, task_type="general"):
        return self.task_configs[task_type]

    def update_task_config(self, task_type, model_sequence, max_retries=None, timeout_seconds=None):
        self.updated_tasks.append((task_type, model_sequence, max_retries, timeout_seconds))
        self.task_configs[task_type] = TaskConfig(
            name=task_type,
            model_sequence=model_sequence,
            max_retries=max_retries or 3,
            timeout_seconds=timeout_seconds or 120,
        )

    def get_agent_info(self, agent_id):
        return {
            "name": agent_id.replace("_", " ").title(),
            "description": f"{agent_id} description",
            "default_task": "general",
        }

    def get_agent_config(self, agent_id):
        return AgentConfig(
            agent_id=agent_id,
            display_name=agent_id,
            model_sequence=[],
            use_default=True,
        )

    def update_agent_config(self, agent_id, model_sequence=None, use_default=None, max_retries=None, timeout_seconds=None):
        self.updated_agents.append((agent_id, model_sequence, use_default, max_retries, timeout_seconds))

    def get_model_sequence(self, task_type="general"):
        return self.task_configs[task_type].model_sequence


class TestWorkbenchModelDefaults(unittest.TestCase):
    def setUp(self):
        FakeLLMConfig.latest_instance = None
        self.config_patch = patch.object(dialogs_module, "LLMConfig", FakeLLMConfig)
        self.config_patch.start()

    def tearDown(self):
        self.config_patch.stop()

    def test_prompts_dialog_contains_model_defaults_tab(self):
        dlg = PromptsDialog()
        try:
            tab_names = [dlg.tabs.tabText(i) for i in range(dlg.tabs.count())]
            self.assertIn("Model Defaults", tab_names)
            self.assertIsInstance(dlg.model_defaults_widget, LLMSettingsWidget)
        finally:
            dlg.deleteLater()

    def test_model_defaults_widget_loads_task_profile_rows(self):
        widget = LLMSettingsWidget()
        try:
            self.assertIn("general", widget.task_widgets)
            rows = widget.task_widgets["general"]["model_rows"]
            self.assertGreaterEqual(len(rows), 1)
            first_row = rows[0]
            self.assertEqual(first_row["provider"].currentText(), "Gemini")
            self.assertEqual(first_row["model"].currentData(), "gemini-2.5-flash")
            self.assertEqual(first_row["max_tokens"].value(), 4096)
            self.assertEqual(widget.task_widgets["general"]["max_retries"].value(), 2)
            self.assertEqual(widget.task_widgets["general"]["timeout"].value(), 90)
        finally:
            widget.deleteLater()

    def test_model_defaults_save_updates_config_and_emits_signal(self):
        widget = LLMSettingsWidget()
        emitted = []
        widget.settings_saved.connect(lambda: emitted.append(True))
        try:
            first_row = widget.task_widgets["general"]["model_rows"][0]
            first_row["provider"].setCurrentText("OpenAI")
            model_index = first_row["model"].findData("gpt-4o-mini")
            self.assertGreaterEqual(model_index, 0)
            first_row["model"].setCurrentIndex(model_index)
            first_row["max_tokens"].setValue(2048)

            with patch.object(dialogs_module.QMessageBox, "information"):
                widget._save_settings()

            self.assertEqual(emitted, [True])
            saved = widget.config.updated_tasks[-1]
            self.assertEqual(saved[0], "general")
            self.assertEqual(saved[1][0].provider, "OpenAI")
            self.assertEqual(saved[1][0].model, "gpt-4o-mini")
            self.assertEqual(saved[1][0].max_tokens, 2048)
        finally:
            widget.deleteLater()

    def test_model_defaults_signals_refresh_selected_agent_panel(self):
        dlg = PromptsDialog()
        try:
            with patch.object(dlg, "_load_agent_model_settings") as reload_settings:
                dlg.model_defaults_widget.settings_saved.emit()
            reload_settings.assert_called_once()

            with patch.object(dlg, "_load_agent_model_settings") as reload_settings:
                dlg.model_defaults_widget.settings_reset.emit()
            reload_settings.assert_called_once()
        finally:
            dlg.deleteLater()

    def test_main_window_no_longer_references_llm_settings_dialog(self):
        source = Path("iCharlotte.py").read_text(encoding="utf-8")
        self.assertNotIn("LLMSettingsDialog", source)
        self.assertNotIn("LLM Settings\", self.open_settings_dialog", source)
        self.assertNotIn("def open_settings_dialog", source)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the new tests and verify they fail for the expected reason**

Run:

```powershell
python -m pytest tests/test_prompts_dialog_model_defaults.py -q
```

Expected result:

```text
ImportError or AttributeError mentioning LLMSettingsWidget
```

The expected failure confirms the tests are checking the new widget and Workbench tab that do not exist yet.

---

### Task 2: Extract The LLM Settings Form Into A Reusable Widget

**Files:**
- Modify: `icharlotte_core/ui/dialogs.py`
- Test: `tests/test_prompts_dialog_model_defaults.py`

- [ ] **Step 1: Rename the existing settings dialog implementation to a widget**

In `icharlotte_core/ui/dialogs.py`, replace the `class LLMSettingsDialog(QDialog):` line and its constructor with this widget class header and constructor. Leave the existing helper methods below it attached to the new class.

```python
class LLMSettingsWidget(QWidget):
    """Reusable widget for configuring LLM model preferences per agent/function."""

    settings_saved = Signal()
    settings_reset = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.config = LLMConfig()

        # Track widgets for agents and task types
        self.agent_widgets = {}
        self.task_widgets = {}

        self._setup_ui()
        self._load_current_settings()
```

- [ ] **Step 2: Update the widget save buttons so the form no longer closes itself**

Inside `LLMSettingsWidget._setup_ui()`, replace the old button block:

```python
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)
```

with no cancel button. The completed button block should be:

```python
        reset_btn = QPushButton("Reset All to Defaults")
        reset_btn.clicked.connect(self._reset_to_defaults)
        btn_layout.addWidget(reset_btn)

        btn_layout.addStretch()

        save_btn = QPushButton("Save")
        save_btn.setStyleSheet("font-weight: bold;")
        save_btn.clicked.connect(self._save_settings)
        btn_layout.addWidget(save_btn)

        layout.addLayout(btn_layout)
```

- [ ] **Step 3: Emit a signal after save instead of accepting a dialog**

In `LLMSettingsWidget._save_settings()`, replace:

```python
            QMessageBox.information(self, "Success", "LLM settings saved successfully.")
            self.accept()
```

with:

```python
            QMessageBox.information(self, "Success", "LLM settings saved successfully.")
            self.settings_saved.emit()
```

- [ ] **Step 4: Emit a signal after reset**

In `LLMSettingsWidget._reset_to_defaults()`, replace:

```python
            # Reload UI
            self._load_current_settings()
            QMessageBox.information(self, "Reset", "Settings reset to defaults.")
```

with:

```python
            # Reload UI
            self._load_current_settings()
            self.settings_reset.emit()
            QMessageBox.information(self, "Reset", "Settings reset to defaults.")
```

- [ ] **Step 5: Add a thin compatibility dialog wrapper**

Immediately after the end of `LLMSettingsWidget._reset_to_defaults()`, add:

```python
class LLMSettingsDialog(QDialog):
    """Compatibility wrapper for the reusable LLM settings widget."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("LLM Settings")
        self.resize(800, 650)

        layout = QVBoxLayout(self)
        self.settings_widget = LLMSettingsWidget(self)
        self.settings_widget.settings_saved.connect(self.accept)
        layout.addWidget(self.settings_widget)

        close_row = QHBoxLayout()
        close_row.addStretch()
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        close_row.addWidget(close_btn)
        layout.addLayout(close_row)

        self.config = self.settings_widget.config
        self.agent_widgets = self.settings_widget.agent_widgets
        self.task_widgets = self.settings_widget.task_widgets
```

- [ ] **Step 6: Run the widget-focused tests and verify the Workbench tab test still fails**

Run:

```powershell
python -m pytest tests/test_prompts_dialog_model_defaults.py -q
```

Expected result:

```text
2 passed, 3 failed
```

The passing tests should be the widget load/save tests. The remaining failures should be the missing Workbench tab, the missing refresh signal wiring, and the still-present main-window reference.

---

### Task 3: Add Model Defaults To The Prompt Engineering Workbench

**Files:**
- Modify: `icharlotte_core/ui/dialogs.py`
- Test: `tests/test_prompts_dialog_model_defaults.py`

- [ ] **Step 1: Add the Model Defaults tab in `PromptsDialog._setup_ui()`**

In `PromptsDialog._setup_ui()`, after the Dashboard tab line:

```python
        self.tabs.addTab(self._create_dashboard_tab(), "Dashboard")
```

add:

```python
        self.model_defaults_widget = LLMSettingsWidget(self)
        self.model_defaults_widget.settings_saved.connect(self._refresh_model_settings_from_defaults)
        self.model_defaults_widget.settings_reset.connect(self._refresh_model_settings_from_defaults)
        self.tabs.addTab(self.model_defaults_widget, "Model Defaults")
```

- [ ] **Step 2: Add the refresh method to `PromptsDialog`**

Add this method near the existing model-settings methods, immediately before `_create_editor_tab()`:

```python
    def _refresh_model_settings_from_defaults(self):
        """Refresh selected-agent model display after default settings change."""
        self.llm_config = LLMConfig()
        self._load_agent_model_settings()
```

- [ ] **Step 3: Run the model-default tests and verify only the main-window test still fails**

Run:

```powershell
python -m pytest tests/test_prompts_dialog_model_defaults.py -q
```

Expected result:

```text
4 passed, 1 failed
```

The remaining failure should be `test_main_window_no_longer_references_llm_settings_dialog`.

---

### Task 4: Remove The Main-Window LLM Settings Entry Point

**Files:**
- Modify: `iCharlotte.py`
- Test: `tests/test_prompts_dialog_model_defaults.py`

- [ ] **Step 1: Remove `LLMSettingsDialog` from the import**

In `iCharlotte.py`, replace:

```python
from icharlotte_core.ui.dialogs import FileNumberDialog, VariablesDialog, PromptsDialog, LLMSettingsDialog
```

with:

```python
from icharlotte_core.ui.dialogs import FileNumberDialog, VariablesDialog, PromptsDialog
```

- [ ] **Step 2: Remove the LLM Settings action and separator from the settings menu**

In `iCharlotte.py`, replace this block:

```python
        self.settings_menu = QMenu(self)
        self.settings_menu.addAction("LLM Settings", self.open_settings_dialog)
        self.settings_menu.addSeparator()
        self.email_monitor_action = self.settings_menu.addAction("Email Monitor")
```

with:

```python
        self.settings_menu = QMenu(self)
        self.email_monitor_action = self.settings_menu.addAction("Email Monitor")
```

- [ ] **Step 3: Remove the now-unused dialog opener method**

Delete this method from `iCharlotte.py`:

```python
    def open_settings_dialog(self):
        """Open the LLM settings dialog."""
        dialog = LLMSettingsDialog(self)
        dialog.exec()
```

- [ ] **Step 4: Run the focused tests and verify they pass**

Run:

```powershell
python -m pytest tests/test_prompts_dialog_model_defaults.py -q
```

Expected result:

```text
5 passed
```

---

### Task 5: Run Regression Checks

**Files:**
- Verify: `iCharlotte.py`
- Verify: `icharlotte_core/ui/dialogs.py`
- Verify: `tests/test_prompts_dialog_loader.py`
- Verify: `tests/test_prompts_dialog_model_defaults.py`

- [ ] **Step 1: Compile the touched Python modules**

Run:

```powershell
python -m py_compile iCharlotte.py icharlotte_core/ui/dialogs.py
```

Expected result:

```text
No output and exit code 0
```

- [ ] **Step 2: Run the new model-default tests**

Run:

```powershell
python -m pytest tests/test_prompts_dialog_model_defaults.py -q
```

Expected result:

```text
5 passed
```

- [ ] **Step 3: Run the existing Prompt Workbench loader tests**

Run:

```powershell
python -m pytest tests/test_prompts_dialog_loader.py -q
```

Expected result:

```text
All tests pass
```

- [ ] **Step 4: Inspect the final touched-file diff**

Run:

```powershell
git diff -- iCharlotte.py icharlotte_core/ui/dialogs.py tests/test_prompts_dialog_model_defaults.py
```

Expected result:

```text
Diff shows only the widget extraction, Workbench tab addition, main-window menu removal, and new tests.
```

- [ ] **Step 5: Commit only in a clean or isolated worktree**

Run this only when the execution workspace has no pre-existing edits in the touched files:

```powershell
git add -- iCharlotte.py icharlotte_core/ui/dialogs.py tests/test_prompts_dialog_model_defaults.py
git commit -m "feat: move llm defaults into prompt workbench"
```

Expected result:

```text
Commit created with only the three touched implementation files.
```
