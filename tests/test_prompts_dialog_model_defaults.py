"""Tests for moving LLM model defaults into the Prompt Engineering Workbench."""

import ast
import os
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from PySide6.QtWidgets import QApplication, QComboBox, QPushButton, QSpinBox

app = QApplication.instance() or QApplication(sys.argv)

from icharlotte_core.llm_config import AgentConfig, ModelSpec, TaskConfig
from icharlotte_core.ui import dialogs as dialogs_module
from icharlotte_core.ui.dialogs import LLMSettingsWidget, PromptsDialog


def _child(widget, cls, name):
    child = widget.findChild(cls, name)
    if child is None:
        raise AssertionError(f"Expected {cls.__name__} named {name!r}")
    return child


def _tab_names(tab_widget):
    return [tab_widget.tabText(i) for i in range(tab_widget.count())]


def _source_tree(path):
    return ast.parse(path.read_text(encoding="utf-8"), filename=str(path))


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
            self.assertIn("Model Defaults", _tab_names(dlg.tabs))
            self.assertIsInstance(dlg.model_defaults_widget, LLMSettingsWidget)
        finally:
            dlg.deleteLater()

    def test_model_defaults_widget_loads_task_profile_rows(self):
        widget = LLMSettingsWidget()
        try:
            provider = _child(widget, QComboBox, "task_provider_general_0")
            model = _child(widget, QComboBox, "task_model_general_0")
            tokens = _child(widget, QSpinBox, "task_tokens_general_0")
            retries = _child(widget, QSpinBox, "task_retries_general")
            timeout = _child(widget, QSpinBox, "task_timeout_general")

            self.assertEqual(provider.currentText(), "Gemini")
            self.assertEqual(model.currentData(), "gemini-2.5-flash")
            self.assertEqual(tokens.value(), 4096)
            self.assertEqual(retries.value(), 2)
            self.assertEqual(timeout.value(), 90)
        finally:
            widget.deleteLater()

    def test_model_defaults_save_updates_config_and_emits_signal(self):
        widget = LLMSettingsWidget()
        emitted = []
        widget.settings_saved.connect(lambda: emitted.append(True))
        try:
            provider = _child(widget, QComboBox, "task_provider_general_0")
            model = _child(widget, QComboBox, "task_model_general_0")
            tokens = _child(widget, QSpinBox, "task_tokens_general_0")
            save_button = _child(widget, QPushButton, "save_llm_settings")

            provider.setCurrentText("OpenAI")
            model_index = model.findData("gpt-4o-mini")
            self.assertGreaterEqual(model_index, 0)
            model.setCurrentIndex(model_index)
            tokens.setValue(2048)

            with patch.object(dialogs_module.QMessageBox, "information"):
                save_button.click()

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
        repo_root = Path(__file__).resolve().parents[1]
        source_path = repo_root / "iCharlotte.py"
        tree = _source_tree(source_path)

        import_from_names = [
            alias.name
            for node in ast.walk(tree)
            if isinstance(node, ast.ImportFrom)
            for alias in node.names
        ]
        method_names = [
            node.name
            for node in ast.walk(tree)
            if isinstance(node, ast.FunctionDef)
        ]
        llm_settings_actions = [
            node
            for node in ast.walk(tree)
            if isinstance(node, ast.Call)
            and isinstance(node.func, ast.Attribute)
            and node.func.attr == "addAction"
            and node.args
            and isinstance(node.args[0], ast.Constant)
            and node.args[0].value == "LLM Settings"
        ]

        self.assertNotIn("LLMSettingsDialog", import_from_names)
        self.assertNotIn("open_settings_dialog", method_names)
        self.assertEqual(llm_settings_actions, [])


if __name__ == "__main__":
    unittest.main()
