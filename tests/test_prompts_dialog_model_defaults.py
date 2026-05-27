"""Tests for pass-scoped model defaults in the Prompt Engineering Workbench."""

import ast
import os
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from PySide6.QtWidgets import QApplication, QComboBox, QLabel, QPushButton

app = QApplication.instance() or QApplication(sys.argv)

from icharlotte_core.llm_config import AgentConfig, ModelSpec, TaskConfig
from icharlotte_core.ui import dialogs as dialogs_module
from icharlotte_core.ui.dialogs import PromptsDialog


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
        self.updated_passes = []
        self.updated_tasks = []
        self.pass_configs = {
            ("agent_sum_depo", "summary"): [
                ModelSpec(provider="OpenAI", model="gpt-4o-mini", max_tokens=-1),
                ModelSpec(provider="Gemini", model="gemini-3.1-pro-preview", max_tokens=-1),
            ]
        }
        self.task_configs = {
            "general": TaskConfig(
                name="general",
                model_sequence=[
                    ModelSpec(provider="Gemini", model="gemini-3.5-flash", max_tokens=4096)
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

    def get_model_sequence_for_agent(self, agent_id, fallback_task=None, pass_name=None):
        sequence = self.pass_configs.get((agent_id, pass_name))
        if sequence:
            return sequence
        return self.get_model_sequence(fallback_task or "general")

    def update_pass_model_config(self, agent_id, pass_name, model_sequence):
        self.updated_passes.append((agent_id, pass_name, model_sequence))
        self.pass_configs[(agent_id, pass_name)] = model_sequence

    def update_agent_config(self, agent_id, model_sequence=None, use_default=None, max_retries=None, timeout_seconds=None):
        self.updated_agents.append((agent_id, model_sequence, use_default, max_retries, timeout_seconds))

    def get_model_sequence(self, task_type="general"):
        return self.task_configs.get(task_type, self.task_configs["general"]).model_sequence


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
            self.assertIsNotNone(dlg.model_defaults_widget)
            self.assertFalse(hasattr(dlg, "model_settings_panel"))
        finally:
            dlg.deleteLater()

    def test_model_defaults_tab_tracks_selected_agent_and_pass(self):
        dlg = PromptsDialog()
        try:
            dlg.agent_combo.setCurrentText("deposition")
            dlg.pass_combo.setCurrentText("summary")

            target = _child(dlg, QLabel, "pass_model_defaults_target")
            primary = _child(dlg, QComboBox, "primary_default_model")
            first_backup = _child(dlg, QComboBox, "first_backup_model")
            second_backup = _child(dlg, QComboBox, "second_backup_model")

            self.assertIn("deposition / summary", target.text())
            self.assertEqual(primary.currentData(), "OpenAI|gpt-4o-mini")
            self.assertEqual(first_backup.currentData(), "Gemini|gemini-3.1-pro-preview")
            self.assertIsNone(second_backup.currentData())
        finally:
            dlg.deleteLater()

    def test_model_defaults_save_updates_selected_agent_pass(self):
        dlg = PromptsDialog()
        try:
            dlg.agent_combo.setCurrentText("deposition")
            dlg.pass_combo.setCurrentText("summary")

            primary = _child(dlg, QComboBox, "primary_default_model")
            first_backup = _child(dlg, QComboBox, "first_backup_model")
            second_backup = _child(dlg, QComboBox, "second_backup_model")
            save_button = _child(dlg, QPushButton, "save_pass_model_defaults")

            primary.setCurrentIndex(primary.findData("Gemini|gemini-3.5-flash"))
            first_backup.setCurrentIndex(first_backup.findData("OpenAI|gpt-4o-mini"))
            second_backup.setCurrentIndex(second_backup.findData(None))

            with patch.object(dialogs_module.QMessageBox, "information"):
                save_button.click()

            saved = dlg.llm_config.updated_passes[-1]
            self.assertEqual(saved[0], "agent_sum_depo")
            self.assertEqual(saved[1], "summary")
            self.assertEqual([(m.provider, m.model) for m in saved[2]], [
                ("Gemini", "gemini-3.5-flash"),
                ("OpenAI", "gpt-4o-mini"),
            ])
        finally:
            dlg.deleteLater()

    def test_pass_change_refreshes_model_defaults_tab(self):
        dlg = PromptsDialog()
        try:
            dlg.agent_combo.setCurrentText("deposition")
            dlg.pass_combo.setCurrentText("summary")
            primary = _child(dlg, QComboBox, "primary_default_model")
            self.assertEqual(primary.currentData(), "OpenAI|gpt-4o-mini")

            dlg.pass_combo.setCurrentText("cross_check")
            self.assertEqual(primary.currentData(), "Gemini|gemini-3.5-flash")
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
