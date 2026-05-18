"""Tests for the med_chron_analysis wizard task registration."""

import sys
from pathlib import Path

import pytest

pytest.importorskip("pytestqt")

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))


def test_med_chron_analysis_task_is_registered():
    from icharlotte_core.ui.wizard.registry import TASK_REGISTRY
    assert "med_chron_analysis" in TASK_REGISTRY


def test_med_chron_task_uses_med_chron_settings_page():
    from icharlotte_core.ui.wizard.registry import get_task
    from icharlotte_core.ui.wizard.pages.med_chron_settings_page import MedChronSettingsPage
    spec = get_task("med_chron_analysis")
    assert spec.settings_page_cls is MedChronSettingsPage


def test_med_chron_task_phase_flags_set():
    from icharlotte_core.ui.wizard.registry import get_task
    spec = get_task("med_chron_analysis")
    assert "--phase=prep" in list(spec.phase1_args)
    assert spec.phase2_flag == "--phase=run"


def test_med_chron_task_script_name():
    from icharlotte_core.ui.wizard.registry import get_task
    spec = get_task("med_chron_analysis")
    assert spec.script_name == "med_chron.py"


def test_existing_deposition_task_phase_flags_unchanged():
    """Defaults must keep existing deposition behavior intact."""
    from icharlotte_core.ui.wizard.registry import get_task
    spec = get_task("summarize_depositions")
    assert spec.phase2_flag == "--phase=summary"
    assert list(spec.phase1_args) == []  # no extra flag for phase 1
