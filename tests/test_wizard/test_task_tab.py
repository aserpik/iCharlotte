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
