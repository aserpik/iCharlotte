"""Smoke test for TaskTab state machine."""
import pytest

pytest.importorskip("pytestqt")
from icharlotte_core.ui.wizard.registry import get_task
from icharlotte_core.ui.wizard.task_tab import TaskTab, PAGE_SETTINGS, PAGE_STATUS, PAGE_OUTPUT


def test_initial_state_is_settings(qtbot):
    tab = TaskTab(get_task("summarize_documents"), files=["/tmp/x.pdf"], case_path="/tmp/case", file_number="0000.000")
    qtbot.addWidget(tab)
    assert tab.current_page == PAGE_SETTINGS


def test_show_output_transitions(qtbot):
    tab = TaskTab(get_task("summarize_documents"), files=["/tmp/x.pdf"], case_path="/tmp/case", file_number="0000.000")
    qtbot.addWidget(tab)
    tab._show_output("/tmp/fake_output.docx")
    assert tab.current_page == PAGE_OUTPUT
    assert tab.output_page.output_path == "/tmp/fake_output.docx"
