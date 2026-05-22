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


def test_spec_property_returns_spec(qtbot):
    spec = get_task("summarize_documents")
    tab = TaskTab(spec, files=["/tmp/x.pdf"], case_path="/tmp/case", file_number="0000.000")
    qtbot.addWidget(tab)
    assert tab.spec is spec


# ---- Worker selection ----


def test_pick_worker_cls_multi_file_summarize_uses_parallel(qtbot):
    """Most multi-file agents should use the per-file parallel runner."""
    from icharlotte_core.ui.wizard.task_tab import TaskTab
    from icharlotte_core.ui.wizard.runners.parallel_subprocess_worker import (
        ParallelSubprocessWorker,
    )
    assert TaskTab._pick_worker_cls("summarize.py", num_files=3) is ParallelSubprocessWorker


def test_pick_worker_cls_multi_file_discovery_uses_dispatcher(qtbot):
    """summarize_discovery.py with multiple files must use the dispatcher
    runner so the script's own party-grouping + consolidation kicks in.
    A per-file parallel runner with --output_path defeats both.
    """
    from icharlotte_core.ui.wizard.task_tab import TaskTab
    from icharlotte_core.ui.wizard.runners.dispatcher_subprocess_worker import (
        DispatcherSubprocessWorker,
    )
    assert (
        TaskTab._pick_worker_cls("summarize_discovery.py", num_files=4)
        is DispatcherSubprocessWorker
    )


def test_pick_worker_cls_single_file_uses_sequential(qtbot):
    """One-file runs always go through the sequential worker, even for
    discovery — process_document already does party-aware appending
    against existing files in NOTES/AI Output/, so we don't need the
    dispatcher.
    """
    from icharlotte_core.ui.wizard.task_tab import TaskTab
    from icharlotte_core.ui.wizard.runners.subprocess_worker import SubprocessWorker
    assert (
        TaskTab._pick_worker_cls("summarize_discovery.py", num_files=1)
        is SubprocessWorker
    )
    assert TaskTab._pick_worker_cls("summarize.py", num_files=1) is SubprocessWorker


def test_pick_worker_cls_two_phase_always_sequential(qtbot):
    """Two-phase agents (AWAITING_INPUT/Phase 2) must use the sequential
    runner regardless of file count."""
    from icharlotte_core.ui.wizard.task_tab import TaskTab
    from icharlotte_core.ui.wizard.runners.subprocess_worker import SubprocessWorker
    assert (
        TaskTab._pick_worker_cls("summarize_deposition.py", num_files=5)
        is SubprocessWorker
    )
    assert (
        TaskTab._pick_worker_cls("med_chron.py", num_files=3)
        is SubprocessWorker
    )
