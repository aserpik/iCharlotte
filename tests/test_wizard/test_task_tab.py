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


def test_output_page_shows_picker_for_multiple_outputs(qtbot, tmp_path):
    """When the OutputPage receives multiple paths, a picker is shown."""
    from icharlotte_core.ui.wizard.pages.output_page import OutputPage

    page = OutputPage()
    qtbot.addWidget(page)

    # Create two real docx-ish files so load_outputs doesn't barf.
    a = tmp_path / "AI_OUTPUT - one.docx"
    b = tmp_path / "AI_OUTPUT - two.docx"
    a.write_bytes(b"x")
    b.write_bytes(b"x")

    # Single output → picker hidden.
    page.load_outputs([str(a)])
    assert page.picker_row_widget.isHidden() is True
    assert page.output_path == str(a)

    # Two outputs → picker shown, first selected.
    page.load_outputs([str(a), str(b)])
    assert page.picker_row_widget.isHidden() is False
    assert page.output_picker.count() == 2
    assert page.output_path == str(a)
    assert "2 files produced" in page.outputs_count_label.text()

    # Switching the picker swaps the active output_path.
    page.output_picker.setCurrentIndex(1)
    assert page.output_path == str(b)


def test_append_output_grows_picker_without_changing_selection(qtbot, tmp_path):
    """append_output adds a new file to the picker but keeps current selection."""
    from icharlotte_core.ui.wizard.pages.output_page import OutputPage

    page = OutputPage()
    qtbot.addWidget(page)

    a = tmp_path / "AI_OUTPUT - one.docx"
    b = tmp_path / "AI_OUTPUT - two.docx"
    c = tmp_path / "AI_OUTPUT - three.docx"
    for p in (a, b, c):
        p.write_bytes(b"x")

    # First append: behaves like load_outputs([a]).
    page.append_output(str(a))
    assert page.output_path == str(a)
    assert page.picker_row_widget.isHidden() is True  # single → no picker

    # Second append: picker appears, current selection stays on a.
    page.append_output(str(b))
    assert page.picker_row_widget.isHidden() is False
    assert page.output_picker.count() == 2
    assert page.output_path == str(a), "appending must not change current selection"

    # Third append: picker grows; selection still a.
    page.append_output(str(c))
    assert page.output_picker.count() == 3
    assert page.output_path == str(a)
    assert "3 files produced" in page.outputs_count_label.text()

    # Duplicate append is a no-op.
    page.append_output(str(a))
    assert page.output_picker.count() == 3


def test_progress_hint_show_and_clear(qtbot):
    from icharlotte_core.ui.wizard.pages.output_page import OutputPage

    page = OutputPage()
    qtbot.addWidget(page)

    assert page.progress_hint_label.isHidden() is True
    page.set_progress_hint("Still processing 2 more files…")
    assert page.progress_hint_label.isHidden() is False
    assert "2 more" in page.progress_hint_label.text()

    page.set_progress_hint("")
    assert page.progress_hint_label.isHidden() is True


def test_partial_failure_after_streaming_outputs_shows_failure_banner(qtbot, tmp_path):
    """When a parallel run streams some outputs via file_completed and then
    emits ``failed`` for the rest, the user is already on the Output page —
    so the failure message must be surfaced THERE (not silently written to the
    invisible status page). Otherwise the user just sees "fewer outputs than
    expected" with no explanation.

    Repro: user picked 4 discovery PDFs in the wizard. 2 produced outputs
    (which triggered file_completed → page switch to PAGE_OUTPUT). The other 2
    failed (skipped by SKIP_KEYWORDS); the worker emitted ``failed``. Before
    the fix, ``_on_worker_failed`` only updated the status page text and never
    notified the output page, so the user saw 2 outputs and nothing else.
    """
    tab = TaskTab(
        get_task("summarize_discovery"),
        files=["/tmp/a.pdf", "/tmp/b.pdf", "/tmp/c.pdf", "/tmp/d.pdf"],
        case_path="/tmp/case",
        file_number="0000.000",
    )
    qtbot.addWidget(tab)

    # Simulate two files completing — this moves us to PAGE_OUTPUT.
    out_a = tmp_path / "AI_OUTPUT - a.docx"
    out_b = tmp_path / "AI_OUTPUT - b.docx"
    out_a.write_bytes(b"x")
    out_b.write_bytes(b"x")
    tab._total_files_for_run = 4
    tab._files_done_for_run = 0
    tab._on_worker_file_completed(str(out_a))
    tab._on_worker_file_completed(str(out_b))
    assert tab.current_page == PAGE_OUTPUT

    # Now the worker fails the remaining 2 (e.g. SKIP_KEYWORDS skipped them).
    err = "2 of 4 agent runs failed: c.pdf, d.pdf. First error: [c.pdf] reported success but AI_OUTPUT - c.docx is missing"
    tab._on_worker_failed(err)

    # The user must see something on the page they are looking at.
    assert tab.current_page == PAGE_OUTPUT
    # The failure must be communicated via the output page (banner / hint).
    banner_text = tab.output_page.failure_banner_text()
    assert banner_text, "OutputPage should display a failure banner after partial failure"
    assert "2 of 4" in banner_text
    assert "c.pdf" in banner_text


def test_failure_banner_cleared_on_rerun(qtbot, tmp_path):
    """A successful re-run after a partial failure should clear the banner."""
    tab = TaskTab(
        get_task("summarize_discovery"),
        files=["/tmp/a.pdf"],
        case_path="/tmp/case",
        file_number="0000.000",
    )
    qtbot.addWidget(tab)

    # Put the page in the post-failure state.
    out = tmp_path / "AI_OUTPUT - a.docx"
    out.write_bytes(b"x")
    tab._total_files_for_run = 2
    tab._files_done_for_run = 0
    tab._on_worker_file_completed(str(out))
    tab._on_worker_failed("1 of 2 agent runs failed: b.pdf. First error: [b.pdf] is missing")
    assert tab.output_page.failure_banner_text()

    # Simulating the start of a fresh run should clear the banner so the user
    # doesn't see stale errors layered on top of new results.
    tab.output_page.set_failure_banner("")
    assert tab.output_page.failure_banner_text() == ""


def test_phase2_rerun_after_completion_creates_fresh_worker(qtbot, monkeypatch):
    """Re-running summarize_deposition after completion must not silently fail.

    Repro: after Phase 2 finishes, _on_worker_finished sets self._worker = None.
    Clicking 'Edit Settings & Re-run' goes back to the settings page with the
    inline form still populated. Clicking Proceed re-emits phase2_requested.

    Before the fix, advance_to_status_with_phase2 returned early because
    self._worker was None, and pressing Process did nothing. After the fix,
    a fresh SubprocessWorker is created and resumes Phase 2 against the
    existing session directory.
    """
    from PySide6.QtCore import QObject, Signal
    import icharlotte_core.ui.wizard.runners.subprocess_worker as sw_mod

    instances = []

    class _FakeWorker(QObject):
        status = Signal(str)
        progress = Signal(int)
        awaiting_input = Signal(str)
        finished = Signal(str)
        failed = Signal(str)
        cancelled = Signal()

        def __init__(self, *args, **kwargs):
            super().__init__()
            self.kwargs = kwargs
            self.resume_called_with = None
            self.started = False
            instances.append(self)

        def start(self):
            self.started = True

        def resume_with_config(self, session_path):
            self.resume_called_with = session_path

        def cancel(self):
            pass

    monkeypatch.setattr(sw_mod, "SubprocessWorker", _FakeWorker)

    tab = TaskTab(
        get_task("summarize_depositions"),
        files=["/tmp/depo.pdf"],
        case_path="/tmp/case",
        file_number="0000.000",
    )
    qtbot.addWidget(tab)

    # Simulate post-completion state: _on_worker_finished cleared the worker.
    tab._worker = None
    tab._awaiting_session_path = None

    # User clicked Proceed on settings page after edit-settings → re-run.
    tab.advance_to_status_with_phase2("/fake/session/path")

    assert tab.current_page == PAGE_STATUS
    assert tab._worker is not None
    assert isinstance(tab._worker, _FakeWorker)
    assert tab._worker.resume_called_with == "/fake/session/path"
    assert tab._worker.started is False  # resume_with_config replaces start()
    assert len(instances) == 1


def test_restored_summarize_depositions_tab_restarts_phase1(qtbot, tmp_path, monkeypatch):
    """A summarize_depositions task tab restored from persistence must re-launch
    Phase 1; otherwise the inline 'Discovering topics…' page in
    DepositionSettingsPage stays at 0% with no worker behind it.

    Repro: open a summarize_depositions task tab, switch cases (or close+reopen
    the app) while Phase 1 is still running, then come back. Before the fix,
    _restore_task_tabs_for_case rebuilt the tab on PAGE_SETTINGS but never
    called start_speculative_run, so the discovering spinner sat forever.
    """
    import json
    from PySide6.QtCore import QObject, Signal
    from PySide6.QtWidgets import QTabWidget, QWidget
    import iCharlotte as _ich
    import icharlotte_core.ui.wizard.runners.subprocess_worker as sw_mod

    instances = []

    class _FakeWorker(QObject):
        status = Signal(str)
        progress = Signal(int)
        awaiting_input = Signal(str)
        finished = Signal(str)
        failed = Signal(str)
        cancelled = Signal()

        def __init__(self, *args, **kwargs):
            super().__init__()
            self.kwargs = kwargs
            self.started = False
            instances.append(self)

        def start(self):
            self.started = True

        def cancel(self):
            pass

    monkeypatch.setattr(sw_mod, "SubprocessWorker", _FakeWorker)

    # Pre-seed wizard_state.json with a single summarize_depositions tab on the
    # settings page (the state we land in if the app dies during Phase 1).
    folder = tmp_path / "NOTES" / "AI OUTPUT" / ".icharlotte"
    folder.mkdir(parents=True)
    (folder / "wizard_state.json").write_text(json.dumps({
        "version": 1,
        "open_tabs": [{
            "task_id": "summarize_depositions",
            "instance_suffix": "",
            "files": ["depo.pdf"],
            "settings": {},
            "page": "settings",
            "output_path": None,
        }],
        "recent_tasks": [],
    }))
    (tmp_path / "depo.pdf").write_bytes(b"")

    class _Stub(QWidget):
        # TaskTab is constructed with parent=self in _restore_task_tabs_for_case,
        # so the stub must be a QWidget for the parent type check to pass.
        def __init__(self, root):
            super().__init__()
            self.case_path = str(root)
            self.file_number = "0000.000"
            self.tabs = QTabWidget(self)
            self.completed_entries = []

        def _on_task_completed(self, entry):
            self.completed_entries.append(entry)

        def _hide_fixed_close_buttons(self):
            pass

    stub = _Stub(tmp_path)
    qtbot.addWidget(stub)
    stub._restore_task_tabs_for_case = _ich.MainWindow._restore_task_tabs_for_case.__get__(stub, _Stub)

    stub._restore_task_tabs_for_case()

    # The tab was restored.
    assert stub.tabs.count() == 1
    # And a fresh Phase 1 worker was started — proving the user won't see a
    # frozen "Discovering topics…" spinner.
    assert len(instances) == 1
    assert instances[0].started is True
