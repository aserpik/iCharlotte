"""Tests for the MedChronConfigForm + MedChronSettingsPage UI."""

import json
import sys
from pathlib import Path

import pytest

pytest.importorskip("pytestqt")  # NOTE: no underscore — pytest_qt silently skips

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))


@pytest.fixture(autouse=True)
def _isolated_custom_analyses_store(tmp_path, monkeypatch):
    """Redirect the global custom-analyses JSON to a per-test tmp path so
    tests don't see (or pollute) the developer's real saved analyses."""
    from icharlotte_core.med_chron import custom_analyses_store
    monkeypatch.setattr(
        custom_analyses_store,
        "_STORE_PATH",
        tmp_path / "store" / "med_chron_custom_analyses.json",
    )
    yield


def _write_session(tmp_path: Path, *, narrative_missing: bool = False) -> Path:
    cache = tmp_path / ".med_chron" / "abc123"
    cache.mkdir(parents=True)
    session_path = cache / "session.json"
    session_path.write_text(json.dumps({
        "version": 1,
        "phase": "awaiting_input",
        "input_path": str(tmp_path / "rec.docx"),
        "narrative_text_path": str(cache / "narrative.txt"),
        "full_text_path": str(cache / "full.txt"),
        "narrative_missing": narrative_missing,
        "provider_name": "Acme PT",
        "file_number": "1234.567",
        "catalog": [
            {"id": "rewrite_chronology", "title": "Rewrite Chronology",
             "description": "Reformats narrative.", "uses_tables": False,
             "default_selected": True},
            {"id": "inconsistencies", "title": "Inconsistency Check",
             "description": "Find contradictions.", "uses_tables": True,
             "default_selected": False},
        ],
        "user_config": None,
    }, indent=2), encoding="utf-8")
    return session_path


def test_form_renders_one_checkbox_per_catalog_entry(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    assert len(form.catalog_checkboxes) == 2


def test_default_selected_checkbox_starts_checked(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    rewrite_cb = form.catalog_checkboxes["rewrite_chronology"]
    other_cb = form.catalog_checkboxes["inconsistencies"]
    assert rewrite_cb.isChecked() is True
    assert other_cb.isChecked() is False


def test_narrative_missing_banner_shown(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    session_path = _write_session(tmp_path, narrative_missing=True)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    form.show()  # banner is parent-of-form; force visibility for assertion
    qtbot.waitExposed(form)
    assert form.narrative_missing_banner.isVisible()


def test_narrative_missing_banner_hidden_by_default(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    session_path = _write_session(tmp_path, narrative_missing=False)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    form.show()
    qtbot.waitExposed(form)
    assert form.narrative_missing_banner.isVisible() is False


def test_proceed_requires_at_least_one_selection(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    form.catalog_checkboxes["rewrite_chronology"].setChecked(False)
    assert form.commit_user_config() is False


def test_custom_row_requires_label_and_instruction(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    form.catalog_checkboxes["rewrite_chronology"].setChecked(False)
    form.add_custom_row()
    row = form.custom_rows[0]
    row.label_edit.setText("Some label")
    assert form.commit_user_config() is False


def test_empty_custom_rows_silently_dropped(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    from icharlotte_core.med_chron import session_manager
    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    form.add_custom_row()
    form.add_custom_row()
    form.add_custom_row()
    form.custom_rows[1].label_edit.setText("Real one")
    form.custom_rows[1].instruction_edit.setPlainText("Do this thing.")
    assert form.commit_user_config() is True
    data = session_manager.read_session(session_path)
    cfg = data["user_config"]
    assert len(cfg["custom_analyses"]) == 1
    assert cfg["custom_analyses"][0]["label"] == "Real one"


def test_commit_flips_phase_to_ready_to_run(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    from icharlotte_core.med_chron import session_manager
    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    assert form.commit_user_config() is True
    data = session_manager.read_session(session_path)
    assert data["phase"] == "ready_to_run"
    assert data["user_config"]["selected_catalog_ids"] == ["rewrite_chronology"]


def test_settings_page_proceed_disabled_until_phase1_completes(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.med_chron_settings_page import MedChronSettingsPage
    from icharlotte_core.ui.wizard.registry import TaskSpec

    spec = TaskSpec(
        task_id="med_chron_analysis",
        title="Med Chron Analysis",
        description="…",
        icon_glyph="🩺",
        script_name="med_chron.py",
    )
    page = MedChronSettingsPage(spec, files=[str(tmp_path / "rec.docx")])
    qtbot.addWidget(page)
    assert page.proceed_btn.isEnabled() is False


def test_settings_page_swaps_to_form_on_awaiting_input(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.med_chron_settings_page import MedChronSettingsPage
    from icharlotte_core.ui.wizard.registry import TaskSpec

    spec = TaskSpec(
        task_id="med_chron_analysis",
        title="Med Chron Analysis",
        description="…",
        icon_glyph="🩺",
        script_name="med_chron.py",
    )
    page = MedChronSettingsPage(spec, files=[str(tmp_path / "rec.docx")])
    qtbot.addWidget(page)

    # Build a valid session for the form to read (use _write_session helper from earlier in the file).
    session_path = _write_session(tmp_path)

    # Simulate Phase 1 completion.
    page._on_phase1_complete(str(session_path))

    assert page._stack.currentIndex() == 1   # form page
    assert page.proceed_btn.isEnabled() is True
    assert page._form is not None


# ---- File-swap behavior (bug repro: settings page edits ignored by Phase 2) ----


def _settings_page_with_phase1_done(qtbot, tmp_path, original_filename: str = "orig.docx"):
    """Build a MedChronSettingsPage that has already completed Phase 1 for one file."""
    from icharlotte_core.ui.wizard.pages.med_chron_settings_page import MedChronSettingsPage
    from icharlotte_core.ui.wizard.registry import TaskSpec

    spec = TaskSpec(
        task_id="med_chron_analysis",
        title="Med Chron Analysis",
        description="…",
        icon_glyph="🩺",
        script_name="med_chron.py",
    )
    original_path = str(tmp_path / original_filename)
    page = MedChronSettingsPage(spec, files=[original_path], case_root=str(tmp_path))
    qtbot.addWidget(page)

    # Simulate the speculative-run path: TaskTab would build the worker and
    # call attach_worker() before launching the process. attach_worker only
    # needs .files (to record _prep_file) plus the four signals it connects.
    from PySide6.QtCore import QObject, Signal

    class _StubWorker(QObject):
        status = Signal(str)
        progress = Signal(int)
        awaiting_input = Signal(str)
        failed = Signal(str)

        def __init__(self, files):
            super().__init__()
            self.files = list(files)

    page.attach_worker(_StubWorker([original_path]))

    # Simulate Phase 1 completion (form appears, _prep_file == original_path).
    session_path = _write_session(tmp_path)
    page._on_phase1_complete(str(session_path))
    return page, original_path, session_path


def test_swapping_file_emits_restart_phase1_with_new_file(qtbot, tmp_path):
    """The bug: changing the file on the settings page must trigger a new
    Phase 1 against the new file (not silently reuse the original session).
    """
    page, original_path, _ = _settings_page_with_phase1_done(qtbot, tmp_path)
    new_path = str(tmp_path / "swapped.docx")

    received = []
    page.restart_phase1_requested.connect(lambda files: received.append(list(files)))

    # Simulate Remove (the only file) + Add (a different file) by mutating
    # the file list the way base SettingsPage does, then routing through the
    # public hook.
    page._files = []
    page._refresh_files_list()
    page._maybe_restart_phase1()           # remove half — emits []

    page._files = [new_path]
    page._refresh_files_list()
    page._maybe_restart_phase1()           # add half — emits [new_path]

    # Both events fired in order: cancel-only then restart-with-new-file.
    assert received == [[], [new_path]]


def test_swapping_file_resets_to_prep_state(qtbot, tmp_path):
    """After a file swap the form must be torn down and the spinner page shown,
    so the user can't click Proceed against the stale session_path.
    """
    page, _, _ = _settings_page_with_phase1_done(qtbot, tmp_path)
    assert page._form is not None                  # precondition: form is up
    assert page._stack.currentIndex() == 1         # form page
    assert page.proceed_btn.isEnabled() is True
    stale_session = page._session_path

    new_path = str(tmp_path / "swapped.docx")
    page._files = [new_path]
    page._refresh_files_list()
    page._maybe_restart_phase1()

    assert page._form is None
    assert page._session_path is None, "old session_path must be cleared"
    assert page._session_path != stale_session
    assert page._stack.currentIndex() == 0         # spinner page
    assert page.proceed_btn.isEnabled() is False


def test_picking_same_file_again_does_not_restart(qtbot, tmp_path):
    """No-op when the user re-picks the file that's already selected."""
    page, original_path, _ = _settings_page_with_phase1_done(qtbot, tmp_path)
    received = []
    page.restart_phase1_requested.connect(lambda files: received.append(list(files)))

    # Same file currently selected → _maybe_restart_phase1 should be a no-op.
    page._maybe_restart_phase1()
    assert received == []


def test_tasktab_restart_handler_swaps_files_and_starts_new_worker(qtbot, tmp_path, monkeypatch):
    """Wiring check: TaskTab._restart_speculative_run cancels the old worker,
    updates self._files, and launches a fresh worker that the settings page
    re-attaches to. Without this wiring the settings page would emit the
    restart signal into the void and Phase 2 would still use the stale cache.
    """
    from PySide6.QtCore import QObject, Signal
    import icharlotte_core.ui.wizard.runners.subprocess_worker as sw_mod
    from icharlotte_core.ui.wizard.registry import get_task
    from icharlotte_core.ui.wizard.task_tab import TaskTab

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
            self.files = list(kwargs.get("files") or [])
            self.started = False
            self.cancel_called = False
            instances.append(self)

        def start(self):
            self.started = True

        def cancel(self):
            self.cancel_called = True

    monkeypatch.setattr(sw_mod, "SubprocessWorker", _FakeWorker)

    spec = get_task("med_chron_analysis")
    tab = TaskTab(spec, files=[str(tmp_path / "orig.docx")], case_path=str(tmp_path), file_number="0000.000")
    qtbot.addWidget(tab)
    tab.start_speculative_run()

    assert len(instances) == 1
    assert instances[0].files == [str(tmp_path / "orig.docx")]

    # Settings page asks TaskTab to swap files.
    new_path = str(tmp_path / "swapped.docx")
    tab.settings_page.restart_phase1_requested.emit([new_path])

    # Old worker was cancelled; new worker was created with new files.
    assert instances[0].cancel_called is True
    assert len(instances) == 2
    assert instances[1].files == [new_path]
    assert instances[1].started is True
    assert tab.files == [new_path]


# ---- Global custom-analyses persistence ----


def test_saved_custom_analyses_prepopulate_form(qtbot, tmp_path):
    from icharlotte_core.med_chron import custom_analyses_store
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm

    # Seed the (isolated, fixture-redirected) store before opening the form.
    custom_analyses_store.save([
        {"label": "Left-knee mentions", "instruction": "Find every entry about the left knee."},
        {"label": "Insurance gaps", "instruction": "Identify periods without coverage."},
    ])

    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)

    assert len(form.custom_rows) == 2
    assert form.custom_rows[0].label() == "Left-knee mentions"
    assert form.custom_rows[0].instruction() == "Find every entry about the left knee."
    assert form.custom_rows[1].label() == "Insurance gaps"
    # Pre-populated saved rows are NOT included by default — the user must
    # explicitly opt them into this run (mirrors the curated list, which also
    # defaults to off). Prevents saved customs silently running unselected.
    assert form.custom_rows[0].is_included() is False
    assert form.custom_rows[1].is_included() is False


def test_prepopulated_saved_rows_do_not_run_unless_checked(qtbot, tmp_path):
    from icharlotte_core.med_chron import custom_analyses_store, session_manager
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm

    custom_analyses_store.save([
        {"label": "Medications", "instruction": "Track medications."},
        {"label": "Inconsistent Reporting", "instruction": "Find inconsistencies."},
    ])

    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)

    # User only selects a curated analysis; leaves the saved customs untouched.
    form.catalog_checkboxes["rewrite_chronology"].setChecked(True)

    assert form.commit_user_config() is True

    session = session_manager.read_session(session_path)
    user_config = session["user_config"]
    assert user_config["selected_catalog_ids"] == ["rewrite_chronology"]
    # No saved custom should run since none were checked.
    assert user_config["custom_analyses"] == []


def test_newly_typed_custom_row_defaults_to_included(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm

    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)

    # A row the user adds fresh in this session is opted in by default —
    # they just added it because they want it to run.
    row = form.add_custom_row()
    assert row.is_included() is True


def test_typed_custom_analyses_save_to_global_store_on_commit(qtbot, tmp_path):
    from icharlotte_core.med_chron import custom_analyses_store
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm

    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)

    # Type a new custom analysis.
    form.add_custom_row()
    form.custom_rows[0].label_edit.setText("Surgeries")
    form.custom_rows[0].instruction_edit.setPlainText("List every surgery date.")

    assert form.commit_user_config() is True

    # Stored to global file.
    saved = custom_analyses_store.load()
    assert saved == [{"label": "Surgeries", "instruction": "List every surgery date."}]


def test_unchecked_saved_row_persists_but_does_not_run(qtbot, tmp_path):
    from icharlotte_core.med_chron import custom_analyses_store, session_manager
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm

    custom_analyses_store.save([
        {"label": "Skip me", "instruction": "Skipped this run."},
        {"label": "Run me", "instruction": "Should run."},
    ])

    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)

    # Saved rows load unchecked; opt only the second one into this run.
    assert form.custom_rows[0].is_included() is False
    assert form.custom_rows[1].is_included() is False
    form.custom_rows[1].include_cb.setChecked(True)

    assert form.commit_user_config() is True

    # Both still saved globally.
    saved = custom_analyses_store.load()
    labels = {a["label"] for a in saved}
    assert labels == {"Skip me", "Run me"}

    # But only the checked one ran (i.e. landed in user_config).
    data = session_manager.read_session(session_path)
    cfg = data["user_config"]
    assert [a["label"] for a in cfg["custom_analyses"]] == ["Run me"]


def test_removed_saved_row_is_dropped_from_global_store(qtbot, tmp_path):
    from icharlotte_core.med_chron import custom_analyses_store
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm

    custom_analyses_store.save([
        {"label": "Keep", "instruction": "K."},
        {"label": "Delete", "instruction": "D."},
    ])

    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)

    # Remove the second pre-populated row (Delete).
    form._remove_custom_row(form.custom_rows[1])

    assert form.commit_user_config() is True

    saved = custom_analyses_store.load()
    assert saved == [{"label": "Keep", "instruction": "K."}]


def test_form_has_vertical_splitter_between_curated_and_custom(qtbot, tmp_path):
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QSplitter
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm

    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    assert isinstance(form._analyses_splitter, QSplitter)
    assert form._analyses_splitter.orientation() == Qt.Orientation.Vertical
    assert form._analyses_splitter.count() == 2
    assert form._analyses_splitter.childrenCollapsible() is False


def test_settings_page_has_vertical_splitter_for_files(qtbot, tmp_path):
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QSplitter
    from icharlotte_core.ui.wizard.pages.med_chron_settings_page import MedChronSettingsPage
    from icharlotte_core.ui.wizard.registry import TaskSpec

    spec = TaskSpec(
        task_id="med_chron_analysis",
        title="Med Chron Analysis",
        description="…",
        icon_glyph="🩺",
        script_name="med_chron.py",
    )
    page = MedChronSettingsPage(spec, files=[str(tmp_path / "rec.docx")])
    qtbot.addWidget(page)
    assert isinstance(page._outer_splitter, QSplitter)
    assert page._outer_splitter.orientation() == Qt.Orientation.Vertical
    assert page._outer_splitter.count() == 2
    assert page._outer_splitter.childrenCollapsible() is False


def test_edited_saved_row_writes_back_edited_text(qtbot, tmp_path):
    from icharlotte_core.med_chron import custom_analyses_store
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm

    custom_analyses_store.save([
        {"label": "Old label", "instruction": "Old instruction."},
    ])

    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)

    form.custom_rows[0].label_edit.setText("New label")
    form.custom_rows[0].instruction_edit.setPlainText("New instruction.")

    assert form.commit_user_config() is True

    saved = custom_analyses_store.load()
    assert saved == [{"label": "New label", "instruction": "New instruction."}]
