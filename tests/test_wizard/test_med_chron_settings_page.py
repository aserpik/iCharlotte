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
    # Pre-populated rows are included by default.
    assert form.custom_rows[0].is_included() is True
    assert form.custom_rows[1].is_included() is True


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

    # Uncheck the first saved row.
    form.custom_rows[0].include_cb.setChecked(False)

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
