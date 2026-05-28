"""TaskTab-level wiring tests for the Depo Prep two-phase flow.

These guard the integration seam that unit tests miss: the per-page tests call
DepoPrepSettingsPage._on_phase1_complete() directly and never exercise how
TaskTab routes the settings page's signals. The Critical defect this file
guards against was: Analyze triggered proceed_requested (which switches to the
Status page and routes awaiting_input to the generic deposition dialog), so the
embedded TopicEditor never appeared and Phase 2 was unreachable.
"""
import pytest

pytest.importorskip("pytestqt")
pytest.importorskip("PySide6")


@pytest.fixture
def spec():
    from icharlotte_core.ui.wizard.registry import get_task
    return get_task("depo_prep")


def _make_tab(spec, tmp_path):
    from icharlotte_core.ui.wizard.task_tab import TaskTab
    return TaskTab(spec, files=[], case_path=str(tmp_path), file_number="1234.567")


def test_tasktab_uses_depo_prep_settings_and_output_pages(qtbot, spec, tmp_path):
    from icharlotte_core.ui.wizard.pages.depo_prep_settings_page import DepoPrepSettingsPage
    from icharlotte_core.ui.wizard.pages.depo_prep_output_page import DepoPrepOutputPage
    tab = _make_tab(spec, tmp_path)
    qtbot.addWidget(tab)
    assert isinstance(tab.settings_page, DepoPrepSettingsPage)
    assert isinstance(tab.output_page, DepoPrepOutputPage)


def test_analyze_requested_routes_to_start_analyze_run_without_page_switch(qtbot, spec, tmp_path):
    """The Critical-regression guard: emitting analyze_requested must invoke the
    settings-owned worker run (NOT proceed_requested's page switch), refresh the
    tab's files to the config path, and leave the tab on the Settings page so the
    inline topic editor can appear."""
    from icharlotte_core.ui.wizard.task_tab import PAGE_SETTINGS
    tab = _make_tab(spec, tmp_path)
    qtbot.addWidget(tab)

    # Stub start_speculative_run so no subprocess is spawned; record that the
    # settings-owned run path was taken.
    calls = {"speculative": 0}
    tab.start_speculative_run = lambda: calls.__setitem__("speculative", calls["speculative"] + 1)

    page = tab.settings_page
    page.set_deponent_name("Jane Doe")
    page.set_deponent_role("Plaintiff")
    (tmp_path / "fake_depo.pdf").write_bytes(b"")
    page.add_deponent_files([str(tmp_path / "fake_depo.pdf")])

    page._on_analyze_clicked()

    # _start_analyze_run refreshed the tab's files from the page (the config path)
    # and delegated to start_speculative_run.
    assert calls["speculative"] == 1
    assert len(tab.files) == 1
    assert tab.files[0].endswith("config.json")
    # Crucially, we did NOT switch to the Status page.
    assert tab.current_page == PAGE_SETTINGS


def test_phase2_requested_connected_to_advance(qtbot, spec, tmp_path):
    """phase2_requested(session_path) must reach advance_to_status_with_phase2."""
    tab = _make_tab(spec, tmp_path)
    qtbot.addWidget(tab)

    captured = {}
    tab.advance_to_status_with_phase2 = lambda p: captured.__setitem__("session", p)
    # Reconnect because we just replaced the bound method after __init__ wired it.
    tab.settings_page.phase2_requested.connect(tab.advance_to_status_with_phase2)

    tab.settings_page.phase2_requested.emit("C:/case/session.json")
    assert captured.get("session") == "C:/case/session.json"
