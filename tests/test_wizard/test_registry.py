"""Tests for the task registry."""
import pytest

from icharlotte_core.ui.wizard.registry import (
    TASK_REGISTRY,
    TaskSpec,
    get_task,
    list_tasks,
)


def test_initial_tasks_registered():
    ids = {t.task_id for t in list_tasks()}
    assert ids == {
        "summarize_documents",
        "summarize_discovery",
        "summarize_depositions",
        "depo_prep",
        "medical_records",
        "med_chron_analysis",
        "med_record_extractor",
        "separate",
        "subpoena_tracker",
        "respond_to_discovery",
        "oppose_motion",
        "generate_motion",
        "mediation_brief",
        "case_intake_docket",
        "chat",
    }


def test_each_task_has_required_metadata():
    for spec in list_tasks():
        assert isinstance(spec, TaskSpec)
        assert spec.task_id
        assert spec.title
        assert spec.description
        assert isinstance(spec.default_folders, list)


def test_default_folders_per_task():
    assert get_task("summarize_documents").default_folders == []
    assert get_task("summarize_discovery").default_folders == ["DISCOVERY/RESPONSES", "DISCOVERY"]
    assert get_task("summarize_depositions").default_folders == ["DISCOVERY/TRANSCRIPTS", "DISCOVERY"]
    assert get_task("medical_records").default_folders == ["RECORDS"]


def test_case_intake_docket_registry_spec():
    spec = get_task("case_intake_docket")
    assert spec.title == "Case Intake & Docket"
    assert (
        spec.description
        == "Extract complaint metadata, review case details, then download and process the court docket."
    )
    assert spec.icon_glyph == "\U0001F5C2"
    assert spec.script_name == ""
    assert spec.category == "General"
    assert spec.default_folders == []
    assert {
        "complaint",
        "docket",
        "case number",
        "venue",
        "intake",
        "hearing",
        "trial",
    } <= set(spec.keywords)


def test_get_task_unknown_raises():
    with pytest.raises(KeyError):
        get_task("not_a_real_task")


def test_each_task_has_script_name():
    assert get_task("summarize_documents").script_name == "summarize.py"
    assert get_task("summarize_discovery").script_name == "summarize_discovery.py"
    assert get_task("summarize_depositions").script_name == "summarize_deposition.py"
    assert get_task("medical_records").script_name == "med_record.py"


def test_each_task_has_settings_page_cls():
    from icharlotte_core.ui.wizard.pages.settings_page import SettingsPage
    for spec in list_tasks():
        assert spec.settings_page_cls is not None
        assert issubclass(spec.settings_page_cls, SettingsPage)
