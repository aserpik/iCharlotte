"""Tests for the task registry."""
import pytest

from icharlotte_core.ui.wizard.registry import (
    TASK_REGISTRY,
    TaskSpec,
    get_task,
    list_tasks,
)


def test_four_initial_tasks_registered():
    ids = {t.task_id for t in list_tasks()}
    assert ids == {
        "summarize_documents",
        "summarize_discovery",
        "summarize_depositions",
        "medical_records",
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


def test_get_task_unknown_raises():
    with pytest.raises(KeyError):
        get_task("not_a_real_task")


def test_each_task_has_script_name():
    assert get_task("summarize_documents").script_name == "summarize.py"
    assert get_task("summarize_discovery").script_name == "summarize_discovery.py"
    assert get_task("summarize_depositions").script_name == "summarize_deposition.py"
    assert get_task("medical_records").script_name == "med_record.py"
