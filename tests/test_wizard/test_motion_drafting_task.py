"""Unified Motion Drafting wizard task tests."""

from pathlib import Path

import pytest

from icharlotte_core.ui.wizard.registry import TASK_REGISTRY, get_task, list_tasks
from icharlotte_core.ui.wizard.task_routing import (
    get_in_process_task_builder_name,
    opens_settings_without_picker,
    requires_initial_file_picker,
)


def test_motion_drafting_replaces_generate_and_oppose_cards():
    ids = {task.task_id for task in list_tasks()}
    assert "motion_drafting" in ids
    assert "generate_motion" not in ids
    assert "oppose_motion" not in ids

    spec = get_task("motion_drafting")
    assert spec.title == "Motion Drafting"
    assert spec.category == "Motions"
    assert spec.script_name == ""
    assert "opposition" in [keyword.lower() for keyword in spec.keywords]
    assert "reply" in [keyword.lower() for keyword in spec.keywords]
    assert TASK_REGISTRY["mediation_brief"].category == "Motions"


def test_motion_drafting_uses_settings_first_in_process_route():
    assert get_in_process_task_builder_name("motion_drafting") == "build_motion_drafting_tab"
    assert opens_settings_without_picker("motion_drafting")
    assert not requires_initial_file_picker("motion_drafting")


def test_legacy_motion_routes_remain_for_saved_tabs():
    assert get_in_process_task_builder_name("generate_motion") == "build_generate_motion_tab"
    assert get_in_process_task_builder_name("oppose_motion") == "build_oppose_motion_tab"


def test_motion_database_options_are_derived_from_taxonomy(tmp_path, monkeypatch):
    from icharlotte_core.motion_drafting.taxonomy import (
        DRAFT_KIND_MOTION,
        DRAFT_KIND_OPPOSITION,
        DRAFT_KIND_REPLY,
        list_motion_type_options,
    )

    root = tmp_path / "MOTION DATABASE"
    dataset = root / "Sample_Pleadings_PDFs"
    (dataset / "Motion - Compel" / "RFP - Document Production").mkdir(parents=True)
    (dataset / "Oppositions" / "Motion to Compel" / "RFP - Document Production").mkdir(parents=True)
    (dataset / "Replies" / "Motion to Compel" / "Separate Statements").mkdir(parents=True)
    (dataset / "_Support - Declarations").mkdir(parents=True)

    monkeypatch.setenv("ICHARLOTTE_MOTION_DATABASE_ROOT", str(root))

    motion_labels = [option.label for option in list_motion_type_options(DRAFT_KIND_MOTION)]
    opposition_labels = [option.label for option in list_motion_type_options(DRAFT_KIND_OPPOSITION)]
    reply_labels = [option.label for option in list_motion_type_options(DRAFT_KIND_REPLY)]

    assert "Motion - Compel" in motion_labels
    assert "Motion - Compel / RFP - Document Production" in motion_labels
    assert "Motion to Compel" in opposition_labels
    assert "Motion to Compel / RFP - Document Production" in opposition_labels
    assert "Motion to Compel / Separate Statements" in reply_labels
    assert all("_Support" not in label for label in motion_labels)


pytest.importorskip("pytestqt")


def test_motion_drafting_settings_updates_motion_types_by_draft_kind(qtbot, tmp_path, monkeypatch):
    from icharlotte_core.motion_drafting.taxonomy import DRAFT_KIND_OPPOSITION, DRAFT_KIND_REPLY
    from icharlotte_core.ui.wizard.pages.motion_drafting_page import MotionDraftingSettingsPage

    root = Path(tmp_path) / "MOTION DATABASE"
    dataset = root / "Sample_Pleadings_PDFs"
    (dataset / "Motion - Demurrer").mkdir(parents=True)
    (dataset / "Oppositions" / "Demurrer").mkdir(parents=True)
    (dataset / "Replies" / "Motion to Compel").mkdir(parents=True)
    monkeypatch.setenv("ICHARLOTTE_MOTION_DATABASE_ROOT", str(root))

    page = MotionDraftingSettingsPage("/case", "123")
    qtbot.addWidget(page)

    page.draft_kind_combo.setCurrentIndex(page.draft_kind_combo.findData(DRAFT_KIND_OPPOSITION))
    assert page.motion_type_combo.findText("Demurrer") >= 0
    assert page.motion_type_combo.findText("Motion - Demurrer") < 0

    page.draft_kind_combo.setCurrentIndex(page.draft_kind_combo.findData(DRAFT_KIND_REPLY))
    assert page.motion_type_combo.findText("Motion to Compel") >= 0
    assert page.intake_settings()["draft_kind"] == DRAFT_KIND_REPLY


def test_motion_drafting_settings_persists_selected_taxonomy_source(qtbot, tmp_path, monkeypatch):
    from icharlotte_core.motion_drafting.taxonomy import DRAFT_KIND_REPLY
    from icharlotte_core.ui.wizard.pages.motion_drafting_page import MotionDraftingSettingsPage

    root = Path(tmp_path) / "MOTION DATABASE"
    dataset = root / "Sample_Pleadings_PDFs"
    source = dataset / "Replies" / "Motion to Compel"
    source.mkdir(parents=True)
    monkeypatch.setenv("ICHARLOTTE_MOTION_DATABASE_ROOT", str(root))

    page = MotionDraftingSettingsPage("/case", "123")
    qtbot.addWidget(page)
    page.draft_kind_combo.setCurrentIndex(page.draft_kind_combo.findData(DRAFT_KIND_REPLY))
    page.motion_type_combo.setCurrentIndex(page.motion_type_combo.findText("Motion to Compel"))

    assert page.intake_settings()["motion_type_source_path"] == str(source)
    assert page.to_dict()["motion_type_source_path"] == str(source)
