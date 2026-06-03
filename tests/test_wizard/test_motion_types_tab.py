"""Tests for the Workbench Motion Types tab (Workbench feature)."""
import pytest

pytest.importorskip("pytestqt")

from icharlotte_core.motion_generation.config import MotionTypeConfig  # noqa: E402
from icharlotte_core.motion_generation.types_registry import MotionTypeRegistry  # noqa: E402
from icharlotte_core.ui.dialogs_motion_types import (  # noqa: E402
    MotionTypesTab,
    _MotionTypeEditDialog,
)


def _tab(qtbot, tmp_path):
    tab = MotionTypesTab(registry_path=str(tmp_path / "mt.json"))
    qtbot.addWidget(tab)
    return tab


def test_seeds_builtins_into_table(qtbot, tmp_path):
    tab = _tab(qtbot, tmp_path)
    ids = {t.type_id for t in tab.registry.list_types()}
    assert {"compel", "demurrer", "strike", "generic"} <= ids
    assert tab.table.rowCount() == len(tab.registry.list_types())


def test_add_type_programmatic_persists(qtbot, tmp_path):
    path = str(tmp_path / "mt.json")
    tab = MotionTypesTab(registry_path=path)
    qtbot.addWidget(tab)
    tab.add_type_programmatic(MotionTypeConfig(
        type_id="protective_order",
        display_name="Motion for Protective Order",
        target_doc_guidance="",
        legal_standard_hint="CCP 2030.090",
        section_plan=["Introduction"],
        placeholder_attachments=[],
    ))
    tab.save()
    reloaded = MotionTypeRegistry.load(path)
    assert reloaded.get("protective_order").display_name == "Motion for Protective Order"


def test_remove_type_programmatic(qtbot, tmp_path):
    tab = _tab(qtbot, tmp_path)
    tab.remove_type_programmatic("strike")
    assert "strike" not in {t.type_id for t in tab.registry.list_types()}


def test_restore_defaults_programmatic(qtbot, tmp_path):
    tab = _tab(qtbot, tmp_path)
    tab.remove_type_programmatic("strike")
    tab.restore_defaults_programmatic()
    assert "strike" in {t.type_id for t in tab.registry.list_types()}


def test_edit_dialog_round_trips_config(qtbot):
    cfg = MotionTypeConfig(
        type_id="compel",
        display_name="Compel",
        target_doc_guidance="guidance",
        legal_standard_hint="LS text",
        section_plan=["Intro", "Argument"],
        placeholder_attachments=["Meet and Confer"],
        analyzer_prompt="AP",
        grounds_prompt="GP",
    )
    dlg = _MotionTypeEditDialog(config=cfg)
    qtbot.addWidget(dlg)
    out = dlg.result_config()
    assert out.type_id == "compel"
    assert out.display_name == "Compel"
    assert out.section_plan == ["Intro", "Argument"]
    assert out.placeholder_attachments == ["Meet and Confer"]
    assert out.legal_standard_hint == "LS text"
    assert out.analyzer_prompt == "AP"


def test_save_reloads_runtime_singleton(qtbot, tmp_path, monkeypatch):
    # Saving the tab must refresh the get_motion_config singleton.
    from icharlotte_core.motion_generation import config as cfgmod

    path = str(tmp_path / "mt.json")
    monkeypatch.setattr(cfgmod, "motion_types_path", lambda: path)
    cfgmod.reload_motion_types()

    tab = MotionTypesTab(registry_path=path)
    qtbot.addWidget(tab)
    tab.update_type_programmatic("compel", display_name="Compel (edited)")
    tab.save()

    assert cfgmod.get_motion_config("compel").display_name == "Compel (edited)"
    monkeypatch.undo()
    cfgmod.reload_motion_types()
