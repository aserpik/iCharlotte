"""get_motion_config reads the editable registry singleton (Workbench feature)."""
from icharlotte_core.motion_generation import config as cfgmod
from icharlotte_core.motion_generation.config import (
    get_motion_config,
    reload_motion_types,
)
from icharlotte_core.motion_generation.types_registry import MotionTypeRegistry


def test_get_motion_config_works_from_seed_without_file():
    reload_motion_types()
    assert get_motion_config("compel").type_id == "compel"
    assert get_motion_config("nope").type_id == "generic"


def test_get_motion_config_reflects_registry_edits(tmp_path, monkeypatch):
    path = str(tmp_path / "motion_types.json")
    # Point the singleton at a temp registry, write an edit, and reload.
    monkeypatch.setattr(cfgmod, "motion_types_path", lambda: path)
    reload_motion_types()

    reg = MotionTypeRegistry.load(path)
    reg.update("compel", display_name="Compel (workbench edit)")
    reg.save()

    reload_motion_types()
    assert get_motion_config("compel").display_name == "Compel (workbench edit)"

    # Cleanup: reset singleton for other tests.
    monkeypatch.undo()
    reload_motion_types()
