"""Tests for the editable motion-types registry (Workbench feature)."""
import json

from icharlotte_core.motion_generation.config import (
    BUILTIN_SEED,
    MotionTypeConfig,
)
from icharlotte_core.motion_generation.types_registry import MotionTypeRegistry


def test_load_missing_file_seeds_builtins_without_writing(tmp_path):
    path = str(tmp_path / "motion_types.json")
    reg = MotionTypeRegistry.load(path)
    ids = {t.type_id for t in reg.list_types()}
    assert {"compel", "demurrer", "strike", "generic"} <= ids
    # Reading must not create the file (no surprise side effects).
    assert not (tmp_path / "motion_types.json").exists()


def test_config_to_dict_from_dict_round_trip():
    cfg = BUILTIN_SEED["compel"]
    data = cfg.to_dict()
    back = MotionTypeConfig.from_dict(data)
    assert back == cfg


def test_save_then_load_round_trips(tmp_path):
    path = str(tmp_path / "mt.json")
    reg = MotionTypeRegistry.load(path)
    reg.add(MotionTypeConfig(
        type_id="protective_order",
        display_name="Motion for Protective Order",
        target_doc_guidance="Add the discovery at issue.",
        legal_standard_hint="CCP 2030.090.",
        section_plan=["Introduction", "Argument", "Conclusion"],
        placeholder_attachments=["Meet and Confer Declaration"],
    ))
    reg.save()
    assert (tmp_path / "mt.json").exists()

    reloaded = MotionTypeRegistry.load(path)
    cfg = reloaded.get("protective_order")
    assert cfg.display_name == "Motion for Protective Order"
    assert cfg.section_plan == ["Introduction", "Argument", "Conclusion"]


def test_get_unknown_returns_generic(tmp_path):
    reg = MotionTypeRegistry.load(str(tmp_path / "mt.json"))
    assert reg.get("does_not_exist").type_id == "generic"


def test_update_existing_type(tmp_path):
    reg = MotionTypeRegistry.load(str(tmp_path / "mt.json"))
    assert reg.update("compel", display_name="Compel (edited)") is True
    assert reg.get("compel").display_name == "Compel (edited)"


def test_remove_type(tmp_path):
    reg = MotionTypeRegistry.load(str(tmp_path / "mt.json"))
    assert reg.remove("strike") is True
    assert "strike" not in {t.type_id for t in reg.list_types()}


def test_restore_defaults(tmp_path):
    reg = MotionTypeRegistry.load(str(tmp_path / "mt.json"))
    reg.remove("strike")
    reg.update("compel", display_name="x")
    reg.restore_defaults()
    ids = {t.type_id for t in reg.list_types()}
    assert "strike" in ids
    assert reg.get("compel").display_name == BUILTIN_SEED["compel"].display_name


def test_restore_single_default(tmp_path):
    reg = MotionTypeRegistry.load(str(tmp_path / "mt.json"))
    reg.update("demurrer", legal_standard_hint="broken")
    assert reg.restore_default("demurrer") is True
    assert reg.get("demurrer").legal_standard_hint == BUILTIN_SEED["demurrer"].legal_standard_hint


def test_saved_json_is_a_list_of_type_dicts(tmp_path):
    path = str(tmp_path / "mt.json")
    reg = MotionTypeRegistry.load(path)
    reg.save()
    data = json.loads((tmp_path / "mt.json").read_text(encoding="utf-8"))
    assert isinstance(data["types"], list)
    assert all("type_id" in t and "legal_standard_hint" in t for t in data["types"])
