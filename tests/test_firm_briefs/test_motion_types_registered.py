from icharlotte_core.motion_generation.config import list_motion_types, get_motion_config

EXPECTED = {"msj", "ex_parte", "ime", "gfs", "dismiss", "leave", "consolidate",
            "quash", "sanctions", "continue_trial", "protective_order",
            "compel", "demurrer", "strike"}


def test_common_types_registered():
    ids = {c.type_id for c in list_motion_types()}
    assert EXPECTED.issubset(ids)


def test_new_types_have_display_and_legal_standard():
    for tid in ["msj", "ex_parte", "ime", "gfs"]:
        cfg = get_motion_config(tid)
        assert cfg.type_id == tid
        assert cfg.display_name
        assert cfg.legal_standard_hint
        assert cfg.section_plan  # non-empty spine


def test_unknown_still_generic():
    assert get_motion_config("totally-unknown-xyz").type_id == "generic"
