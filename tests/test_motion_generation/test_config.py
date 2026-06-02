"""Tests for per-type motion configs (spec #7)."""
from icharlotte_core.motion_generation.config import (
    MOTION_TYPE_CONFIGS,
    MotionTypeConfig,
    get_motion_config,
)


def test_configured_type_set():
    assert set(MOTION_TYPE_CONFIGS.keys()) == {
        "compel",
        "demurrer",
        "strike",
        "generic",
    }


def test_each_config_is_a_motiontypeconfig():
    for cfg in MOTION_TYPE_CONFIGS.values():
        assert isinstance(cfg, MotionTypeConfig)
        assert cfg.type_id
        assert cfg.display_name


def test_compel_legal_standard_cites_discovery_statutes():
    hint = MOTION_TYPE_CONFIGS["compel"].legal_standard_hint
    assert "2030.300" in hint
    assert "2031.310" in hint


def test_demurrer_legal_standard_cites_430_10():
    assert "430.10" in MOTION_TYPE_CONFIGS["demurrer"].legal_standard_hint


def test_strike_legal_standard_cites_435_436():
    hint = MOTION_TYPE_CONFIGS["strike"].legal_standard_hint
    assert "435" in hint
    assert "436" in hint


def test_non_generic_types_have_section_plan():
    for type_id in ("compel", "demurrer", "strike"):
        assert MOTION_TYPE_CONFIGS[type_id].section_plan, type_id


def test_compel_has_meet_confer_and_separate_statement_placeholders():
    attachments = [a.lower() for a in MOTION_TYPE_CONFIGS["compel"].placeholder_attachments]
    assert any("meet" in a for a in attachments)
    assert any("separate statement" in a for a in attachments)


def test_demurrer_and_strike_have_meet_confer_placeholder():
    for type_id in ("demurrer", "strike"):
        attachments = [a.lower() for a in MOTION_TYPE_CONFIGS[type_id].placeholder_attachments]
        assert any("meet" in a for a in attachments), type_id


def test_generic_has_no_placeholder_attachments():
    assert MOTION_TYPE_CONFIGS["generic"].placeholder_attachments == []


def test_get_motion_config_returns_requested_type():
    assert get_motion_config("compel").type_id == "compel"
    assert get_motion_config("demurrer").type_id == "demurrer"


def test_get_motion_config_unknown_falls_back_to_generic():
    assert get_motion_config("motion_for_summary_judgment").type_id == "generic"
    assert get_motion_config("").type_id == "generic"
    assert get_motion_config(None).type_id == "generic"


def test_configs_carry_analyzer_and_grounds_prompts():
    for type_id in ("compel", "demurrer", "strike", "generic"):
        cfg = MOTION_TYPE_CONFIGS[type_id]
        assert cfg.analyzer_prompt
        assert cfg.grounds_prompt
