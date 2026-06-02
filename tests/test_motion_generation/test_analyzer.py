"""Tests for the Generate Motion target-document analyzer (spec #7).

Uses a fake llm_callback so no network/LLM is required.
"""
import json

from icharlotte_core.motion_generation.analyzer import (
    analyze_target,
    outline_from_config,
)
from icharlotte_core.motion_generation.config import get_motion_config


def _fake_llm(payload: dict):
    """Return an llm_callback that always responds with `payload` as JSON, and
    records the prompts it was called with."""
    calls = {}

    def _cb(system_prompt, user_prompt):
        calls["system"] = system_prompt
        calls["user"] = user_prompt
        return json.dumps(payload)

    _cb.calls = calls
    return _cb


def test_analyze_target_extracts_grounds_and_relief():
    cfg = get_motion_config("compel")
    llm = _fake_llm({
        "relief_requested": "Order compelling further responses to RFP Nos. 1-5",
        "principal_arguments": [
            "Responses are evasive boilerplate objections",
            "No privilege log was served",
        ],
    })
    md = analyze_target(cfg, "RESPONSE TO RFP: objection, vague...", llm_callback=llm)
    assert "further responses" in md.relief_requested.lower()
    assert len(md.principal_arguments) == 2


def test_analyze_target_sets_motion_type_from_config():
    cfg = get_motion_config("demurrer")
    llm = _fake_llm({"relief_requested": "x", "principal_arguments": ["y"]})
    md = analyze_target(cfg, "COMPLAINT...", llm_callback=llm)
    assert md.motion_type == cfg.display_name


def test_analyze_target_sends_target_text_and_type_prompt_to_llm():
    cfg = get_motion_config("strike")
    llm = _fake_llm({"relief_requested": "x", "principal_arguments": ["y"]})
    analyze_target(cfg, "PUNITIVE DAMAGES ALLEGATION 42", llm_callback=llm)
    user = llm.calls["user"]
    assert "PUNITIVE DAMAGES ALLEGATION 42" in user
    # The per-type analyzer guidance must steer the prompt.
    assert "irrelevant, false, or improper" in user


def test_analyze_target_tolerates_bad_json():
    cfg = get_motion_config("compel")

    def _bad(system_prompt, user_prompt):
        return "Sorry, I cannot help with that."

    md = analyze_target(cfg, "x", llm_callback=_bad)
    assert md.motion_type == cfg.display_name
    assert md.principal_arguments == []


def test_outline_from_config_uses_section_plan():
    cfg = get_motion_config("demurrer")
    nodes = outline_from_config(cfg)
    texts = [n.text for n in nodes]
    assert texts == cfg.section_plan
    # Normalized nodes carry ids.
    assert all(n.id for n in nodes)


def test_outline_from_config_generic_still_has_sections():
    nodes = outline_from_config(get_motion_config("generic"))
    assert [n.text for n in nodes] == get_motion_config("generic").section_plan
