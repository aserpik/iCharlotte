"""draft_motion / analyze_target read Workbench-editable prompts with fallback."""
import json

from icharlotte_core.motion_generation import analyzer as analyzer_mod
from icharlotte_core.motion_generation import drafter as drafter_mod
from icharlotte_core.motion_generation.analyzer import analyze_target
from icharlotte_core.motion_generation.config import get_motion_config
from icharlotte_core.motion_generation.drafter import draft_motion
from icharlotte_core.motion_generation.prompts import (
    DEFAULT_ANALYZE_TEMPLATE,
    MOTION_DRAFT_PROMPT,
)
from icharlotte_core.opposition.models import MotionMetadata


def _capture_llm(payload):
    calls = {}

    def _cb(system_prompt, user_prompt):
        calls["user"] = user_prompt
        return json.dumps(payload)

    _cb.calls = calls
    return _cb


def _meta():
    return MotionMetadata(
        motion_type="Motion to Compel Further Responses",
        relief_requested="r",
        principal_arguments=["g"],
    )


def test_draft_motion_uses_registered_prompt_when_present(monkeypatch):
    monkeypatch.setattr(
        drafter_mod, "get_prompt",
        lambda agent, pass_name: "OVERRIDE " + MOTION_DRAFT_PROMPT,
    )
    llm = _capture_llm({"title": "X", "body_text": "Body."})
    draft_motion(get_motion_config("compel"), _meta(), [], "T", "", style_exemplars=[], llm_callback=llm)
    assert llm.calls["user"].startswith("OVERRIDE")


def test_draft_motion_falls_back_to_code_default(monkeypatch):
    monkeypatch.setattr(drafter_mod, "get_prompt", lambda agent, pass_name: None)
    llm = _capture_llm({"title": "X", "body_text": "Body."})
    draft = draft_motion(get_motion_config("compel"), _meta(), [], "T", "", style_exemplars=[], llm_callback=llm)
    assert draft.body_text == "Body."
    assert not llm.calls["user"].startswith("OVERRIDE")


def test_analyze_target_uses_registered_prompt_when_present(monkeypatch):
    monkeypatch.setattr(
        analyzer_mod, "get_prompt",
        lambda agent, pass_name: "ANALYZE_OVERRIDE " + DEFAULT_ANALYZE_TEMPLATE,
    )
    llm = _capture_llm({"relief_requested": "x", "principal_arguments": ["y"]})
    analyze_target(get_motion_config("strike"), "TARGET", llm_callback=llm)
    assert llm.calls["user"].startswith("ANALYZE_OVERRIDE")


def test_analyze_target_falls_back_to_code_default(monkeypatch):
    monkeypatch.setattr(analyzer_mod, "get_prompt", lambda agent, pass_name: None)
    llm = _capture_llm({"relief_requested": "x", "principal_arguments": ["y"]})
    analyze_target(get_motion_config("strike"), "PUNITIVE DAMAGES 42", llm_callback=llm)
    # Default template still injects the per-type analyzer guidance + target text.
    assert "irrelevant, false, or improper" in llm.calls["user"]
    assert "PUNITIVE DAMAGES 42" in llm.calls["user"]
