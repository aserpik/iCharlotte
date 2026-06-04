"""Tests for the moving-party motion drafter (spec #7).

The opposition drafter rejects drafts that *support* a motion; a moving motion
must NOT be rejected for that reason. Uses a fake llm_callback.
"""
import json

from icharlotte_core.motion_generation.config import get_motion_config
from icharlotte_core.motion_generation.drafter import draft_motion
from icharlotte_core.opposition.models import MotionMetadata


def _meta():
    return MotionMetadata(
        motion_type="Motion to Compel Further Responses",
        moving_party="Defendant",
        relief_requested="Order compelling further responses to RFP Nos. 1-5",
        principal_arguments=[
            "The responses consist of boilerplate objections without merit",
            "No privilege log was served",
        ],
    )


def _fake_llm(payload):
    calls = {}

    def _cb(system_prompt, user_prompt):
        calls["system"] = system_prompt
        calls["user"] = user_prompt
        return json.dumps(payload) if isinstance(payload, dict) else payload

    _cb.calls = calls
    return _cb


def test_draft_motion_returns_body_supporting_the_motion():
    cfg = get_motion_config("compel")
    llm = _fake_llm({
        "title": "Motion to Compel Further Responses",
        "body_text": (
            "This Court should GRANT the motion and compel further responses. "
            "The responding party's boilerplate objections lack merit."
        ),
    })
    draft = draft_motion(
        cfg, _meta(), [], "RESPONSES...", "",
        style_exemplars=[], llm_callback=llm,
    )
    # Must NOT be rejected as "wrong side" for supporting the motion.
    assert draft.body_text
    assert "GRANT" in draft.body_text
    assert not draft.rejection_reason


def test_draft_motion_uses_moving_party_voice_and_legal_standard():
    cfg = get_motion_config("compel")
    llm = _fake_llm({"title": "X", "body_text": "Body."})
    draft_motion(cfg, _meta(), [], "TARGET", "", style_exemplars=[], llm_callback=llm)
    assert "moving party" in llm.calls["system"].lower()
    # Per-type legal standard should be injected.
    assert "2031.310" in llm.calls["user"]
    # Grounds + target text flow into the prompt.
    assert "boilerplate objections" in llm.calls["user"]
    assert "TARGET" in llm.calls["user"]


def test_draft_motion_empty_body_is_rejected():
    cfg = get_motion_config("demurrer")
    llm = _fake_llm({"title": "X", "body_text": "   "})
    draft = draft_motion(cfg, _meta(), [], "C", "", style_exemplars=[], llm_callback=llm)
    assert draft.body_text == ""
    assert draft.rejection_reason


def test_draft_motion_bad_json_is_rejected():
    cfg = get_motion_config("strike")
    llm = _fake_llm("not json at all")
    draft = draft_motion(cfg, _meta(), [], "C", "", style_exemplars=[], llm_callback=llm)
    assert draft.body_text == ""
    assert draft.rejection_reason


def test_draft_motion_prompt_carries_motion_type_and_guardrail():
    from icharlotte_core.motion_generation.drafter import draft_motion
    from icharlotte_core.motion_generation.config import get_motion_config
    from icharlotte_core.opposition.models import MotionMetadata

    captured = {}

    def fake_llm(system_prompt, user_prompt):
        captured["system"] = system_prompt
        captured["user"] = user_prompt
        return '{"title": "T", "body_text": "Argument in favor of granting the motion."}'

    cfg = get_motion_config("generic")
    md = MotionMetadata(motion_type="Motion in Limine to Exclude Witnesses",
                        relief_requested="Exclude", principal_arguments=["A"])
    draft_motion(cfg, md, [], "facts", "", style_exemplars=[], llm_callback=fake_llm)

    blob = (captured["system"] + "\n" + captured["user"]).lower()
    assert "motion in limine to exclude witnesses" in blob
    assert "summary judgment" in blob  # the "do not convert into an MSJ" guardrail
