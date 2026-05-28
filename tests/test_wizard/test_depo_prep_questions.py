import json
from unittest.mock import MagicMock

import pytest

from Scripts.depo_prep_lib.questions import generate_questions_for_topic
from Scripts.depo_prep_lib.schemas import Topic


def _topic(refs=None):
    return {
        "id": "t01", "title": "Pre-existing conditions",
        "strategic_note": "Establish chronic LBP since 2019",
        "relevant_digest_refs": refs or ["med.json#factual_anchors[0]"],
        "default_checked": True, "lawyer_added": False,
    }


def _digest():
    return {
        "source_id": "med.json",
        "factual_anchors": [{"fact": "2019-03 PT intake: chronic LBP", "location": "p.12",
                              "topic_tags": ["injury"]}],
        "deponent_statements": [], "inconsistencies": [],
        "source_kind": "medical_records", "summary": "",
    }


def _llm(payload):
    caller = MagicMock()
    caller.call = MagicMock(return_value=json.dumps(payload))
    return caller


def test_generate_questions_returns_topic_questions():
    payload = {"topic_id": "t01", "questions": [
        {"n": 1, "text": "Before 2024, did you have back pain?"}
    ]}
    caller = _llm(payload)
    result = generate_questions_for_topic(
        topic=_topic(), digests_by_source={"med.json": _digest()},
        llm_caller=caller, deponent_name="Jane", deponent_role="P",
        style="lockdown", free_text_notes="",
        flags={"strategic_note": False, "source_facts": False,
               "impeachment_hook": False, "objection_alts": False},
    )
    assert result["topic_id"] == "t01"
    assert len(result["questions"]) == 1
    assert "error" not in result


def test_generate_questions_returns_error_on_llm_failure():
    caller = MagicMock()
    caller.call = MagicMock(side_effect=RuntimeError("LLM timeout"))
    result = generate_questions_for_topic(
        topic=_topic(), digests_by_source={"med.json": _digest()},
        llm_caller=caller, deponent_name="Jane", deponent_role="P",
        style="discovery", free_text_notes="",
        flags={"strategic_note": False, "source_facts": False,
               "impeachment_hook": False, "objection_alts": False},
    )
    assert "error" in result
    assert "LLM timeout" in result["error"]
    assert result["questions"] == []


def test_generate_questions_resolves_relevant_refs():
    """Only the referenced digest entries are sent in the prompt text."""
    payload = {"topic_id": "t01", "questions": [{"n": 1, "text": "Q"}]}
    caller = _llm(payload)
    generate_questions_for_topic(
        topic=_topic(refs=["med.json#factual_anchors[0]"]),
        digests_by_source={
            "med.json": _digest(),
            "other.json": {"source_id": "other.json", "source_kind": "other",
                            "factual_anchors": [{"fact": "irrelevant", "location": "z",
                                                  "topic_tags": []}],
                            "deponent_statements": [], "inconsistencies": [], "summary": ""},
        },
        llm_caller=caller, deponent_name="J", deponent_role="P",
        style="discovery", free_text_notes="",
        flags={"strategic_note": False, "source_facts": False,
               "impeachment_hook": False, "objection_alts": False},
    )
    # The text payload should contain the chronic LBP fact, not the irrelevant one.
    call = caller.call.call_args
    text = call.kwargs.get("text") or (call.args[1] if len(call.args) > 1 else "")
    assert "chronic LBP" in text
    assert "irrelevant" not in text


def test_generate_questions_lawyer_added_uses_full_digest():
    """Lawyer-added topics with empty refs see all digest entries."""
    payload = {"topic_id": "t99", "questions": [{"n": 1, "text": "Q"}]}
    caller = _llm(payload)
    topic = {"id": "t99", "title": "Custom", "strategic_note": "s",
             "relevant_digest_refs": [], "default_checked": True, "lawyer_added": True}
    generate_questions_for_topic(
        topic=topic,
        digests_by_source={"med.json": _digest()},
        llm_caller=caller, deponent_name="J", deponent_role="P",
        style="discovery", free_text_notes="",
        flags={"strategic_note": False, "source_facts": False,
               "impeachment_hook": False, "objection_alts": False},
    )
    call = caller.call.call_args
    text = call.kwargs.get("text") or (call.args[1] if len(call.args) > 1 else "")
    # Whole digest should be in the payload for lawyer-added topics.
    assert "chronic LBP" in text


def test_generate_questions_passes_flags_to_prompt():
    payload = {"topic_id": "t01", "questions": [{"n": 1, "text": "Q"}]}
    caller = _llm(payload)
    generate_questions_for_topic(
        topic=_topic(), digests_by_source={"med.json": _digest()},
        llm_caller=caller, deponent_name="J", deponent_role="P",
        style="discovery", free_text_notes="",
        flags={"strategic_note": True, "source_facts": True,
               "impeachment_hook": True, "objection_alts": True},
    )
    prompt = caller.call.call_args.kwargs.get("prompt") or caller.call.call_args.args[0]
    assert "purpose" in prompt.lower()
    assert "impeachment" in prompt.lower()
