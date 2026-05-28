import json
from unittest.mock import MagicMock

import pytest

from Scripts.depo_prep_lib.merge import dedup_and_coverage, apply_dedup


def test_apply_dedup_drops_marked_questions_and_renumbers():
    topic_outputs = [
        {"topic_id": "t01", "questions": [
            {"n": 1, "text": "Q1"}, {"n": 2, "text": "Q2"}, {"n": 3, "text": "Q3"}]},
        {"topic_id": "t02", "questions": [
            {"n": 1, "text": "DupOfT01Q2"}, {"n": 2, "text": "Q5"}]},
    ]
    dedup = {
        "duplicates": [{"keep": "t01.q2", "drop": "t02.q1", "reason": "same"}],
        "coverage_gaps": [], "renumber_after_dedup": True,
    }
    result = apply_dedup(topic_outputs, dedup)
    t02 = next(t for t in result if t["topic_id"] == "t02")
    assert len(t02["questions"]) == 1
    assert t02["questions"][0]["n"] == 1
    assert t02["questions"][0]["text"] == "Q5"


def test_apply_dedup_handles_missing_topic_or_question_gracefully():
    topic_outputs = [{"topic_id": "t01", "questions": [{"n": 1, "text": "Q"}]}]
    dedup = {"duplicates": [{"keep": "t99.q1", "drop": "t01.q99", "reason": "?"}],
             "coverage_gaps": [], "renumber_after_dedup": True}
    # Should not raise; nothing dropped.
    result = apply_dedup(topic_outputs, dedup)
    assert len(result[0]["questions"]) == 1


def test_dedup_and_coverage_returns_parsed_dedup():
    caller = MagicMock()
    caller.call = MagicMock(return_value=json.dumps({
        "duplicates": [{"keep": "t01.q1", "drop": "t02.q1", "reason": "same"}],
        "coverage_gaps": ["Missing X"],
        "renumber_after_dedup": True,
    }))

    topic_outputs = [
        {"topic_id": "t01", "questions": [{"n": 1, "text": "Q1"}]},
        {"topic_id": "t02", "questions": [{"n": 1, "text": "Q1 again"}]},
    ]
    dedup = dedup_and_coverage(
        topic_outputs=topic_outputs,
        digests_by_source={"x.json": {"summary": "..."}},
        llm_caller=caller,
    )
    assert dedup["duplicates"][0]["drop"] == "t02.q1"
    assert dedup["coverage_gaps"] == ["Missing X"]


def test_dedup_returns_empty_on_llm_error():
    """Dedup must not fail Phase 2 if the LLM throws."""
    caller = MagicMock()
    caller.call = MagicMock(side_effect=RuntimeError("boom"))
    result = dedup_and_coverage(
        topic_outputs=[], digests_by_source={}, llm_caller=caller,
    )
    assert result == {"duplicates": [], "coverage_gaps": [],
                       "renumber_after_dedup": False}
