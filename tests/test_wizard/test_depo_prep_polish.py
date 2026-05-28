import json
from unittest.mock import MagicMock

import pytest

from Scripts.depo_prep_lib.polish import polish_outline


def _outline(topics):
    return {"topics": topics}


def _t(tid, n):
    return {"topic_id": tid, "questions": [{"n": i, "text": f"Q{i}"} for i in range(1, n + 1)]}


def test_polish_accepts_when_question_counts_match():
    orig = _outline([_t("t01", 3), _t("t02", 2)])
    polished_payload = _outline([
        {"topic_id": "t01", "questions": [
            {"n": 1, "text": "Q1 polished"}, {"n": 2, "text": "Q2 polished"},
            {"n": 3, "text": "Q3 polished"}]},
        {"topic_id": "t02", "questions": [
            {"n": 1, "text": "Q1p"}, {"n": 2, "text": "Q2p"}]},
    ])
    caller = MagicMock()
    caller.call = MagicMock(return_value=json.dumps(polished_payload))

    result = polish_outline(outline=orig, llm_caller=caller)
    assert result["topics"][0]["questions"][0]["text"] == "Q1 polished"


def test_polish_rejects_when_topic_added():
    orig = _outline([_t("t01", 2)])
    polished_payload = _outline([_t("t01", 2), _t("t99", 1)])  # added a topic
    caller = MagicMock()
    caller.call = MagicMock(return_value=json.dumps(polished_payload))

    result = polish_outline(outline=orig, llm_caller=caller)
    # Reverts to original.
    assert len(result["topics"]) == 1


def test_polish_rejects_when_question_dropped():
    orig = _outline([_t("t01", 3)])
    polished_payload = _outline([_t("t01", 2)])  # one Q missing
    caller = MagicMock()
    caller.call = MagicMock(return_value=json.dumps(polished_payload))

    result = polish_outline(outline=orig, llm_caller=caller)
    assert len(result["topics"][0]["questions"]) == 3
    # Original Qs preserved.
    assert result["topics"][0]["questions"][2]["text"] == "Q3"


def test_polish_returns_original_on_llm_error():
    orig = _outline([_t("t01", 2)])
    caller = MagicMock()
    caller.call = MagicMock(side_effect=RuntimeError("boom"))
    result = polish_outline(outline=orig, llm_caller=caller)
    assert result == orig
