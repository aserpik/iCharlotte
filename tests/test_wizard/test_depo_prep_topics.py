import json
from unittest.mock import MagicMock

import pytest

from Scripts.depo_prep_lib.topics import cluster_topics, TopicsResult


def _payload(topics):
    return json.dumps({"topics": topics})


def _topic(i, **kw):
    base = {
        "id": f"t{i:02d}",
        "title": f"Topic {i}",
        "strategic_note": "note",
        "relevant_digest_refs": [],
        "default_checked": True,
    }
    base.update(kw)
    return base


def test_cluster_topics_returns_topics_list():
    caller = MagicMock()
    caller.call = MagicMock(return_value=_payload([_topic(i) for i in range(1, 11)]))

    result = cluster_topics(
        digests=[{"source_id": "x.pdf", "summary": "s"}],
        llm_caller=caller,
        deponent_name="Jane", deponent_role="P",
        style="discovery", free_text_notes="",
    )
    assert isinstance(result, TopicsResult)
    assert len(result.topics) == 10
    assert result.warning is None


def test_cluster_topics_warns_below_3():
    caller = MagicMock()
    caller.call = MagicMock(return_value=_payload([_topic(1), _topic(2)]))

    result = cluster_topics(
        digests=[{"source_id": "x.pdf", "summary": "s"}],
        llm_caller=caller,
        deponent_name="J", deponent_role="P",
        style="discovery", free_text_notes="",
    )
    assert result.warning is not None
    assert "thin" in result.warning.lower() or "few" in result.warning.lower()


def test_cluster_topics_truncates_above_20():
    caller = MagicMock()
    caller.call = MagicMock(return_value=_payload([_topic(i) for i in range(1, 30)]))

    result = cluster_topics(
        digests=[{"source_id": "x.pdf", "summary": "s"}],
        llm_caller=caller,
        deponent_name="J", deponent_role="P",
        style="discovery", free_text_notes="",
    )
    assert len(result.topics) == 20
    assert result.warning is not None
    assert "truncat" in result.warning.lower()


def test_cluster_topics_raises_on_bad_payload():
    caller = MagicMock()
    caller.call = MagicMock(return_value="not json")

    with pytest.raises(ValueError):
        cluster_topics(digests=[], llm_caller=caller,
                       deponent_name="J", deponent_role="P",
                       style="discovery", free_text_notes="")
