"""Tests for the argument_research grounding module."""

from __future__ import annotations

from unittest.mock import MagicMock

from icharlotte_core.opposition.argument_research import generate_search_queries


def test_generate_search_queries_parses_json():
    llm = MagicMock(return_value='{"queries": ["discovery cutoff abuse of discretion", "late motion to compel"]}')
    queries = generate_search_queries("The motion is untimely under the discovery cutoff", llm_callback=llm)
    assert queries == ["discovery cutoff abuse of discretion", "late motion to compel"]


def test_generate_search_queries_caps_at_two():
    llm = MagicMock(return_value='{"queries": ["a", "b", "c", "d"]}')
    queries = generate_search_queries("x", llm_callback=llm)
    assert len(queries) == 2


def test_generate_search_queries_handles_garbage():
    llm = MagicMock(return_value="not json")
    assert generate_search_queries("x", llm_callback=llm) == []


from icharlotte_core.opposition.argument_research import select_authorities


def _candidate(cluster_id, text, name="Case v. Name", citation="1 Cal.5th 1"):
    return {"cluster_id": cluster_id, "case_name": name, "citation": citation, "text": text,
            "opinion_url": f"https://www.courtlistener.com/opinion/{cluster_id}/"}


def test_select_authorities_builds_from_metadata_not_model():
    cands = [_candidate("111", "The court held discretion is broad here.", name="A v. B", citation="2 Cal.5th 2")]
    # Model returns a DIFFERENT (hallucinated) citation; we must ignore it and
    # use the candidate's metadata citation.
    llm = MagicMock(return_value='{"selections": [{"id": "111", "supports": "discretion is broad", '
                                 '"passage": "The court held discretion is broad here."}]}')
    out = select_authorities("discretion is broad", cands, argument_text="arg", llm_callback=llm)
    assert len(out) == 1
    assert out[0].cluster_id == "111"
    assert out[0].citation == "2 Cal.5th 2"        # from metadata, not the model
    assert out[0].case_name == "A v. B"
    assert out[0].argument_text == "arg"


def test_select_authorities_drops_unverifiable_passage():
    cands = [_candidate("111", "Real opinion text about timeliness.")]
    # Passage is NOT a substring of the candidate text -> drop it.
    llm = MagicMock(return_value='{"selections": [{"id": "111", "supports": "x", '
                                 '"passage": "A fabricated holding never written."}]}')
    out = select_authorities("x", cands, argument_text="arg", llm_callback=llm)
    assert out == []


def test_select_authorities_ignores_unknown_ids():
    cands = [_candidate("111", "text one")]
    llm = MagicMock(return_value='{"selections": [{"id": "999", "supports": "x", "passage": "text one"}]}')
    out = select_authorities("x", cands, argument_text="arg", llm_callback=llm)
    assert out == []
