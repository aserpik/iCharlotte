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


from icharlotte_core.opposition.argument_research import research_argument
from icharlotte_core.legal_research.models import CaseResult


def _case(cluster_id, name="A v. B", citation="2 Cal.5th 2"):
    return CaseResult(name=name, citation=citation, date="2015-01-01", court="cal",
                      snippet="snip", url=f"https://cl/opinion/{cluster_id}/", cluster_id=cluster_id)


def test_research_argument_happy_path(tmp_path):
    cl = MagicMock()
    cl.search_opinions.return_value = [_case(111)]
    cl.get_opinion_text.return_value = "The court held discretion is broad here."

    query_llm = MagicMock(return_value='{"queries": ["discovery cutoff"]}')
    rerank_llm = MagicMock(return_value='{"selections": [{"id": "111", "supports": "broad discretion", '
                                        '"passage": "The court held discretion is broad here."}]}')

    out = research_argument(
        "The motion is untimely", cl_client=cl,
        query_llm=query_llm, rerank_llm=rerank_llm, cache_dir=str(tmp_path),
    )
    assert len(out) == 1
    assert out[0].cluster_id == "111"
    assert out[0].argument_text == "The motion is untimely"


def test_research_argument_unions_semantic_and_keyword(tmp_path):
    cl = MagicMock()
    cl.search_opinions.return_value = [_case(111), _case(222)]
    cl.get_opinion_text.return_value = "text"
    query_llm = MagicMock(return_value='{"queries": ["q1"]}')
    rerank_llm = MagicMock(return_value='{"selections": []}')

    research_argument("arg", cl_client=cl, query_llm=query_llm, rerank_llm=rerank_llm, cache_dir=str(tmp_path))

    # The plain argument is prepended as a natural-language query, so two queries
    # run ("arg" + "q1"), each firing a semantic and a keyword pass = 4 calls.
    assert cl.search_opinions.call_count == 4
    first_kwargs = cl.search_opinions.call_args_list[0].kwargs
    assert first_kwargs.get("semantic") is True  # semantic pass fires first


def test_research_argument_empty_retries_once(tmp_path):
    cl = MagicMock()
    cl.search_opinions.return_value = []
    query_llm = MagicMock(return_value='{"queries": ["alpha beta gamma"]}')
    rerank_llm = MagicMock(return_value='{"selections": []}')

    # Multi-word argument so the broaden-retry on the (prepended) natural-language
    # lead query actually fires.
    out = research_argument("the motion is untimely", cl_client=cl, query_llm=query_llm,
                            rerank_llm=rerank_llm, cache_dir=str(tmp_path))
    assert out == []
    # Two queries ("the motion is untimely" + "alpha beta gamma") x (semantic+keyword)
    # = 4 calls, plus one broadened retry of the lead query (2 calls) = 6.
    assert cl.search_opinions.call_count == 6


def test_research_argument_stamps_goodlaw_signals(tmp_path):
    cl = MagicMock()
    cl.search_opinions.return_value = [_case(111)]
    cl.get_opinion_text.return_value = "The court held discretion is broad here."
    cl.get_authority_signals.return_value = {"citation_count": 42, "latest_citing_year": "2022"}
    query_llm = MagicMock(return_value='{"queries": ["q"]}')
    rerank_llm = MagicMock(return_value='{"selections": [{"id": "111", "supports": "s", '
                                        '"passage": "The court held discretion is broad here."}]}')

    out = research_argument("arg", cl_client=cl, query_llm=query_llm, rerank_llm=rerank_llm, cache_dir=str(tmp_path))
    assert out[0].citation_count == 42
    assert out[0].latest_citing_year == "2022"


from icharlotte_core.opposition.argument_research import research_arguments


def test_research_arguments_runs_each_and_emits_progress(tmp_path):
    cl = MagicMock()
    cl.search_opinions.return_value = [_case(111)]
    cl.get_opinion_text.return_value = "The court held discretion is broad here."
    query_llm = MagicMock(return_value='{"queries": ["q"]}')
    rerank_llm = MagicMock(return_value='{"selections": [{"id": "111", "supports": "s", '
                                        '"passage": "The court held discretion is broad here."}]}')
    messages = []

    out = research_arguments(
        ["arg one", "arg two"], cl_client=cl, query_llm=query_llm, rerank_llm=rerank_llm,
        max_workers=2, on_progress=messages.append, cache_dir=str(tmp_path),
    )
    # One authority per argument.
    assert len(out) == 2
    assert {a.argument_text for a in out} == {"arg one", "arg two"}
    assert len(messages) >= 2


def test_research_arguments_empty_list():
    assert research_arguments([], cl_client=MagicMock(), query_llm=MagicMock(), rerank_llm=MagicMock()) == []
