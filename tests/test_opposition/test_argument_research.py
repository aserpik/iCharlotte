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
