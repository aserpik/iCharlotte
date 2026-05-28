"""Tests for semantic + published-only params on search_opinions."""

from __future__ import annotations

from unittest.mock import MagicMock, patch

from icharlotte_core.legal_research.sources.courtlistener import CourtListenerClient


def _fake_response(results):
    resp = MagicMock()
    resp.json.return_value = {"results": results}
    resp.raise_for_status.return_value = None
    resp.status_code = 200
    return resp


def test_semantic_flag_adds_param():
    client = CourtListenerClient(token="x")
    with patch(
        "icharlotte_core.legal_research.sources.courtlistener.requests.get",
        return_value=_fake_response([]),
    ) as mock_get:
        client.search_opinions("discovery cutoff", semantic=True, published_only=True)
    _, kwargs = mock_get.call_args
    params = kwargs["params"]
    assert params["semantic"] == "true"
    assert params["stat_Published"] == "on"
    assert params["type"] == "o"


def test_keyword_default_has_no_semantic_param():
    client = CourtListenerClient(token="x")
    with patch(
        "icharlotte_core.legal_research.sources.courtlistener.requests.get",
        return_value=_fake_response([]),
    ) as mock_get:
        client.search_opinions("discovery cutoff")
    _, kwargs = mock_get.call_args
    assert "semantic" not in kwargs["params"]
