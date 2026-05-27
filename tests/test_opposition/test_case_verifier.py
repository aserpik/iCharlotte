"""Tests for the case verifier — CourtListener lookup + opinion fetch + verdict."""

from __future__ import annotations

import json
import os
import tempfile
from unittest.mock import MagicMock

import pytest

from icharlotte_core.opposition.case_verifier import CaseVerifier
from icharlotte_core.opposition.citation_parser import Citation


@pytest.fixture
def tmp_cache_dir():
    with tempfile.TemporaryDirectory() as tmp:
        yield tmp


def make_citation(
    case_name="Cottini v. Enloe Medical Center",
    year="2014",
    reporter="226 Cal.App.4th 401",
    proposition="Trial courts retain discretion to deny untimely motions.",
):
    return Citation(
        kind="case",
        raw_text=f"*{case_name}* ({year}) {reporter}",
        normalized=f"{case_name} {reporter}",
        proposition=proposition,
        body_offset=0,
        case_name=case_name,
        year=year,
        reporter_citation=reporter,
    )


def test_no_cluster_returns_not_found(tmp_cache_dir):
    cl = MagicMock()
    cl.lookup_citations.return_value = [{"status": 404}]
    llm = MagicMock()

    v = CaseVerifier(courtlistener_client=cl, llm_callback=llm, cache_dir=tmp_cache_dir)
    cv = v.verify(make_citation(case_name="Smith v. Imaginary", reporter="35 Cal.5th 999"))

    assert cv.verdict == "NOT_FOUND"
    assert cv.kind == "case"
    llm.assert_not_called()


def test_cluster_found_triggers_opinion_fetch_and_llm(tmp_cache_dir):
    cl = MagicMock()
    cl.lookup_citations.return_value = [
        {
            "status": 200,
            "clusters": [{"id": 12345, "case_name": "Cottini v. Enloe Medical Center", "absolute_url": "/opinion/12345/cottini/"}],
            "normalized_citations": ["226 Cal.App.4th 401"],
        }
    ]
    cl.get_opinion_text.return_value = "The trial court did not abuse its discretion in denying the late-filed motion."
    llm = MagicMock(return_value='{"verdict": "SUPPORTED", "evidence": "The trial court did not abuse its discretion", "note": "Direct support."}')

    v = CaseVerifier(courtlistener_client=cl, llm_callback=llm, cache_dir=tmp_cache_dir)
    cv = v.verify(make_citation())

    cl.get_opinion_text.assert_called_once_with(12345)
    assert cv.verdict == "SUPPORTED"
    assert cv.cluster_id == "12345"
    assert "courtlistener.com" in cv.opinion_url


def test_cluster_found_but_no_opinion_text_falls_back_to_unverified(tmp_cache_dir):
    cl = MagicMock()
    cl.lookup_citations.return_value = [
        {"status": 200, "clusters": [{"id": 99, "case_name": "X", "absolute_url": "/opinion/99/x/"}]}
    ]
    cl.get_opinion_text.return_value = None
    llm = MagicMock()

    v = CaseVerifier(courtlistener_client=cl, llm_callback=llm, cache_dir=tmp_cache_dir)
    cv = v.verify(make_citation())

    assert cv.verdict == "UNVERIFIED"
    llm.assert_not_called()


def test_opinion_text_cached_after_first_fetch(tmp_cache_dir):
    cl = MagicMock()
    cl.lookup_citations.return_value = [
        {"status": 200, "clusters": [{"id": 12345, "case_name": "Cottini", "absolute_url": "/opinion/12345/"}]}
    ]
    cl.get_opinion_text.return_value = "Opinion text."
    llm = MagicMock(return_value='{"verdict": "PARTIAL", "evidence": "x", "note": "y"}')

    v = CaseVerifier(courtlistener_client=cl, llm_callback=llm, cache_dir=tmp_cache_dir)
    v.verify(make_citation())

    cache_file = os.path.join(tmp_cache_dir, "12345.json")
    assert os.path.exists(cache_file)
    with open(cache_file, "r", encoding="utf-8") as f:
        cached = json.load(f)
    assert cached["text"] == "Opinion text."

    # Second call: get_opinion_text NOT called again.
    cl.get_opinion_text.reset_mock()
    v.verify(make_citation())
    cl.get_opinion_text.assert_not_called()
