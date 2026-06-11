import pytest

from icharlotte_core.chat.legal_research import (
    ChatAuthorityCandidate,
    ChatResearchSource,
)
from icharlotte_core.chat.legal_research_eval import evaluate_candidate_quality


def _candidate(
    *,
    case_name,
    citation,
    snippet,
    source_kind="local_corpus",
    candidate_id=None,
):
    return ChatAuthorityCandidate(
        id=candidate_id or citation,
        proposition="summary judgment burden",
        case_name=case_name,
        citation=citation,
        snippet=snippet,
        sources=[ChatResearchSource(kind=source_kind, label=source_kind)],
    )


def test_evaluate_candidate_quality_reports_expected_citations_and_terms():
    candidates = [
        _candidate(
            case_name="Aguilar v. Atlantic Richfield Co.",
            citation="25 Cal. 4th 826",
            snippet="A defendant moving for summary judgment bears the initial burden.",
        ),
        _candidate(
            case_name="Triable v. Issue",
            citation="99 Cal.App.5th 1",
            snippet="A triable issue prevents summary judgment.",
            source_kind="courtlistener",
        ),
    ]

    report = evaluate_candidate_quality(
        "summary judgment burden triable issue",
        candidates,
        expected_citations=("25 Cal. 4th 826", "100 Cal.App.5th 2"),
        expected_terms=("summary", "judgment", "burden", "triable", "issue"),
        top_n=2,
    )

    assert report.candidate_count == 2
    assert report.expected_citation_hits == ["25 Cal. 4th 826"]
    assert report.expected_citation_misses == ["100 Cal.App.5th 2"]
    assert report.term_coverage_top1 == pytest.approx(3 / 5)
    assert report.term_coverage_top_n == pytest.approx(1.0)
    assert report.source_counts == {"local_corpus": 1, "courtlistener": 1}
    assert report.to_dict()["top_cases"][0]["citation"] == "25 Cal. 4th 826"


def test_evaluate_candidate_quality_handles_empty_candidates():
    report = evaluate_candidate_quality(
        "discovery sanctions",
        [],
        expected_citations=("10 Cal.App.5th 1",),
        expected_terms=("discovery", "sanctions"),
    )

    assert report.candidate_count == 0
    assert report.expected_citation_hits == []
    assert report.expected_citation_misses == ["10 Cal.App.5th 1"]
    assert report.term_coverage_top1 == 0.0
    assert report.term_coverage_top_n == 0.0
    assert report.to_dict()["top_cases"] == []
