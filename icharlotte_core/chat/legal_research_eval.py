"""Deterministic quality checks for Chat legal-research retrieval."""
from __future__ import annotations

from dataclasses import dataclass, field
import re
from typing import Iterable, Sequence

from icharlotte_core.chat.legal_research import ChatAuthorityCandidate


def _norm_citation(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (value or "").lower())


def _candidate_text(candidate: ChatAuthorityCandidate) -> str:
    return " ".join(
        str(value or "")
        for value in (
            candidate.case_name,
            candidate.citation,
            candidate.snippet,
            candidate.text,
        )
    ).lower()


def _source_counts(candidates: Sequence[ChatAuthorityCandidate]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for candidate in candidates:
        for source in candidate.sources:
            kind = source.kind or "unknown"
            counts[kind] = counts.get(kind, 0) + 1
    return counts


def _term_coverage(
    candidates: Sequence[ChatAuthorityCandidate],
    expected_terms: Sequence[str],
) -> float:
    terms = [term.lower() for term in expected_terms if str(term).strip()]
    if not terms:
        return 1.0
    haystack = " ".join(_candidate_text(candidate) for candidate in candidates)
    if not haystack:
        return 0.0
    hits = sum(1 for term in terms if term in haystack)
    return hits / len(terms)


@dataclass(frozen=True)
class CandidateQualityReport:
    query: str
    candidate_count: int
    top_n: int
    expected_citation_hits: list[str] = field(default_factory=list)
    expected_citation_misses: list[str] = field(default_factory=list)
    term_coverage_top1: float = 0.0
    term_coverage_top_n: float = 0.0
    source_counts: dict[str, int] = field(default_factory=dict)
    top_cases: list[dict[str, str]] = field(default_factory=list)

    def to_dict(self) -> dict[str, object]:
        return {
            "query": self.query,
            "candidate_count": self.candidate_count,
            "top_n": self.top_n,
            "expected_citation_hits": list(self.expected_citation_hits),
            "expected_citation_misses": list(self.expected_citation_misses),
            "term_coverage_top1": self.term_coverage_top1,
            "term_coverage_top_n": self.term_coverage_top_n,
            "source_counts": dict(self.source_counts),
            "top_cases": list(self.top_cases),
        }


def evaluate_candidate_quality(
    query: str,
    candidates: Iterable[ChatAuthorityCandidate],
    *,
    expected_citations: Sequence[str] = (),
    expected_terms: Sequence[str] = (),
    top_n: int = 5,
) -> CandidateQualityReport:
    """Score a retrieved candidate list against a small gold expectation."""
    candidate_list = list(candidates)
    top_n = max(1, int(top_n or 1))
    top_candidates = candidate_list[:top_n]
    top_citation_norms = {
        _norm_citation(candidate.citation)
        for candidate in top_candidates
        if candidate.citation
    }

    hits: list[str] = []
    misses: list[str] = []
    for citation in expected_citations:
        normalized = _norm_citation(citation)
        if normalized and normalized in top_citation_norms:
            hits.append(citation)
        else:
            misses.append(citation)

    top_cases = [
        {
            "id": candidate.id,
            "case_name": candidate.case_name,
            "citation": candidate.citation,
        }
        for candidate in top_candidates
    ]

    return CandidateQualityReport(
        query=query,
        candidate_count=len(candidate_list),
        top_n=top_n,
        expected_citation_hits=hits,
        expected_citation_misses=misses,
        term_coverage_top1=_term_coverage(candidate_list[:1], expected_terms),
        term_coverage_top_n=_term_coverage(top_candidates, expected_terms),
        source_counts=_source_counts(candidate_list),
        top_cases=top_cases,
    )
