"""Tests for the find_replacement_candidates worker function."""

from __future__ import annotations

from unittest.mock import MagicMock

from icharlotte_core.opposition.models import CitationVerification
from icharlotte_core.opposition.verifier import find_replacement_candidates


def test_returns_verified_candidates_only():
    failed = CitationVerification(
        citation_text="Sinaiko Healthcare (2007) 148 Cal.App.4th 390",
        verdict="NOT_SUPPORTED",
        proposition="Serving discovery responses moots a motion to compel.",
        note="Sinaiko addresses waiver, not mootness.",
    )

    # LLM proposes 3 candidates.
    llm = MagicMock(return_value='{"candidates": ['
        '{"citation_text": "*Smith v. Jones* (2010) 50 Cal.4th 100", "kind": "case", "reason": "directly on point"},'
        '{"citation_text": "*Brown v. Davis* (2015) 60 Cal.App.4th 200", "kind": "case", "reason": "supports mootness"},'
        '{"citation_text": "CCP § 2024.020", "kind": "statute", "reason": "deadline"}'
        ']}')

    # Verifier returns SUPPORTED for first, NOT_SUPPORTED for second, NOT_FOUND for third.
    verifier = MagicMock()
    verifier.verify_all.side_effect = lambda cites, **_: [
        CitationVerification(citation_text=c.raw_text, verdict=v)
        for c, v in zip(cites, ["SUPPORTED", "NOT_SUPPORTED", "NOT_FOUND"])
    ]

    candidates = find_replacement_candidates(
        failed_citation=failed,
        verifier=verifier,
        llm_callback=llm,
    )
    # All three returned; caller decides what to do, but verdicts populated.
    assert len(candidates) == 3
    verdicts = [c.verdict for c in candidates]
    assert "SUPPORTED" in verdicts
