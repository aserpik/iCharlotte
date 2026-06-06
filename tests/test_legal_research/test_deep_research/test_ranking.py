from icharlotte_core.legal_research.deep_research import AuthorityCandidate, TreatmentSignal
from icharlotte_core.legal_research.deep_research.ranking import (
    dampen_duplicate_parentheticals,
    score_candidate,
)


def test_parenthetical_bonus_is_capped_at_ten_percent():
    candidate = AuthorityCandidate(
        candidate_id="c1",
        semantic_score=0.40,
        keyword_score=0.10,
        parenthetical_match_score=0.95,
    )

    score = score_candidate(candidate)

    assert score == 0.60


def test_parenthetical_match_cannot_overcome_direct_text_support_gap():
    direct_support = AuthorityCandidate(
        candidate_id="direct",
        semantic_score=0.70,
        keyword_score=0.10,
        parenthetical_match_score=0.00,
        full_text_available=True,
    )
    parenthetical_only = AuthorityCandidate(
        candidate_id="parenthetical",
        semantic_score=0.30,
        keyword_score=0.05,
        parenthetical_match_score=1.00,
        full_text_available=False,
    )

    assert score_candidate(direct_support) > score_candidate(parenthetical_only)


def test_negative_signal_reduces_score():
    candidate = AuthorityCandidate(
        semantic_score=0.50,
        keyword_score=0.10,
        negative_signal=0.20,
    )

    assert score_candidate(candidate) == 0.40


def test_duplicate_parentheticals_are_dampened():
    signals = [
        TreatmentSignal(
            signal_id="1",
            parenthetical_text="holding that notice was required",
            confidence=0.50,
        ),
        TreatmentSignal(
            signal_id="2",
            parenthetical_text="Holding that notice was required.",
            confidence=0.90,
        ),
        TreatmentSignal(
            signal_id="3",
            parenthetical_text="distinguishing cases involving late notice",
            confidence=0.80,
        ),
    ]

    dampened = dampen_duplicate_parentheticals(signals)

    assert [signal.signal_id for signal in dampened] == ["2", "3"]
