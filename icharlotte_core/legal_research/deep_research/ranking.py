"""Ranking helpers for deep legal research candidates."""
from __future__ import annotations

import re

from .models import AuthorityCandidate, ParentheticalWeightPolicy, TreatmentSignal


def _clamp_score(value: float) -> float:
    return max(0.0, min(float(value or 0.0), 1.0))


def score_candidate(
    candidate: AuthorityCandidate,
    *,
    parenthetical_policy: ParentheticalWeightPolicy | None = None,
) -> float:
    """Return a deterministic score with parentheticals capped as secondary signal."""
    policy = parenthetical_policy or ParentheticalWeightPolicy.default()
    direct_score = (
        _clamp_score(candidate.semantic_score)
        + _clamp_score(candidate.keyword_score)
        + _clamp_score(candidate.recency_score)
        + _clamp_score(candidate.authority_signal_score)
        + _clamp_score(candidate.source_count_score)
        + _clamp_score(candidate.firm_prior_score)
        - _clamp_score(candidate.negative_signal)
    )
    parenthetical_bonus = min(
        _clamp_score(candidate.parenthetical_match_score),
        _clamp_score(policy.max_score_contribution),
    )
    return round(max(0.0, direct_score + parenthetical_bonus), 6)


def _normalize_parenthetical(text: str) -> str:
    text = re.sub(r"[^a-z0-9]+", " ", (text or "").lower())
    return re.sub(r"\s+", " ", text).strip()


def dampen_duplicate_parentheticals(
    signals: list[TreatmentSignal],
) -> list[TreatmentSignal]:
    """Collapse near-identical parentheticals, keeping the highest-confidence one."""
    by_text: dict[str, TreatmentSignal] = {}
    order: list[str] = []
    for signal in signals:
        key = _normalize_parenthetical(signal.parenthetical_text)
        if not key:
            continue
        current = by_text.get(key)
        if current is None:
            by_text[key] = signal
            order.append(key)
            continue
        if signal.confidence > current.confidence:
            by_text[key] = signal
    return [by_text[key] for key in order]
