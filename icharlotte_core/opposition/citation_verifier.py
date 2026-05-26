"""Citation verification helpers for opposition draft generation."""

from __future__ import annotations

import re
from collections.abc import Iterable
from typing import Any, Callable

from icharlotte_core.legal_research.sources.courtlistener import opinion_url_for_cluster
from icharlotte_core.opposition.models import CitationVerification


STATUS_MAP = {
    400: "invalid",
    300: "ambiguous",
    404: "not_found",
    429: "throttled",
}

_STOP_WORDS = {
    "a",
    "an",
    "and",
    "are",
    "as",
    "at",
    "be",
    "by",
    "for",
    "from",
    "has",
    "have",
    "in",
    "into",
    "is",
    "it",
    "its",
    "of",
    "on",
    "or",
    "that",
    "the",
    "their",
    "to",
    "was",
    "were",
    "when",
    "with",
}


def _first_string(value: Any) -> str:
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, Iterable) and not isinstance(value, (dict, bytes)):
        for item in value:
            text = _first_string(item)
            if text:
                return text
    return ""


def _status_code(record: dict) -> int | None:
    raw_status = record.get("status", record.get("status_code"))
    try:
        return int(raw_status)
    except (TypeError, ValueError):
        return None


def _first_cluster(record: dict) -> dict:
    clusters = record.get("clusters") or record.get("cluster") or []
    if isinstance(clusters, dict):
        return clusters
    if isinstance(clusters, list) and clusters and isinstance(clusters[0], dict):
        return clusters[0]
    return {}


def _cluster_id(record: dict, cluster: dict) -> str:
    value = (
        cluster.get("id")
        or cluster.get("cluster_id")
        or cluster.get("clusterId")
        or record.get("cluster_id")
        or record.get("clusterId")
        or record.get("id")
    )
    return "" if value is None else str(value).strip()


def _warning(record: dict) -> str:
    return _first_string(
        [
            record.get("warning"),
            record.get("warnings"),
            record.get("error"),
            record.get("message"),
        ]
    )


def map_lookup_record(record: dict) -> CitationVerification:
    """Map one CourtListener citation-lookup record to an in-app record."""
    record = record or {}
    normalized_citation = _first_string(record.get("normalized_citations"))
    cluster = _first_cluster(record)
    cluster_id = _cluster_id(record, cluster)
    status_code = _status_code(record)
    status = (
        "exists_support_unconfirmed"
        if status_code == 200
        else STATUS_MAP.get(status_code, "unknown")
    )

    if not cluster and record.get("absolute_url"):
        cluster = {"absolute_url": record.get("absolute_url")}

    return CitationVerification(
        citation_text=_first_string(
            [
                record.get("citation"),
                record.get("citation_text"),
                record.get("original_citation"),
            ]
        ),
        normalized_citation=normalized_citation,
        status=status,
        case_name=_first_string(
            [
                cluster.get("case_name"),
                cluster.get("caseName"),
                record.get("case_name"),
                record.get("caseName"),
            ]
        ),
        court=_first_string([cluster.get("court"), record.get("court")]),
        date=_first_string(
            [
                cluster.get("date_filed"),
                cluster.get("dateFiled"),
                record.get("date_filed"),
                record.get("dateFiled"),
            ]
        ),
        opinion_url=opinion_url_for_cluster(cluster),
        cluster_id=cluster_id,
        warning=_warning(record),
        replacement_candidates=[
            text
            for text in (
                _first_string(candidate)
                for candidate in record.get("replacement_candidates", []) or []
            )
            if text
        ],
    )


def _stem_word(word: str) -> str:
    if len(word) > 4 and word.endswith("ies"):
        return f"{word[:-3]}y"
    if len(word) > 5 and word.endswith("ing"):
        return word[:-3]
    if len(word) > 4 and word.endswith("ed"):
        return word[:-2]
    if len(word) > 3 and word.endswith("s"):
        return word[:-1]
    return word


def _meaningful_words(text: str) -> set[str]:
    words = re.findall(r"[A-Za-z][A-Za-z']+", text.lower())
    return {_stem_word(word) for word in words if word not in _STOP_WORDS and len(word) > 2}


def _passages(opinion_text: str) -> list[str]:
    return [passage for passage, _start, _end in _passage_spans(opinion_text)]


def _passage_spans(opinion_text: str) -> list[tuple[str, int, int]]:
    spans: list[tuple[str, int, int]] = []
    for match in re.finditer(r".*?(?:[.!?](?=\s|$)|$)", opinion_text or "", re.DOTALL):
        raw = match.group(0)
        if not raw:
            continue
        compact = re.sub(r"\s+", " ", raw).strip()
        if compact:
            spans.append((compact, match.start(), match.end()))
    return spans


def find_supporting_passage(proposition: str, opinion_text: str) -> str:
    """Return the best sentence-like passage supporting a proposition."""
    passage, _start, _end = _find_supporting_passage_with_offsets(
        proposition,
        opinion_text,
    )
    return passage


def replacement_candidates(proposition: str, courtlistener) -> list[dict]:
    """Return candidate replacement authorities for an unsupported citation."""
    if not proposition or not proposition.strip():
        return []
    search = getattr(courtlistener, "search_opinions", None)
    if not callable(search):
        return []
    try:
        results = search(proposition, max_results=5)
    except Exception:
        return []

    candidates: list[dict] = []
    for case in results or []:
        candidates.append(
            {
                "case_name": getattr(case, "name", ""),
                "citation": getattr(case, "formatted_citation", "")
                or getattr(case, "citation", ""),
                "court": getattr(case, "court", ""),
                "date": getattr(case, "date", ""),
                "opinion_url": getattr(case, "url", ""),
                "reason": getattr(case, "snippet", ""),
            }
        )
    return candidates


def _find_supporting_passage_with_offsets(
    proposition: str,
    opinion_text: str,
) -> tuple[str, int | None, int | None]:
    proposition_words = _meaningful_words(proposition)
    if len(proposition_words) < 2:
        return "", None, None

    best_passage = ""
    best_start: int | None = None
    best_end: int | None = None
    best_score = 0.0
    best_overlap_count = 0
    for passage, start, end in _passage_spans(opinion_text):
        passage_words = _meaningful_words(passage)
        if not passage_words:
            continue
        overlap_count = len(proposition_words & passage_words)
        score = overlap_count / len(proposition_words)
        if (score, overlap_count) > (best_score, best_overlap_count):
            best_passage = passage
            best_start = start
            best_end = end
            best_score = score
            best_overlap_count = overlap_count

    if best_overlap_count >= 3 and best_score >= 0.4:
        return best_passage, best_start, best_end
    return "", None, None


def verify_citations(
    draft_text: str,
    *,
    citation_propositions: dict[str, str],
    courtlistener,
    support_checker: Callable[[CitationVerification, str, str], bool] | None = None,
) -> list[CitationVerification]:
    """Verify citations found in draft text without changing the draft."""
    records = courtlistener.lookup_citations(draft_text)
    verifications = [map_lookup_record(record) for record in records]
    opinion_text_cache: dict[str, str] = {}

    for verification in verifications:
        if verification.status != "exists_support_unconfirmed":
            continue
        if not verification.cluster_id:
            continue

        proposition = (
            citation_propositions.get(verification.normalized_citation)
            or citation_propositions.get(verification.citation_text)
            or ""
        )
        if not proposition:
            continue

        if verification.cluster_id not in opinion_text_cache:
            opinion_text_cache[verification.cluster_id] = (
                courtlistener.get_opinion_text(verification.cluster_id) or ""
            )
        opinion_text = opinion_text_cache[verification.cluster_id]
        passage, start, end = _find_supporting_passage_with_offsets(
            proposition,
            opinion_text,
        )
        if not passage:
            verification.replacement_candidates = replacement_candidates(
                proposition,
                courtlistener,
            )
            continue

        verification.status = "support_passage_found"
        verification.supporting_passage = passage
        if start is not None and end is not None:
            verification.support_start = start
            verification.support_end = end
        if support_checker and support_checker(verification, proposition, passage):
            verification.status = "verified"

    return verifications
