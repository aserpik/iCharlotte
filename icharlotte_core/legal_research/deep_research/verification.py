"""Pure verification helpers for deep legal research packets."""
from __future__ import annotations

import re

from .models import (
    CitationAudit,
    CitationAuditItem,
    CitationAuditStatus,
    ResearchPacket,
    SelectedAuthority,
)


_CASE_CITATION_RE = re.compile(
    r"""
    (?P<full>
        [A-Z][A-Za-z0-9&'.,\- ]+
        \s+v\.\s+
        [A-Z][A-Za-z0-9&'.,\- ]+
        \s+\((?P<year>\d{4})\)\s+
        (?P<reporter>
            \d+\s+
            Cal\.(?:App\.)?(?:\d+(?:th|d|nd|rd)?|[A-Za-z]*)\s+
            \d+
        )
    )
    """,
    re.VERBOSE,
)


def _normalize_ws(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip().lower()


def _normalize_citation(text: str) -> str:
    normalized = re.sub(r"[*_`~\[\]()>.,;:]+", " ", text or "")
    return re.sub(r"[^a-z0-9]+", "", normalized.lower())


def _normalize_case_name(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (text or "").lower())


def contains_verbatim_quote(source_text: str, quote: str) -> bool:
    """Return True when quote appears unchanged after whitespace normalization."""
    normalized_source = _normalize_ws(source_text)
    normalized_quote = _normalize_ws(quote)
    if not normalized_source or not normalized_quote:
        return False
    return normalized_quote in normalized_source


def _is_prompt_safe_authority(authority: SelectedAuthority) -> bool:
    return (authority.verification_status or "").strip().lower() == "verified"


def _known_citation_keys(packet: ResearchPacket) -> set[str]:
    keys: set[str] = set()
    for authority in packet.selected_authorities:
        if not _is_prompt_safe_authority(authority):
            continue
        case_key = _normalize_case_name(authority.case_name)
        reporter_key = _normalize_citation(authority.citation)
        if case_key and reporter_key:
            keys.add(f"{case_key}:{reporter_key}")
    return keys


def audit_citations_against_packet(text: str, packet: ResearchPacket) -> CitationAudit:
    """Audit California-style case citations against verified packet authorities."""
    known_keys = _known_citation_keys(packet)
    items: list[CitationAuditItem] = []

    for match in _CASE_CITATION_RE.finditer(text or ""):
        citation_text = match.group("full").strip()
        case_key = _normalize_case_name(citation_text.rsplit("(", 1)[0])
        reporter_key = _normalize_citation(match.group("reporter"))
        candidate_key = f"{case_key}:{reporter_key}" if case_key and reporter_key else ""
        status = (
            CitationAuditStatus.SUPPORTED
            if candidate_key in known_keys
            else CitationAuditStatus.OFF_PACKET
        )
        detail = (
            "Citation appears in verified research packet."
            if status == CitationAuditStatus.SUPPORTED
            else "Citation was not found in verified research packet authorities."
        )
        items.append(
            CitationAuditItem(
                citation_text=citation_text,
                status=status,
                detail=detail,
            )
        )

    return CitationAudit(items=items)
