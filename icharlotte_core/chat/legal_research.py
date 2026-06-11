"""Chat-tab legal research orchestration.

This module is intentionally Qt-free. The Chat tab owns widgets and persistence;
this module owns source settings, retrieval, quote verification, and prompt
packet formatting.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
import html as html_lib
import json
import os
import re
from typing import Any, Callable, Iterable, Optional
from urllib.parse import urlparse

from icharlotte_core.legal_research.local_corpus.cap_urls import (
    is_broken_cap_static_id_url,
    repair_broken_cap_url,
)
from icharlotte_core.legal_research.models import CaseResult

try:
    from icharlotte_core.legal_research.sources.courtlistener import CourtListenerClient
except Exception:
    CourtListenerClient = None

LLMCallback = Callable[[str, str], str]
StatusCallback = Optional[Callable[[str], None]]
DebugCallback = Optional[Callable[..., None]]


class ChatResearchError(RuntimeError):
    """Raised when research mode cannot produce a verified research basis."""


class CourtListenerMode(str, Enum):
    OFF = "off"
    FALLBACK_CURRENT_LAW = "fallback_current_law"
    ALWAYS_SEARCH = "always_search"


class ChatResearchOutputMode(str, Enum):
    QUICK_ANSWER = "quick_answer"
    RESEARCH_MEMO = "research_memo"


def _bool_value(value: Any, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"1", "true", "yes", "on"}:
            return True
        if lowered in {"0", "false", "no", "off"}:
            return False
    if value is None:
        return default
    return bool(value)


@dataclass(frozen=True)
class ChatResearchSettings:
    firm_authority: bool = True
    local_corpus: bool = True
    courtlistener_mode: CourtListenerMode = CourtListenerMode.FALLBACK_CURRENT_LAW
    output_mode: ChatResearchOutputMode = ChatResearchOutputMode.QUICK_ANSWER

    @classmethod
    def default(cls) -> "ChatResearchSettings":
        return cls()

    @classmethod
    def from_values(
        cls,
        *,
        firm_authority: Any = True,
        local_corpus: Any = True,
        courtlistener_mode: Any = CourtListenerMode.FALLBACK_CURRENT_LAW,
        output_mode: Any = ChatResearchOutputMode.QUICK_ANSWER,
    ) -> "ChatResearchSettings":
        if isinstance(courtlistener_mode, CourtListenerMode):
            mode = courtlistener_mode
        else:
            try:
                mode = CourtListenerMode(str(courtlistener_mode))
            except ValueError:
                mode = CourtListenerMode.FALLBACK_CURRENT_LAW
        if isinstance(output_mode, ChatResearchOutputMode):
            selected_output_mode = output_mode
        else:
            try:
                selected_output_mode = ChatResearchOutputMode(str(output_mode))
            except ValueError:
                selected_output_mode = ChatResearchOutputMode.QUICK_ANSWER
        return normalize_settings(
            cls(
                firm_authority=_bool_value(firm_authority, True),
                local_corpus=_bool_value(local_corpus, True),
                courtlistener_mode=mode,
                output_mode=selected_output_mode,
            )
        )


def normalize_settings(settings: ChatResearchSettings) -> ChatResearchSettings:
    if (
        not settings.firm_authority
        and not settings.local_corpus
        and settings.courtlistener_mode == CourtListenerMode.FALLBACK_CURRENT_LAW
    ):
        return ChatResearchSettings(
            firm_authority=False,
            local_corpus=False,
            courtlistener_mode=CourtListenerMode.ALWAYS_SEARCH,
            output_mode=settings.output_mode,
        )
    return settings


@dataclass
class ChatResearchSource:
    kind: str
    label: str
    verification: str = "unknown"
    reference: str = ""


@dataclass
class ChatAuthorityCandidate:
    id: str
    proposition: str
    case_name: str
    citation: str
    year: str = ""
    court: str = ""
    url: str = ""
    text: str = ""
    snippet: str = ""
    snippet_page_label: str = ""
    sources: list[ChatResearchSource] = field(default_factory=list)
    citation_count: int | None = None
    latest_citing_year: str = ""


@dataclass
class ChatSelectedAuthority:
    id: str
    proposition: str
    case_name: str
    citation: str
    year: str = ""
    court: str = ""
    url: str = ""
    reason: str = ""
    supports: str = ""
    quote: str = ""
    caveat: str = ""
    verification: str = "verified"
    sources: list[ChatResearchSource] = field(default_factory=list)

    @property
    def formatted_citation(self) -> str:
        if self.year:
            return f"{self.case_name} ({self.year}) {self.citation}".strip()
        return f"{self.case_name} {self.citation}".strip()


RERANK_SELECT_PROMPT = """You are selecting California legal authorities for a Chat answer.

You are given one research proposition and candidate opinions or firm-authority sources already retrieved by search.

Select the 1 to 4 candidates that best support the proposition. For each selected candidate return:
- id: candidate id exactly as shown;
- reason: one sentence explaining why this authority was selected;
- supports: one sentence stating the rule or proposition the authority supports;
- quote: a short verbatim quote copied from the candidate excerpt;
- caveat: a short limitation or empty string.

Return strict JSON only: {"selections":[{"id":"cap:1","reason":"This authority states the governing rule.","supports":"The case supports the requested legal proposition.","quote":"A short verbatim quote from the candidate excerpt.","caveat":""}]}.
Never invent a quote. The quote must be copied from the candidate excerpt.
"""

RESEARCH_PROMPT_INSTRUCTION = """LEGAL RESEARCH MODE IS ENABLED.

You must cite only authorities in [CHAT LEGAL RESEARCH AUTHORITY].
Do not add citations from memory. Do not invent, recall, or add outside citations.
If the selected authorities do not support a requested proposition, say that the selected sources did not provide support.
Write a direct answer to the user's question using the verified authorities below.
"""

RESEARCH_MEMO_PROMPT_INSTRUCTION = """Research Memo mode is enabled.

Organize the main answer as polished attorney work product. Use this structure unless the user's prompt clearly requests a different structure:
1. Summary
2. Governing Rule
3. Best Supporting Cases
4. Limitations / Adverse Authority
5. Suggested Argument Framing

Use the Presentation role hints in the authority block to organize the discussion. Do not repeat the full Legal Research Basis appendix inside the main answer. The application will append the audit trail separately.
"""


@dataclass
class ChatResearchPacket:
    query: str
    settings: ChatResearchSettings
    propositions: list[str] = field(default_factory=list)
    searches: list[str] = field(default_factory=list)
    selected_authorities: list[ChatSelectedAuthority] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def get_known_case_names(self) -> list[str]:
        return [a.case_name for a in self.selected_authorities if a.case_name]

    def format_authority_block(self) -> str:
        lines = ["[CHAT LEGAL RESEARCH AUTHORITY]"]
        if self.propositions:
            lines.append("Research questions:")
            for idx, prop in enumerate(self.propositions, start=1):
                lines.append(f"{idx}. {prop}")
        if self.searches:
            lines.append("")
            lines.append("Searches run:")
            for item in self.searches:
                lines.append(f"- {item}")
        if self.selected_authorities:
            lines.append("")
            lines.append("Selected authorities:")
            for authority in self.selected_authorities:
                source_labels = ", ".join(s.label for s in authority.sources) or "unknown source"
                lines.append(f"- {authority.formatted_citation}")
                lines.append(f"  Proposition: {authority.proposition}")
                if authority.reason:
                    lines.append(
                        f"  Untrusted selector reason: {_sanitize_selector_text(authority.reason)}"
                    )
                if authority.supports:
                    lines.append(
                        f"  Untrusted selector support summary: {_sanitize_selector_text(authority.supports)}"
                    )
                lines.append(f"  Source: {source_labels}")
                if self.settings.output_mode == ChatResearchOutputMode.RESEARCH_MEMO:
                    lines.append(
                        f"  Presentation role: {_authority_presentation_role(authority)}"
                    )
                if authority.quote:
                    lines.append(f"  Quote: \"{_sanitize_selector_text(authority.quote)}\"")
                if authority.caveat:
                    lines.append(
                        f"  Untrusted selector caveat: {_sanitize_selector_text(authority.caveat)}"
                    )
                view_url = _authority_view_url(authority)
                if view_url:
                    lines.append(f"  URL: {view_url}")
        if self.warnings:
            lines.append("")
            lines.append("Warnings:")
            for warning in self.warnings:
                lines.append(f"- {warning}")
        lines.append("[/CHAT LEGAL RESEARCH AUTHORITY]")
        return "\n".join(lines)

    def format_research_basis_html(self) -> list[str]:
        lines = ["<b>Legal Research Basis</b>"]
        if self.searches:
            lines.append("<i>Searches run:</i>")
            for item in self.searches:
                lines.append(f"- {_escape_html(item)}")
        if self.selected_authorities:
            lines.append("<i>Authorities selected:</i>")
            for authority in self.selected_authorities:
                source_labels = ", ".join(s.label for s in authority.sources) or "unknown source"
                line = (
                    f"- <b>{_escape_html(authority.formatted_citation)}</b> "
                    f"[{_escape_html(source_labels)}]"
                )
                safe_url = _authority_view_url(authority)
                if safe_url:
                    line += f' <a href="{_escape_html(safe_url)}">View</a>'
                lines.append(line)
                if authority.reason:
                    lines.append(f"  Why: {_escape_html(authority.reason)}")
                if authority.quote:
                    lines.append(f"  Quote: &quot;{_escape_html(authority.quote)}&quot;")
        if self.warnings:
            lines.append("<i>Warnings:</i>")
            for warning in self.warnings:
                lines.append(f"- {_escape_html(warning)}")
        return lines

    def build_augmented_system_prompt(self, base_system_prompt: str) -> str:
        parts = [
            base_system_prompt,
            RESEARCH_PROMPT_INSTRUCTION,
        ]
        if self.settings.output_mode == ChatResearchOutputMode.RESEARCH_MEMO:
            parts.append(RESEARCH_MEMO_PROMPT_INSTRUCTION)
        parts.append(self.format_authority_block())
        return "\n\n".join(parts)


def _loads_json(text: str) -> dict[str, Any]:
    if not isinstance(text, str):
        return {}
    cleaned = text.strip()
    fenced = re.fullmatch(r"```(?:json)?\s*(.*?)\s*```", cleaned, re.DOTALL | re.I)
    if fenced:
        cleaned = fenced.group(1).strip()
    try:
        data = json.loads(cleaned)
    except (TypeError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def _normalize_ws(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip().lower()


def _is_california_supreme_court(authority: ChatSelectedAuthority) -> bool:
    court = (authority.court or "").lower()
    citation = authority.citation or ""
    if "supreme" in court:
        return True
    return (
        bool(re.search(r"\bCal\.\s*(?:2d|3d|4th|5th)\b", citation))
        and "Cal.App" not in citation
    )


def _authority_presentation_role(authority: ChatSelectedAuthority) -> str:
    caveat = (authority.caveat or "").lower()
    combined = " ".join(
        [
            authority.reason or "",
            authority.supports or "",
            authority.proposition or "",
            caveat,
        ]
    ).lower()
    if any(
        term in combined
        for term in ("adverse", "contrary", "distinguish", "cuts against")
    ):
        return "adverse"
    if caveat:
        return "limiting"
    if _is_california_supreme_court(authority):
        return "foundational"
    if any(term in combined for term in ("directly", "squarely", "direct support")):
        return "direct"
    return "background"


def _escape_html(text: str) -> str:
    return html_lib.escape(str(text or ""), quote=True)


def _safe_http_url(url: str) -> str:
    value = str(url or "").strip()
    if not value:
        return ""
    parsed = urlparse(value)
    if parsed.scheme.lower() not in {"http", "https"}:
        return ""
    if not parsed.netloc:
        return ""
    return value


def _authority_view_url(authority: ChatSelectedAuthority) -> str:
    safe_url = _safe_http_url(authority.url)
    is_local_cap = (
        str(authority.id or "").lower().startswith("cap:")
        and any(source.kind == "local_corpus" for source in authority.sources)
    )
    if is_local_cap and (not safe_url or is_broken_cap_static_id_url(safe_url)):
        return _safe_http_url(repair_broken_cap_url(safe_url, authority.citation))
    return safe_url


PROMPT_CONTROL_MARKER_RE = re.compile(
    r"[\[\(<{]\s*/?\s*(?:system|assistant|user|developer|instruction|instructions|prompt|"
    r"chat\s+legal\s+research\s+authority)[^\]\)>}]*[\]\)>}]",
    re.I,
)
PROMPT_CONTROL_PHRASE_RE = re.compile(
    r"\b(?:ignore|disregard|override)\s+(?:all\s+)?(?:previous|prior|earlier|above)\s+"
    r"(?:instructions?|prompts?|messages?)\b",
    re.I,
)


def _sanitize_selector_text(text: str, *, max_chars: int = 600) -> str:
    cleaned = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]+", " ", str(text or ""))
    cleaned = PROMPT_CONTROL_MARKER_RE.sub("[removed prompt-control marker]", cleaned)
    cleaned = PROMPT_CONTROL_PHRASE_RE.sub("[removed prompt-control phrase]", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if len(cleaned) > max_chars:
        return cleaned[: max_chars - 3].rstrip() + "..."
    return cleaned


def _relevant_excerpt(text: str, proposition: str, *, max_chars: int = 3500) -> str:
    text = text or ""
    if len(text) <= max_chars:
        return text
    terms = {
        term
        for term in re.findall(r"[a-z]{4,}", (proposition or "").lower())
        if term not in {"that", "this", "with", "from", "court", "case", "rule"}
    }
    if not terms:
        return text[:max_chars]
    lower_text = text.lower()
    hits: list[int] = []
    for term in terms:
        start = 0
        while True:
            idx = lower_text.find(term, start)
            if idx < 0:
                break
            hits.append(idx)
            start = idx + len(term)
    if not hits:
        return text[:max_chars]
    hits.sort()
    import bisect

    best_start = 0
    best_score = -1
    for hit in hits:
        window_start = max(0, hit - 250)
        window_end = window_start + max_chars
        score = bisect.bisect_right(hits, window_end) - bisect.bisect_left(
            hits,
            window_start,
        )
        if score > best_score:
            best_score = score
            best_start = window_start
    end = min(len(text), best_start + max_chars)
    prefix = "[excerpt begins mid-opinion] " if best_start > 0 else ""
    suffix = " [excerpt continues]" if end < len(text) else ""
    return f"{prefix}{text[best_start:end]}{suffix}"


def _format_candidates(candidates: list[ChatAuthorityCandidate], proposition: str) -> str:
    blocks: list[str] = []
    for candidate in candidates:
        source_labels = ", ".join(s.label for s in candidate.sources)
        text = candidate.text or candidate.snippet
        excerpt = _relevant_excerpt(text, proposition)
        lines = [
            f"[{candidate.id}] {candidate.case_name}, {candidate.citation}",
            f"Sources: {source_labels}",
        ]
        if candidate.snippet:
            if candidate.snippet_page_label:
                lines.append(f"Snippet page label: {candidate.snippet_page_label}")
            lines.extend(["Best matching snippet:", candidate.snippet])
        lines.extend(["Opinion excerpt:", excerpt])
        blocks.append("\n".join(lines))
    return "\n\n".join(blocks)


PROPOSITION_EXTRACTION_PROMPT = """You are extracting focused California legal research questions for a litigation attorney.

Read the user's request and context. Return 1 to 5 propositions or legal questions that should be researched.

Rules:
- Focus on California legal doctrine, elements, defenses, standards, and procedural rules.
- Do not include party names unless needed to identify a specific case.
- Do not include drafting instructions or formatting instructions.
- Return strict JSON only: {"propositions":["landlord duty to repair stairs"]}.
"""

CURRENT_LAW_RE = re.compile(
    r"\b(most recent|recent|current law|current authority|new cases|latest|up[- ]to[- ]date|updated authority)\b",
    re.I,
)


def is_current_law_query(text: str) -> bool:
    return bool(CURRENT_LAW_RE.search(text or ""))


FRESHNESS_MAX_AGE_DAYS = 548
THIN_RESULT_THRESHOLD = 2
COURTLISTENER_KEYWORD_MIN_OVERLAP = 0.35
_MISSING = object()
SEARCH_QUALITY_STOPWORDS = {
    "about",
    "action",
    "appeal",
    "california",
    "cause",
    "civil",
    "code",
    "court",
    "legal",
    "liability",
    "rule",
    "under",
}
_LOCAL_DOCTRINE_EXPANSIONS: tuple[tuple[set[str], str], ...] = (
    (
        {
            "premises",
            "landowner",
            "landlord",
            "owner",
            "ownership",
            "possessor",
            "possession",
            "control",
            "controlled",
            "property",
        },
        "premises liability ownership possession control landowner duty dangerous condition",
    ),
    (
        {
            "summary",
            "judgment",
            "adjudication",
            "triable",
            "burden",
            "moving",
        },
        "summary judgment burden moving party triable issue Aguilar",
    ),
    (
        {
            "discovery",
            "sanction",
            "sanctions",
            "misuse",
            "evasive",
            "terminating",
        },
        "discovery sanctions misuse discovery process terminating sanctions evasive response",
    ),
    (
        {
            "negligence",
            "foreseeability",
            "rowland",
            "civil",
        },
        "Civil Code 1714 Rowland factors foreseeability duty negligence",
    ),
)
_LOCAL_LANDMARK_CITATIONS: tuple[tuple[set[str], tuple[str, ...]], ...] = (
    (
        {
            "summary",
            "judgment",
            "adjudication",
            "triable",
            "burden",
            "moving",
        },
        ("25 Cal. 4th 826",),
    ),
    (
        {
            "premises",
            "landowner",
            "landlord",
            "owner",
            "ownership",
            "possessor",
            "possession",
            "control",
            "controlled",
            "property",
        },
        ("14 Cal. 4th 1149",),
    ),
    (
        {
            "discovery",
            "sanction",
            "sanctions",
            "misuse",
            "evasive",
            "terminating",
        },
        ("174 Cal. App. 4th 967",),
    ),
)


def _year(value: str) -> str:
    text = str(value or "")
    return text[:4] if len(text) >= 4 and text[:4].isdigit() else ""


def _candidate_key(candidate: ChatAuthorityCandidate) -> str:
    proposition = _normalize_ws(candidate.proposition)
    proposition_part = f":prop:{proposition}" if proposition else ""
    cite = re.sub(r"[^a-z0-9]+", "", (candidate.citation or "").lower())
    if cite:
        return f"cite:{cite}{proposition_part}"
    if candidate.id:
        return f"id:{candidate.id}{proposition_part}"
    name = re.sub(r"[^a-z0-9]+", "", (candidate.case_name or "").lower())
    return f"name:{name}:{candidate.year}{proposition_part}"


def _only_unverified_firm_sources(candidate: ChatAuthorityCandidate) -> bool:
    if not candidate.sources:
        return False
    for source in candidate.sources:
        if source.kind != "firm":
            return False
        if str(source.verification or "").lower() != "unverified_firm":
            return False
    return True


def _search_quality_terms(proposition: str) -> set[str]:
    return {
        term
        for term in re.findall(r"[a-z]{4,}", (proposition or "").lower())
        if term not in SEARCH_QUALITY_STOPWORDS
    }


def expand_local_legal_queries(proposition: str) -> list[str]:
    """Return deterministic local-corpus query variants for common CA doctrines."""
    base = _normalize_ws(proposition)
    if not base:
        return []
    lower = base.lower()
    terms = set(re.findall(r"[a-z]+", lower))
    queries = [base]
    for triggers, expansion in _LOCAL_DOCTRINE_EXPANSIONS:
        if terms & triggers and expansion.lower() not in lower:
            queries.append(expansion)
    return list(dict.fromkeys(queries))


def _expanded_local_search_query(proposition: str) -> str:
    queries = expand_local_legal_queries(proposition)
    if len(queries) <= 1:
        return queries[0] if queries else proposition
    return " ".join(queries)


def _local_landmark_citations(proposition: str) -> list[str]:
    terms = set(re.findall(r"[a-z]+", (proposition or "").lower()))
    citations: list[str] = []
    for triggers, mapped_citations in _LOCAL_LANDMARK_CITATIONS:
        if terms & triggers:
            citations.extend(mapped_citations)
    return list(dict.fromkeys(citations))


def _case_search_text(case: CaseResult) -> str:
    return " ".join(
        str(value or "")
        for value in (
            getattr(case, "name", ""),
            getattr(case, "citation", ""),
            getattr(case, "snippet", ""),
        )
    ).lower()


def _search_overlap_score(proposition: str, cases: list[CaseResult], *, top_n: int = 3) -> float:
    terms = _search_quality_terms(proposition)
    if not terms:
        return 1.0
    haystack = " ".join(_case_search_text(case) for case in cases[:top_n])
    if not haystack:
        return 0.0
    matched = sum(1 for term in terms if term in haystack)
    return matched / len(terms)


def _candidate_relevance_score(candidate: ChatAuthorityCandidate) -> float:
    terms = _search_quality_terms(candidate.proposition)
    if not terms:
        return 0.0
    snippet_text = (candidate.snippet or "").lower()
    text_excerpt = (candidate.text or "")[:8000].lower()
    name_text = (candidate.case_name or "").lower()
    snippet_matches = sum(1 for term in terms if term in snippet_text)
    text_matches = sum(1 for term in terms if term in text_excerpt)
    name_matches = sum(1 for term in terms if term in name_text)
    score = (snippet_matches * 5.0) + (text_matches * 1.0) + (name_matches * 0.25)
    if snippet_matches == 0 and text_matches == 0 and name_matches > 0:
        score -= 2.0
    try:
        citation_count = int(candidate.citation_count or 0)
    except (TypeError, ValueError):
        citation_count = 0
    if citation_count > 0:
        import math
        score += min(1.0, math.log10(citation_count + 1) / 4.0)
    latest = str(candidate.latest_citing_year or "")
    if latest.isdigit():
        score += max(0.0, min(0.25, (int(latest) - 2000) / 100.0))
    court = (candidate.court or "").lower()
    if court in {"cal.", "cal"} or "supreme" in court:
        score += 0.35
    elif "app" in court:
        score += 0.12
    return score


def _rank_candidates_for_selection(
    candidates: list[ChatAuthorityCandidate],
) -> list[ChatAuthorityCandidate]:
    return [
        candidate
        for _idx, candidate in sorted(
            enumerate(candidates),
            key=lambda item: (-_candidate_relevance_score(item[1]), item[0]),
        )
    ]


def _case_result_candidate(
    case: CaseResult,
    *,
    proposition: str,
    source_kind: str,
    source_label: str,
    text: str,
    authority_signals: dict[str, Any] | None = None,
) -> ChatAuthorityCandidate:
    cluster_id = str(getattr(case, "cluster_id", "") or "")
    year = _year(getattr(case, "date", "") or "")
    signals = authority_signals or {}
    case_url = case.url
    if source_kind == "local_corpus" and cluster_id.lower().startswith("cap:"):
        case_url = repair_broken_cap_url(case_url, case.citation)
    return ChatAuthorityCandidate(
        id=cluster_id or f"{source_kind}:{case.name}:{case.citation}",
        proposition=proposition,
        case_name=case.name,
        citation=case.citation,
        year=year,
        court=case.court,
        url=case_url,
        text=text or "",
        snippet=case.snippet or "",
        snippet_page_label=getattr(case, "snippet_page_label", "") or "",
        sources=[ChatResearchSource(kind=source_kind, label=source_label, verification="verified")],
        citation_count=signals.get("citation_count", getattr(case, "citation_count", None)),
        latest_citing_year=str(
            signals.get("latest_citing_year", getattr(case, "latest_citing_year", "") or "")
            or ""
        ),
    )


def _firm_candidate(row: dict[str, Any], *, proposition: str) -> ChatAuthorityCandidate:
    verification = str(row.get("verification") or "unverified_firm")
    source_brief = str(row.get("source_brief") or "")
    return ChatAuthorityCandidate(
        id=str(row.get("cluster_id") or row.get("citation") or row.get("case_name") or ""),
        proposition=proposition,
        case_name=str(row.get("case_name") or ""),
        citation=str(row.get("citation") or ""),
        year=str(row.get("year") or ""),
        court=str(row.get("court") or ""),
        url=str(row.get("opinion_url") or row.get("url") or ""),
        text=str(row.get("text") or ""),
        snippet=str(row.get("passage") or row.get("proposition") or ""),
        sources=[
            ChatResearchSource(
                kind="firm",
                label="Firm/sample-motion authority",
                verification=verification,
                reference=source_brief,
            )
        ],
    )


def _merge_candidates(candidates: Iterable[ChatAuthorityCandidate]) -> list[ChatAuthorityCandidate]:
    merged: dict[str, ChatAuthorityCandidate] = {}
    for candidate in candidates:
        key = _candidate_key(candidate)
        existing = merged.get(key)
        if existing is None:
            merged[key] = candidate
            continue
        seen_sources = {(s.kind, s.label, s.reference) for s in existing.sources}
        for source in candidate.sources:
            source_key = (source.kind, source.label, source.reference)
            if source_key not in seen_sources:
                existing.sources.append(source)
                seen_sources.add(source_key)
        if not existing.text and candidate.text:
            existing.text = candidate.text
        if not existing.snippet and candidate.snippet:
            existing.snippet = candidate.snippet
        if not existing.snippet_page_label and candidate.snippet_page_label:
            existing.snippet_page_label = candidate.snippet_page_label
        if not existing.url and candidate.url:
            existing.url = candidate.url
        if existing.citation_count is None and candidate.citation_count is not None:
            existing.citation_count = candidate.citation_count
        if not existing.latest_citing_year and candidate.latest_citing_year:
            existing.latest_citing_year = candidate.latest_citing_year
    return list(merged.values())


def _local_freshness_warning(local_corpus: Any) -> str:
    if local_corpus is None or not hasattr(local_corpus, "corpus_metadata"):
        return ""
    try:
        metadata = local_corpus.corpus_metadata() or {}
    except Exception:
        return ""
    source_counts = metadata.get("source_counts") or {}
    cl_count = int(source_counts.get("cl") or 0) if isinstance(source_counts, dict) else 0
    if cl_count <= 0:
        return "Local corpus has no CourtListener recent slice."
    max_date = str(metadata.get("max_decision_date") or "")[:10]
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", max_date):
        return "Local corpus has no max decision date metadata."
    from datetime import date, datetime

    age_days = (date.today() - datetime.strptime(max_date, "%Y-%m-%d").date()).days
    if age_days > FRESHNESS_MAX_AGE_DAYS:
        return f"Local corpus is stale; newest decision is {max_date}."
    return ""


def make_local_corpus() -> Any:
    try:
        from icharlotte_core.config import CASELAW_DATA_DIR
        from icharlotte_core.legal_research.local_corpus.corpus import LocalCaseCorpus
        from icharlotte_core.legal_research.local_corpus.embedder import OnnxEmbedder

        db_path = os.path.join(CASELAW_DATA_DIR, "corpus.db")
        vectors_path = os.path.join(CASELAW_DATA_DIR, "vectors.f16")
        if not (os.path.exists(db_path) and os.path.exists(vectors_path)):
            return None
        return LocalCaseCorpus(
            db_path=db_path,
            vectors_path=vectors_path,
            embedder=OnnxEmbedder(),
        )
    except Exception:
        return None


def make_firm_provider(corpus: Any, token: str) -> Any:
    try:
        from icharlotte_core.firm_briefs import factory
        from icharlotte_core.firm_briefs.provider import FirmAuthorityProvider

        index = factory.make_index()
        if index is None:
            return None
        cl_client = None
        if token and CourtListenerClient is not None:
            cl_client = _CourtListenerCitationAdapter(CourtListenerClient(token))
        return FirmAuthorityProvider(index, corpus, cl_client=cl_client)
    except Exception:
        return None


class _CourtListenerCitationAdapter:
    def __init__(self, client: Any) -> None:
        self.client = client
        self._citation_to_cluster_id: dict[str, str] = {}

    def lookup_by_citation(self, cite: str) -> dict[str, str] | None:
        citation = str(cite or "").strip()
        if not citation:
            return None
        try:
            records = self.client.lookup_citations(citation) or []
        except Exception:
            return None
        for record in records:
            cluster_id = self._extract_cluster_id(record)
            if cluster_id:
                self._remember_citation_cluster(citation, cluster_id)
                record_citation = str(
                    record.get("citation")
                    or record.get("matched_citation")
                    or citation
                )
                self._remember_citation_cluster(record_citation, cluster_id)
                return {"case_uid": cluster_id, "citation": record_citation}
        return None

    def get_opinion_text(self, cluster_id: str) -> str:
        identifier = str(cluster_id or "").strip()
        if not identifier:
            return ""
        resolved = self._citation_to_cluster_id.get(_citation_lookup_key(identifier))
        if resolved is None and _looks_like_citation(identifier):
            hit = self.lookup_by_citation(identifier)
            resolved = str((hit or {}).get("case_uid") or "")
        if not resolved:
            resolved = identifier
        try:
            return self.client.get_opinion_text(resolved) or ""
        except Exception:
            return ""

    def _remember_citation_cluster(self, citation: str, cluster_id: str) -> None:
        key = _citation_lookup_key(citation)
        if key:
            self._citation_to_cluster_id[key] = str(cluster_id)

    @staticmethod
    def _extract_cluster_id(record: Any) -> str:
        if not isinstance(record, dict):
            return ""
        for key in ("case_uid", "cluster_id"):
            value = record.get(key)
            if value:
                return str(value)
        cluster = record.get("cluster")
        if isinstance(cluster, dict):
            value = cluster.get("id") or cluster.get("cluster_id")
            if value:
                return str(value)
        elif cluster:
            value = _last_path_number(str(cluster))
            if value:
                return value
        clusters = record.get("clusters")
        if isinstance(clusters, list):
            for item in clusters:
                if isinstance(item, dict):
                    value = item.get("id") or item.get("cluster_id")
                    if value:
                        return str(value)
                elif item:
                    value = _last_path_number(str(item))
                    if value:
                        return value
        return ""


def _citation_lookup_key(citation: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(citation or "").lower())


def _looks_like_citation(value: str) -> bool:
    text = str(value or "")
    return bool(re.search(r"\b(?:cal|u\.s\.|f\.?\s?\d*d?|s\.ct)\b", text, re.I))


def _last_path_number(value: str) -> str:
    match = re.search(r"/(\d+)/?$", value)
    return match.group(1) if match else ""


class ChatLegalResearchService:
    def __init__(
        self,
        llm_callback: LLMCallback,
        *,
        local_corpus: Any = None,
        firm_provider: Any = None,
        courtlistener_client: Any = None,
        courtlistener_token: str = "",
        max_results_per_source: int = 8,
    ) -> None:
        self.llm_callback = llm_callback
        self.local_corpus = local_corpus
        self.firm_provider = firm_provider
        self.courtlistener_client = courtlistener_client
        self.courtlistener_token = courtlistener_token
        self.max_results_per_source = max_results_per_source

    @classmethod
    def from_environment(
        cls,
        *,
        llm_callback: LLMCallback,
        courtlistener_token: str | None = None,
    ) -> "ChatLegalResearchService":
        token = (
            courtlistener_token
            if courtlistener_token is not None
            else os.environ.get("COURTLISTENER_API_TOKEN", "")
        )
        token = (token or "").strip()
        local_corpus = make_local_corpus()
        firm_provider = make_firm_provider(local_corpus, token)
        courtlistener_client = (
            CourtListenerClient(token)
            if token and CourtListenerClient is not None
            else None
        )
        return cls(
            llm_callback=llm_callback,
            local_corpus=local_corpus,
            firm_provider=firm_provider,
            courtlistener_client=courtlistener_client,
            courtlistener_token=token,
        )

    def extract_propositions(self, *, user_text: str, context_text: str) -> list[str]:
        research_input = (
            f"USER REQUEST:\n{user_text or ''}\n\n"
            f"CONTEXT:\n{(context_text or '')[:50000]}"
        )
        try:
            raw = self.llm_callback(PROPOSITION_EXTRACTION_PROMPT, research_input) or ""
        except Exception:
            raw = ""
        data = _loads_json(raw)
        values = data.get("propositions") if isinstance(data, dict) else None
        if isinstance(values, list):
            propositions = [str(item).strip() for item in values if str(item).strip()]
        else:
            propositions = []
        if not propositions:
            fallback = (user_text or "").strip()
            if fallback:
                propositions = [fallback]
        return propositions[:5]

    def collect_candidates(
        self,
        *,
        propositions: list[str],
        settings: ChatResearchSettings,
        original_query: str,
        debug_callback: DebugCallback = None,
    ) -> tuple[list[ChatAuthorityCandidate], list[str], list[str]]:
        def debug(
            phase: str,
            message: str,
            *,
            level: str = "info",
            details: dict[str, Any] | None = None,
        ) -> None:
            if debug_callback:
                try:
                    debug_callback(
                        phase=phase,
                        message=message,
                        level=level,
                        details=details or {},
                    )
                except Exception:
                    pass

        settings = normalize_settings(settings)
        warnings: list[str] = []
        searches: list[str] = []
        all_candidates: list[ChatAuthorityCandidate] = []
        current_law = is_current_law_query(original_query)
        local_warning = _local_freshness_warning(self.local_corpus) if settings.local_corpus else ""

        for proposition in propositions:
            debug(
                "proposition",
                f"Researching proposition: {proposition}",
                details={"proposition": proposition},
            )
            local_count = 0
            non_live_candidates: list[ChatAuthorityCandidate] = []
            if settings.firm_authority:
                debug(
                    "source_search",
                    f"Searching firm/sample-motion authority for: {proposition}",
                    details={"source": "firm_authority", "proposition": proposition},
                )
                firm_candidates = self._collect_firm(proposition, settings=settings)
                non_live_candidates.extend(firm_candidates)
                all_candidates.extend(firm_candidates)
                searches.append(f"Firm/sample-motion authority: {proposition}")
                debug(
                    "source_result",
                    f"Firm/sample-motion authority returned {len(firm_candidates)} candidate(s)",
                    details={
                        "source": "firm_authority",
                        "proposition": proposition,
                        "candidate_count": len(firm_candidates),
                    },
                )
                if self.firm_provider is None:
                    warning = "Firm/sample-motion authority selected but the firm authority index is unavailable."
                    warnings.append(warning)
                    debug(
                        "warning",
                        warning,
                        level="warning",
                        details={"source": "firm_authority"},
                    )

            if settings.local_corpus:
                local_search_query = _expanded_local_search_query(proposition)
                local_query_variants = expand_local_legal_queries(proposition)
                local_landmarks = _local_landmark_citations(proposition)
                debug(
                    "source_search",
                    f"Searching Local California corpus for: {proposition}",
                    details={
                        "source": "local_corpus",
                        "proposition": proposition,
                        "query": local_search_query,
                        "query_variants": local_query_variants,
                        "landmark_citations": local_landmarks,
                    },
                )
                local_candidates, _local_hits = self._collect_case_client(
                    self.local_corpus,
                    proposition=local_search_query,
                    candidate_proposition=proposition,
                    source_kind="local_corpus",
                    source_label="Local California corpus",
                    # LocalCaseCorpus.search_opinions(semantic=True) already
                    # fuses metadata, BM25/FTS, parenthetical recall, and
                    # semantic ranking when embeddings are available. Calling
                    # semantic=False afterward repeats the expensive local DB
                    # search without improving the first-page candidate set.
                    semantic_modes=(True,),
                )
                landmark_candidates: list[ChatAuthorityCandidate] = []
                for citation in local_landmarks:
                    citation_candidates, citation_hits = self._collect_case_client(
                        self.local_corpus,
                        proposition=f"{proposition} {citation}",
                        candidate_proposition=proposition,
                        source_kind="local_corpus",
                        source_label="Local California corpus",
                        semantic_modes=(False,),
                    )
                    for candidate in citation_candidates:
                        if candidate.text:
                            candidate.snippet = _relevant_excerpt(
                                candidate.text,
                                proposition,
                                max_chars=650,
                            )
                    landmark_candidates.extend(citation_candidates)
                    _local_hits += citation_hits
                if landmark_candidates:
                    local_candidates = _rank_candidates_for_selection(
                        _merge_candidates([*local_candidates, *landmark_candidates])
                    )[: self.max_results_per_source]
                local_count = len(local_candidates)
                non_live_candidates.extend(local_candidates)
                all_candidates.extend(local_candidates)
                searches.append(f"Local California corpus: {proposition}")
                debug(
                    "source_result",
                    f"Local California corpus returned {local_count} candidate(s)",
                    details={
                        "source": "local_corpus",
                        "proposition": proposition,
                        "query": local_search_query,
                        "query_variants": local_query_variants,
                        "landmark_citations": local_landmarks,
                        "candidate_count": local_count,
                        "raw_hit_count": _local_hits,
                    },
                )
                if self.local_corpus is None:
                    warning = "Local California corpus selected but the local corpus is unavailable."
                    warnings.append(warning)
                    debug("warning", warning, level="warning", details={"source": "local_corpus"})
                elif local_count == 0:
                    warning = f"Local corpus returned thin results for: {proposition}"
                    warnings.append(warning)
                    debug("warning", warning, level="warning", details={"source": "local_corpus"})
                if local_warning:
                    warnings.append(local_warning)
                    debug(
                        "warning",
                        local_warning,
                        level="warning",
                        details={"source": "local_corpus"},
                    )

            non_live_count = len(_merge_candidates(non_live_candidates))
            fallback_thin = (
                settings.courtlistener_mode == CourtListenerMode.FALLBACK_CURRENT_LAW
                and non_live_count < THIN_RESULT_THRESHOLD
            )
            if fallback_thin and settings.local_corpus and local_count:
                warning = f"Local corpus returned thin results for: {proposition}"
                warnings.append(warning)
                debug("warning", warning, level="warning", details={"source": "local_corpus"})
            elif fallback_thin and (settings.firm_authority or settings.local_corpus):
                warning = f"Selected non-live sources returned thin results for: {proposition}"
                warnings.append(warning)
                debug("warning", warning, level="warning", details={"source": "non_live"})

            if self._should_call_courtlistener(
                settings=settings,
                non_live_count=non_live_count,
                local_warning=local_warning,
                current_law=current_law,
            ):
                if not self.courtlistener_token or self.courtlistener_client is None:
                    warning = "CourtListener API selected but COURTLISTENER_API_TOKEN is not set."
                    warnings.append(warning)
                    debug(
                        "warning",
                        warning,
                        level="warning",
                        details={"source": "courtlistener"},
                    )
                else:
                    debug(
                        "source_search",
                        f"Searching CourtListener API for: {proposition}",
                        details={"source": "courtlistener", "proposition": proposition},
                    )
                    cl_candidates, _cl_hits = self._collect_case_client(
                        self.courtlistener_client,
                        proposition=proposition,
                        source_kind="courtlistener",
                        source_label="CourtListener API",
                        enrich_results=False,
                        semantic_modes=(False, True),
                        semantic_fallback_threshold=THIN_RESULT_THRESHOLD,
                        semantic_fallback_min_overlap=COURTLISTENER_KEYWORD_MIN_OVERLAP,
                    )
                    all_candidates.extend(cl_candidates)
                    searches.append(f"CourtListener API: {proposition}")
                    debug(
                        "source_result",
                        f"CourtListener API returned {len(cl_candidates)} candidate(s)",
                        details={
                            "source": "courtlistener",
                            "proposition": proposition,
                            "candidate_count": len(cl_candidates),
                            "raw_hit_count": _cl_hits,
                        },
                    )

        unique_warnings = list(dict.fromkeys(warnings))
        unique_searches = list(dict.fromkeys(searches))
        merged_candidates = _rank_candidates_for_selection(_merge_candidates(all_candidates))
        debug(
            "candidate_summary",
            f"Collected {len(merged_candidates)} unique candidate(s)",
            details={
                "candidate_count": len(merged_candidates),
                "warning_count": len(unique_warnings),
                "search_count": len(unique_searches),
                "sources": list(unique_searches),
            },
        )
        return merged_candidates, unique_warnings, unique_searches

    def _collect_firm(
        self,
        proposition: str,
        *,
        settings: ChatResearchSettings,
    ) -> list[ChatAuthorityCandidate]:
        if self.firm_provider is None:
            return []
        restore_cl_client = _MISSING
        if settings.courtlistener_mode == CourtListenerMode.OFF and hasattr(
            self.firm_provider,
            "cl_client",
        ):
            try:
                restore_cl_client = self.firm_provider.cl_client
                self.firm_provider.cl_client = None
            except Exception:
                restore_cl_client = _MISSING
        try:
            rows = self.firm_provider.candidates_for(
                proposition,
                motion_type="",
                side="",
                limit=self.max_results_per_source,
            ) or []
        except Exception:
            return []
        finally:
            if restore_cl_client is not _MISSING:
                try:
                    self.firm_provider.cl_client = restore_cl_client
                except Exception:
                    pass
        return [_firm_candidate(row, proposition=proposition) for row in rows]

    def _collect_case_client(
        self,
        client: Any,
        *,
        proposition: str,
        candidate_proposition: str | None = None,
        source_kind: str,
        source_label: str,
        enrich_results: bool = True,
        semantic_modes: tuple[bool, ...] = (True, False),
        semantic_fallback_threshold: int | None = None,
        semantic_fallback_min_overlap: float | None = None,
    ) -> tuple[list[ChatAuthorityCandidate], int]:
        if client is None:
            return [], 0
        results: list[CaseResult] = []
        hit_count = 0
        seen_result_keys: set[str] = set()
        fallback_due_to_quality = False
        uses_fallback_policy = (
            semantic_fallback_threshold is not None
            or semantic_fallback_min_overlap is not None
        )
        for semantic in semantic_modes:
            try:
                batch = client.search_opinions(
                    proposition,
                    semantic=semantic,
                    max_results=self.max_results_per_source,
                    published_only=True,
                ) or []
            except Exception:
                batch = []
            hit_count += len(batch)
            if semantic and fallback_due_to_quality and batch:
                results = batch + results
            else:
                results.extend(batch)
            for case in batch:
                cluster_id = str(getattr(case, "cluster_id", "") or "")
                key = cluster_id or f"{getattr(case, 'name', '')}:{getattr(case, 'citation', '')}"
                if key:
                    seen_result_keys.add(key)
            if not uses_fallback_policy:
                continue
            if semantic:
                break
            enough_results = (
                semantic_fallback_threshold is None
                or len(seen_result_keys) >= semantic_fallback_threshold
            )
            quality_ok = (
                semantic_fallback_min_overlap is None
                or _search_overlap_score(proposition, batch) >= semantic_fallback_min_overlap
            )
            if enough_results and quality_ok:
                break
            fallback_due_to_quality = enough_results and not quality_ok

        candidates: list[ChatAuthorityCandidate] = []
        seen: set[str] = set()
        output_proposition = candidate_proposition or proposition
        for case in results:
            if len(candidates) >= self.max_results_per_source:
                break
            cluster_id = str(getattr(case, "cluster_id", "") or "")
            case_key = cluster_id or f"{getattr(case, 'name', '')}:{getattr(case, 'citation', '')}"
            if case_key in seen:
                continue
            seen.add(case_key)

            text = ""
            if enrich_results and cluster_id:
                try:
                    text = client.get_opinion_text(cluster_id) or ""
                except Exception:
                    text = ""

            signals: dict[str, Any] = {}
            if enrich_results and cluster_id and hasattr(client, "get_authority_signals"):
                try:
                    signals = client.get_authority_signals(cluster_id) or {}
                except Exception:
                    signals = {}

            candidates.append(
                _case_result_candidate(
                    case,
                    proposition=output_proposition,
                    source_kind=source_kind,
                    source_label=source_label,
                    text=text,
                    authority_signals=signals,
                )
            )
        return candidates, hit_count

    def select_authorities(
        self,
        proposition: str,
        candidates: list[ChatAuthorityCandidate],
    ) -> list[ChatSelectedAuthority]:
        if not candidates:
            return []
        by_id = {candidate.id: candidate for candidate in candidates}
        user_prompt = (
            f"PROPOSITION:\n{proposition}\n\n"
            f"CANDIDATES:\n{_format_candidates(candidates, proposition)}"
        )
        try:
            raw = self.llm_callback(RERANK_SELECT_PROMPT, user_prompt) or ""
        except Exception:
            raw = ""
        data = _loads_json(raw)
        selections = data.get("selections") if isinstance(data, dict) else None
        if not isinstance(selections, list):
            return []
        selected: list[ChatSelectedAuthority] = []
        for item in selections:
            if not isinstance(item, dict):
                continue
            candidate = by_id.get(str(item.get("id") or ""))
            if candidate is None:
                continue
            quote = str(item.get("quote") or "").strip()
            normalized_quote = _normalize_ws(quote)
            if not normalized_quote:
                continue
            text_match = normalized_quote in _normalize_ws(
                candidate.text or candidate.snippet
            )
            if not text_match:
                continue
            if _only_unverified_firm_sources(candidate):
                continue
            selected.append(
                ChatSelectedAuthority(
                    id=candidate.id,
                    proposition=proposition,
                    case_name=candidate.case_name,
                    citation=candidate.citation,
                    year=candidate.year,
                    court=candidate.court,
                    url=candidate.url,
                    reason=str(item.get("reason") or "").strip(),
                    supports=str(item.get("supports") or "").strip(),
                    quote=quote,
                    caveat=str(item.get("caveat") or "").strip(),
                    verification="verified",
                    sources=list(candidate.sources),
                )
            )
        return selected

    def research(
        self,
        *,
        user_text: str,
        context_text: str,
        settings: ChatResearchSettings,
        status_callback: StatusCallback = None,
        debug_callback: DebugCallback = None,
    ) -> ChatResearchPacket:
        def status(message: str) -> None:
            if status_callback:
                status_callback(message)

        def debug(
            phase: str,
            message: str,
            *,
            level: str = "info",
            details: dict[str, Any] | None = None,
        ) -> None:
            if debug_callback:
                try:
                    debug_callback(
                        phase=phase,
                        message=message,
                        level=level,
                        details=details or {},
                    )
                except Exception:
                    pass

        settings = normalize_settings(settings)
        query = (user_text or "").strip()
        if context_text:
            query = f"{query}\n\nContext:\n{context_text[:50000]}"
        debug(
            "settings",
            "Legal research settings resolved",
            details={
                "firm_authority": settings.firm_authority,
                "local_corpus": settings.local_corpus,
                "courtlistener_mode": settings.courtlistener_mode.value,
                "user_text_length": len(user_text or ""),
                "context_text_length": len(context_text or ""),
            },
        )
        status("Extracting legal research questions")
        debug("extract", "Extracting legal research questions")
        propositions = self.extract_propositions(
            user_text=user_text,
            context_text=context_text,
        )
        debug(
            "extract",
            f"Extracted {len(propositions)} legal research question(s)",
            details={"propositions": list(propositions)},
        )
        status("Searching selected legal research sources")
        candidates, warnings, searches = self.collect_candidates(
            propositions=propositions,
            settings=settings,
            original_query=query,
            debug_callback=debug,
        )
        debug(
            "candidate_summary",
            f"Ready to select authorities from {len(candidates)} candidate(s)",
            details={
                "candidate_count": len(candidates),
                "warnings": list(warnings),
                "searches": list(searches),
            },
        )
        selected: list[ChatSelectedAuthority] = []
        for proposition in propositions:
            prop_candidates = [
                candidate
                for candidate in candidates
                if candidate.proposition == proposition
            ]
            status(f"Selecting authorities for: {proposition}")
            debug(
                "select",
                f"Selecting authorities for: {proposition}",
                details={
                    "proposition": proposition,
                    "candidate_count": len(prop_candidates),
                },
            )
            before_count = len(selected)
            selected.extend(self.select_authorities(proposition, prop_candidates))
            debug(
                "select_result",
                f"Selected {len(selected) - before_count} authorit(ies) for: {proposition}",
                details={
                    "proposition": proposition,
                    "selected_count": len(selected) - before_count,
                },
            )
        selected = self._dedupe_selected(selected)
        if not selected:
            if warnings:
                debug(
                    "error",
                    "No verified legal authorities were found.",
                    level="error",
                    details={"warnings": list(warnings)},
                )
                raise ChatResearchError(
                    "No verified legal authorities were found. " + " ".join(warnings)
                )
            debug(
                "error",
                "No verified legal authorities were found for the selected research sources.",
                level="error",
            )
            raise ChatResearchError(
                "No verified legal authorities were found for the selected research sources."
            )
        status("Research complete.")
        debug(
            "complete",
            "Research complete.",
            details={
                "selected_count": len(selected),
                "warning_count": len(warnings),
                "search_count": len(searches),
            },
        )
        return ChatResearchPacket(
            query=query,
            settings=settings,
            propositions=propositions,
            searches=searches,
            selected_authorities=selected,
            warnings=warnings,
        )

    @staticmethod
    def _dedupe_selected(
        authorities: list[ChatSelectedAuthority],
    ) -> list[ChatSelectedAuthority]:
        seen: set[str] = set()
        out: list[ChatSelectedAuthority] = []
        for authority in authorities:
            citation_key = re.sub(r"[^a-z0-9]+", "", authority.citation.lower()) or authority.id
            proposition_key = _normalize_ws(authority.proposition)
            key = f"{citation_key}:{proposition_key}"
            if key in seen:
                continue
            seen.add(key)
            out.append(authority)
        return out

    @staticmethod
    def _should_call_courtlistener(
        *,
        settings: ChatResearchSettings,
        non_live_count: int,
        local_warning: str,
        current_law: bool,
    ) -> bool:
        if settings.courtlistener_mode == CourtListenerMode.OFF:
            return False
        if settings.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH:
            return True
        return non_live_count < THIN_RESULT_THRESHOLD or bool(local_warning) or current_law
