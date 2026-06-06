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
import re
from typing import Any, Callable, Iterable, Optional
from urllib.parse import urlparse

from icharlotte_core.legal_research.models import CaseResult

LLMCallback = Callable[[str, str], str]
StatusCallback = Optional[Callable[[str], None]]


class ChatResearchError(RuntimeError):
    """Raised when research mode cannot produce a verified research basis."""


class CourtListenerMode(str, Enum):
    OFF = "off"
    FALLBACK_CURRENT_LAW = "fallback_current_law"
    ALWAYS_SEARCH = "always_search"


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
    ) -> "ChatResearchSettings":
        if isinstance(courtlistener_mode, CourtListenerMode):
            mode = courtlistener_mode
        else:
            try:
                mode = CourtListenerMode(str(courtlistener_mode))
            except ValueError:
                mode = CourtListenerMode.FALLBACK_CURRENT_LAW
        return normalize_settings(
            cls(
                firm_authority=_bool_value(firm_authority, True),
                local_corpus=_bool_value(local_corpus, True),
                courtlistener_mode=mode,
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
Do not invent, recall, or add citations from memory.
If the selected authorities do not support a requested proposition, say that the selected sources did not provide support.
Include a concise section titled "Research Basis" explaining searches run, sources searched, why cited authorities were selected, and the quoted support.
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
                if authority.quote:
                    lines.append(f"  Quote: \"{authority.quote}\"")
                if authority.caveat:
                    lines.append(
                        f"  Untrusted selector caveat: {_sanitize_selector_text(authority.caveat)}"
                    )
                if authority.url:
                    lines.append(f"  URL: {authority.url}")
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
                safe_url = _safe_http_url(authority.url)
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
        return "\n\n".join(
            [
                base_system_prompt,
                RESEARCH_PROMPT_INSTRUCTION,
                self.format_authority_block(),
            ]
        )


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
        if candidate.text and candidate.snippet and candidate.snippet not in candidate.text:
            excerpt = f"{excerpt}\n\nSnippet:\n{candidate.snippet}"
        blocks.append(
            f"[{candidate.id}] {candidate.case_name}, {candidate.citation}\n"
            f"Sources: {source_labels}\n"
            f"Excerpt:\n{excerpt}"
        )
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
_MISSING = object()


def _year(value: str) -> str:
    text = str(value or "")
    return text[:4] if len(text) >= 4 and text[:4].isdigit() else ""


def _candidate_key(candidate: ChatAuthorityCandidate) -> str:
    cite = re.sub(r"[^a-z0-9]+", "", (candidate.citation or "").lower())
    if cite:
        return f"cite:{cite}"
    if candidate.id:
        return f"id:{candidate.id}"
    name = re.sub(r"[^a-z0-9]+", "", (candidate.case_name or "").lower())
    return f"name:{name}:{candidate.year}"


def _only_unverified_firm_sources(candidate: ChatAuthorityCandidate) -> bool:
    if not candidate.sources:
        return False
    for source in candidate.sources:
        if source.kind != "firm":
            return False
        if str(source.verification or "").lower() != "unverified_firm":
            return False
    return True


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
    return ChatAuthorityCandidate(
        id=cluster_id or f"{source_kind}:{case.name}:{case.citation}",
        proposition=proposition,
        case_name=case.name,
        citation=case.citation,
        year=year,
        court=case.court,
        url=case.url,
        text=text or "",
        snippet=case.snippet or "",
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
    ) -> tuple[list[ChatAuthorityCandidate], list[str], list[str]]:
        settings = normalize_settings(settings)
        warnings: list[str] = []
        searches: list[str] = []
        all_candidates: list[ChatAuthorityCandidate] = []
        current_law = is_current_law_query(original_query)
        local_warning = _local_freshness_warning(self.local_corpus) if settings.local_corpus else ""

        for proposition in propositions:
            local_count = 0
            non_live_candidates: list[ChatAuthorityCandidate] = []
            if settings.firm_authority:
                firm_candidates = self._collect_firm(proposition, settings=settings)
                non_live_candidates.extend(firm_candidates)
                all_candidates.extend(firm_candidates)
                searches.append(f"Firm/sample-motion authority: {proposition}")
                if self.firm_provider is None:
                    warnings.append(
                        "Firm/sample-motion authority selected but the firm authority index is unavailable."
                    )

            if settings.local_corpus:
                local_candidates, _local_hits = self._collect_case_client(
                    self.local_corpus,
                    proposition=proposition,
                    source_kind="local_corpus",
                    source_label="Local California corpus",
                )
                local_count = len(local_candidates)
                non_live_candidates.extend(local_candidates)
                all_candidates.extend(local_candidates)
                searches.append(f"Local California corpus: {proposition}")
                if self.local_corpus is None:
                    warnings.append("Local California corpus selected but the local corpus is unavailable.")
                elif local_count == 0:
                    warnings.append(f"Local corpus returned thin results for: {proposition}")
                if local_warning:
                    warnings.append(local_warning)

            non_live_count = len(_merge_candidates(non_live_candidates))
            fallback_thin = (
                settings.courtlistener_mode == CourtListenerMode.FALLBACK_CURRENT_LAW
                and non_live_count < THIN_RESULT_THRESHOLD
            )
            if fallback_thin and settings.local_corpus and local_count:
                warnings.append(f"Local corpus returned thin results for: {proposition}")
            elif fallback_thin and (settings.firm_authority or settings.local_corpus):
                warnings.append(f"Selected non-live sources returned thin results for: {proposition}")

            if self._should_call_courtlistener(
                settings=settings,
                non_live_count=non_live_count,
                local_warning=local_warning,
                current_law=current_law,
            ):
                if not self.courtlistener_token or self.courtlistener_client is None:
                    warnings.append("CourtListener API selected but COURTLISTENER_API_TOKEN is not set.")
                else:
                    cl_candidates, _cl_hits = self._collect_case_client(
                        self.courtlistener_client,
                        proposition=proposition,
                        source_kind="courtlistener",
                        source_label="CourtListener API",
                    )
                    all_candidates.extend(cl_candidates)
                    searches.append(f"CourtListener API: {proposition}")

        unique_warnings = list(dict.fromkeys(warnings))
        unique_searches = list(dict.fromkeys(searches))
        return _merge_candidates(all_candidates), unique_warnings, unique_searches

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
        source_kind: str,
        source_label: str,
    ) -> tuple[list[ChatAuthorityCandidate], int]:
        if client is None:
            return [], 0
        results: list[CaseResult] = []
        hit_count = 0
        for semantic in (True, False):
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
            results.extend(batch)

        candidates: list[ChatAuthorityCandidate] = []
        seen: set[str] = set()
        for case in results:
            cluster_id = str(getattr(case, "cluster_id", "") or "")
            case_key = cluster_id or f"{getattr(case, 'name', '')}:{getattr(case, 'citation', '')}"
            if case_key in seen:
                continue
            seen.add(case_key)

            text = ""
            if cluster_id:
                try:
                    text = client.get_opinion_text(cluster_id) or ""
                except Exception:
                    text = ""

            signals: dict[str, Any] = {}
            if cluster_id and hasattr(client, "get_authority_signals"):
                try:
                    signals = client.get_authority_signals(cluster_id) or {}
                except Exception:
                    signals = {}

            candidates.append(
                _case_result_candidate(
                    case,
                    proposition=proposition,
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
            text_match = normalized_quote in _normalize_ws(candidate.text)
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
    ) -> ChatResearchPacket:
        def status(message: str) -> None:
            if status_callback:
                status_callback(message)

        settings = normalize_settings(settings)
        query = (user_text or "").strip()
        if context_text:
            query = f"{query}\n\nContext:\n{context_text[:50000]}"
        status("Extracting legal research questions")
        propositions = self.extract_propositions(
            user_text=user_text,
            context_text=context_text,
        )
        status("Searching selected legal research sources")
        candidates, warnings, searches = self.collect_candidates(
            propositions=propositions,
            settings=settings,
            original_query=query,
        )
        selected: list[ChatSelectedAuthority] = []
        for proposition in propositions:
            prop_candidates = [
                candidate
                for candidate in candidates
                if candidate.proposition == proposition
            ]
            status(f"Selecting authorities for: {proposition}")
            selected.extend(self.select_authorities(proposition, prop_candidates))
        selected = self._dedupe_selected(selected)
        if not selected:
            if warnings:
                raise ChatResearchError(
                    "No verified legal authorities were found. " + " ".join(warnings)
                )
            raise ChatResearchError(
                "No verified legal authorities were found for the selected research sources."
            )
        status("Research complete.")
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
