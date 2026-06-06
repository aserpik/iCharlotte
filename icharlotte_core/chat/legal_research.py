"""Chat-tab legal research orchestration.

This module is intentionally Qt-free. The Chat tab owns widgets and persistence;
this module owns source settings, retrieval, quote verification, and prompt
packet formatting.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
import json
import os
import re
from typing import Any, Callable, Iterable, Optional

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
                lines.append(f"  Selected because: {authority.reason}")
                lines.append(f"  Supports: {authority.supports}")
                lines.append(f"  Source: {source_labels}")
                if authority.quote:
                    lines.append(f"  Quote: \"{authority.quote}\"")
                if authority.caveat:
                    lines.append(f"  Caveat: {authority.caveat}")
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
                lines.append(f"- {item}")
        if self.selected_authorities:
            lines.append("<i>Authorities selected:</i>")
            for authority in self.selected_authorities:
                source_labels = ", ".join(s.label for s in authority.sources) or "unknown source"
                line = f"- <b>{authority.formatted_citation}</b> [{source_labels}]"
                if authority.url:
                    line += f' <a href="{authority.url}">View</a>'
                lines.append(line)
                if authority.reason:
                    lines.append(f"  Why: {authority.reason}")
                if authority.quote:
                    lines.append(f"  Quote: &quot;{authority.quote}&quot;")
        if self.warnings:
            lines.append("<i>Warnings:</i>")
            for warning in self.warnings:
                lines.append(f"- {warning}")
        return lines


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
