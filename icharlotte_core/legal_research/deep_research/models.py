"""Shared data contracts for agentic legal deep research."""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import Enum
from typing import Any


class ResearchSurface(str, Enum):
    CHAT = "chat"
    WORD_ASSISTANT = "word_assistant"
    OPPOSE_MOTION = "oppose_motion"
    GENERATE_MOTION = "generate_motion"
    CASE_ASSESSMENT = "case_assessment"


class ResearchTaskType(str, Enum):
    DISCRETE_QUESTION = "discrete_question"
    MOTION_ARGUMENT = "motion_argument"
    BRIEF_SECTION = "brief_section"
    STATUTORY_INTERPRETATION = "statutory_interpretation"
    MIXED = "mixed"


class ResearchStatus(str, Enum):
    COMPLETE = "complete"
    PARTIAL = "partial"
    FAILED = "failed"


class CourtListenerMode(str, Enum):
    OFF = "off"
    FALLBACK_CURRENT_LAW = "fallback_current_law"
    ALWAYS_SEARCH = "always_search"


class TreatmentClassification(str, Enum):
    SUPPORTING = "supporting"
    LIMITING = "limiting"
    DISTINGUISHING = "distinguishing"
    CONTRARY = "contrary"
    BACKGROUND = "background"
    UNKNOWN = "unknown"


class CitationAuditStatus(str, Enum):
    SUPPORTED = "supported"
    OFF_PACKET = "off_packet"
    UNVERIFIED = "unverified"


def _bool_value(value: Any, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"1", "true", "yes", "on"}:
            return True
        if lowered in {"0", "false", "no", "off"}:
            return False
        return default
    if value is None:
        return default
    return bool(value)


def _enum_value(enum_type, value: Any, default):
    if isinstance(value, enum_type):
        return value
    if isinstance(value, str):
        try:
            return enum_type(value.strip().lower())
        except ValueError:
            return default
    return default


@dataclass(frozen=True)
class SourcePolicy:
    firm_authority: bool = True
    local_corpus: bool = True
    courtlistener_mode: CourtListenerMode = CourtListenerMode.FALLBACK_CURRENT_LAW
    ca_leginfo: bool = True
    ca_courts_recent: bool = True
    fail_closed: bool = True

    @classmethod
    def default(cls) -> "SourcePolicy":
        return cls()

    @classmethod
    def from_values(
        cls,
        *,
        firm_authority: Any = True,
        local_corpus: Any = True,
        courtlistener_mode: Any = CourtListenerMode.FALLBACK_CURRENT_LAW,
        ca_leginfo: Any = True,
        ca_courts_recent: Any = True,
        fail_closed: Any = True,
    ) -> "SourcePolicy":
        return cls(
            firm_authority=_bool_value(firm_authority, True),
            local_corpus=_bool_value(local_corpus, True),
            courtlistener_mode=_enum_value(
                CourtListenerMode,
                courtlistener_mode,
                CourtListenerMode.FALLBACK_CURRENT_LAW,
            ),
            ca_leginfo=_bool_value(ca_leginfo, True),
            ca_courts_recent=_bool_value(ca_courts_recent, True),
            fail_closed=_bool_value(fail_closed, True),
        )


def normalize_source_policy(policy: SourcePolicy | None) -> SourcePolicy:
    policy = policy or SourcePolicy.default()
    if (
        not policy.firm_authority
        and not policy.local_corpus
        and policy.courtlistener_mode == CourtListenerMode.FALLBACK_CURRENT_LAW
    ):
        return SourcePolicy(
            firm_authority=False,
            local_corpus=False,
            courtlistener_mode=CourtListenerMode.ALWAYS_SEARCH,
            ca_leginfo=policy.ca_leginfo,
            ca_courts_recent=policy.ca_courts_recent,
            fail_closed=policy.fail_closed,
        )
    return policy


@dataclass(frozen=True)
class ParentheticalWeightPolicy:
    max_score_contribution: float = 0.10
    allow_parenthetical_as_sole_support: bool = False
    duplicate_similarity_threshold: float = 0.92

    @classmethod
    def default(cls) -> "ParentheticalWeightPolicy":
        return cls()


@dataclass
class DeepResearchRequest:
    question: str = ""
    surface: ResearchSurface = ResearchSurface.CHAT
    task_type: ResearchTaskType = ResearchTaskType.DISCRETE_QUESTION
    jurisdiction: str = "California"
    side: str = "neutral"
    matter_context: str = ""
    source_policy: SourcePolicy = field(default_factory=SourcePolicy.default)
    freshness_policy: str = "fallback_current_law"
    max_questions: int = 5
    max_iterations: int = 2
    fail_closed: bool = True


@dataclass
class ResearchQuestion:
    question_id: str
    text: str
    material_facts: list[str] = field(default_factory=list)
    statutory_refs: list[str] = field(default_factory=list)


@dataclass
class ResearchPlan:
    questions: list[ResearchQuestion] = field(default_factory=list)
    required_source_types: list[str] = field(default_factory=list)
    strategy_notes: list[str] = field(default_factory=list)
    current_law_required: bool = False
    adverse_authority_required: bool = False


@dataclass
class ResearchStep:
    phase: str
    input_summary: str = ""
    tool_or_source: str = ""
    output_summary: str = ""
    decision: str = ""
    warnings: list[str] = field(default_factory=list)
    duration_ms: int = 0


@dataclass
class TreatmentSignal:
    signal_id: str = ""
    source: str = "courtlistener_parenthetical"
    described_case_uid: str = ""
    described_cluster_id: str = ""
    described_citation: str = ""
    citing_case_uid: str = ""
    citing_cluster_id: str = ""
    citing_case_name: str = ""
    citing_citation: str = ""
    citing_year: str = ""
    citing_court: str = ""
    parenthetical_text: str = ""
    depth: int | None = None
    classification: TreatmentClassification = TreatmentClassification.UNKNOWN
    confidence: float = 0.0
    provenance: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["classification"] = self.classification.value
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any] | None) -> "TreatmentSignal":
        data = dict(data or {})
        data["classification"] = _enum_value(
            TreatmentClassification,
            data.get("classification", TreatmentClassification.UNKNOWN),
            TreatmentClassification.UNKNOWN,
        )
        return cls(**data)


@dataclass
class AuthorityCandidate:
    candidate_id: str = ""
    source: str = ""
    case_name: str = ""
    citation: str = ""
    year: str = ""
    court: str = ""
    cluster_id: str = ""
    case_uid: str = ""
    source_url: str = ""
    snippet: str = ""
    full_text: str = ""
    full_text_available: bool = False
    proposition_match: str = ""
    retrieval_score: float = 0.0
    semantic_score: float = 0.0
    keyword_score: float = 0.0
    recency_score: float = 0.0
    authority_signal_score: float = 0.0
    source_count_score: float = 0.0
    firm_prior_score: float = 0.0
    citation_count: int | None = None
    latest_citing_year: str = ""
    negative_signal: float = 0.0
    parentheticals: list[str] = field(default_factory=list)
    parenthetical_match_score: float = 0.0
    treatment_signals: list[TreatmentSignal] = field(default_factory=list)
    provenance: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["treatment_signals"] = [signal.to_dict() for signal in self.treatment_signals]
        return data


@dataclass
class SelectedAuthority:
    candidate_id: str = ""
    case_name: str = ""
    citation: str = ""
    year: str = ""
    court: str = ""
    source: str = ""
    source_url: str = ""
    supports: str = ""
    verbatim_quote: str = ""
    quote_location: str = ""
    selection_reason: str = ""
    limitations: str = ""
    adverse_or_distinguishable: bool = False
    parenthetical_summary: str = ""
    parenthetical_source: str = ""
    verification_status: str = "verified"
    alternatives: list[dict[str, Any]] = field(default_factory=list)

    @property
    def formatted_citation(self) -> str:
        if self.year and self.citation:
            return f"{self.case_name} ({self.year}) {self.citation}"
        return " ".join(part for part in [self.case_name, self.citation] if part)


@dataclass
class StatutoryMaterial:
    code: str = ""
    section: str = ""
    title: str = ""
    text: str = ""
    url: str = ""


@dataclass
class CitationAuditItem:
    citation_text: str
    status: CitationAuditStatus
    detail: str = ""


@dataclass
class CitationAudit:
    items: list[CitationAuditItem] = field(default_factory=list)

    @property
    def has_off_packet_citations(self) -> bool:
        return any(item.status == CitationAuditStatus.OFF_PACKET for item in self.items)


@dataclass
class ResearchPacket:
    selected_authorities: list[SelectedAuthority] = field(default_factory=list)
    statutory_materials: list[StatutoryMaterial] = field(default_factory=list)
    adverse_authorities: list[SelectedAuthority] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    searches_run: list[str] = field(default_factory=list)

    def known_case_names(self) -> list[str]:
        return [authority.case_name for authority in self.selected_authorities if authority.case_name]

    def known_reporter_citations(self) -> list[str]:
        return [authority.citation for authority in self.selected_authorities if authority.citation]

    def to_prompt_block(self) -> str:
        from .packets import packet_to_prompt_block

        return packet_to_prompt_block(self)

    def to_research_basis_markdown(self) -> str:
        from .packets import packet_to_research_basis_markdown

        return packet_to_research_basis_markdown(self)


@dataclass
class ResearchRun:
    run_id: str = ""
    request: DeepResearchRequest = field(default_factory=DeepResearchRequest)
    status: ResearchStatus = ResearchStatus.PARTIAL
    plan: ResearchPlan = field(default_factory=ResearchPlan)
    questions: list[ResearchQuestion] = field(default_factory=list)
    steps: list[ResearchStep] = field(default_factory=list)
    candidates: list[AuthorityCandidate] = field(default_factory=list)
    selected_authorities: list[SelectedAuthority] = field(default_factory=list)
    statutory_materials: list[StatutoryMaterial] = field(default_factory=list)
    adverse_authorities: list[SelectedAuthority] = field(default_factory=list)
    treatment_signals: list[TreatmentSignal] = field(default_factory=list)
    citation_audit: CitationAudit = field(default_factory=CitationAudit)
    synthesis: str = ""
    warnings: list[str] = field(default_factory=list)
    packet: ResearchPacket = field(default_factory=ResearchPacket)
    diagnostics: dict[str, Any] = field(default_factory=dict)
