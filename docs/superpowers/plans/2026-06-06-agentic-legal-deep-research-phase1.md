# Agentic Legal Deep Research Phase 1 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add the shared deep-research contract layer: models, source policy normalization, capped parenthetical weighting, packet formatting, and citation-audit primitives.

**Architecture:** Create a new Qt-free `icharlotte_core/legal_research/deep_research/` package. Phase 1 does not replace current Chat, Word Assistant, Oppose Motion, or Generate Motion retrieval flows; it adds stable contracts and tested guardrails they can adopt incrementally. Parentheticals are modeled as treatment signals and capped as secondary ranking inputs, never as a substitute for direct opinion-text support.

**Tech Stack:** Python 3.12, dataclasses, enum, re, pytest.

---

## Scope

Implement Phase 1 from `docs/superpowers/specs/2026-06-06-agentic-legal-deep-research-design.md`.

This plan creates contracts and utility behavior only. It does not:

- call live CourtListener;
- change local corpus build scripts;
- consume the parenthetical bulk files;
- modify Chat, Word Assistant, Oppose Motion, or Generate Motion UI behavior;
- replace `LegalResearchEngine`;
- stage or commit unrelated dirty files.

The repository currently has unrelated dirty files. Every commit step stages only the files named in that task.

## File Structure

- Create `icharlotte_core/legal_research/deep_research/__init__.py`
  - Public exports for the new contract layer.
- Create `icharlotte_core/legal_research/deep_research/models.py`
  - Enums and dataclasses for requests, plans, candidates, treatment signals, selected authority, packets, runs, and source policies.
- Create `icharlotte_core/legal_research/deep_research/ranking.py`
  - Deterministic candidate scoring and parenthetical duplicate dampening.
- Create `icharlotte_core/legal_research/deep_research/packets.py`
  - Prompt-block and research-basis formatting helpers.
- Create `icharlotte_core/legal_research/deep_research/verification.py`
  - Verbatim quote checks and off-packet citation audit.
- Create `icharlotte_core/legal_research/deep_research/orchestrator.py`
  - Minimal fail-closed public API scaffold for Phase 1.
- Create `tests/test_legal_research/test_deep_research/__init__.py`
  - Test package marker.
- Create `tests/test_legal_research/test_deep_research/test_models.py`
  - Source-policy, model serialization, and treatment-signal tests.
- Create `tests/test_legal_research/test_deep_research/test_ranking.py`
  - Parenthetical cap and duplicate dampening tests.
- Create `tests/test_legal_research/test_deep_research/test_packets.py`
  - Prompt packet and research basis tests.
- Create `tests/test_legal_research/test_deep_research/test_verification.py`
  - Quote verification and off-packet citation audit tests.
- Create `tests/test_legal_research/test_deep_research/test_orchestrator.py`
  - Public API fail-closed scaffold tests.

## Task 1: Contract Models And Source Policy

**Files:**
- Create: `icharlotte_core/legal_research/deep_research/__init__.py`
- Create: `icharlotte_core/legal_research/deep_research/models.py`
- Create: `tests/test_legal_research/test_deep_research/__init__.py`
- Create: `tests/test_legal_research/test_deep_research/test_models.py`

- [ ] **Step 1: Write failing model tests**

Create `tests/test_legal_research/test_deep_research/__init__.py` as an empty file.

Create `tests/test_legal_research/test_deep_research/test_models.py`:

```python
from icharlotte_core.legal_research.deep_research import (
    AuthorityCandidate,
    CourtListenerMode,
    DeepResearchRequest,
    ParentheticalWeightPolicy,
    ResearchSurface,
    ResearchTaskType,
    SourcePolicy,
    TreatmentClassification,
    TreatmentSignal,
    normalize_source_policy,
)


def test_default_request_is_california_fail_closed():
    request = DeepResearchRequest(question="What is the summary judgment standard?")

    assert request.surface == ResearchSurface.CHAT
    assert request.task_type == ResearchTaskType.DISCRETE_QUESTION
    assert request.jurisdiction == "California"
    assert request.fail_closed is True
    assert request.max_questions == 5


def test_source_policy_defaults_to_firm_local_and_courtlistener_fallback():
    policy = SourcePolicy.default()

    assert policy.firm_authority is True
    assert policy.local_corpus is True
    assert policy.courtlistener_mode == CourtListenerMode.FALLBACK_CURRENT_LAW
    assert policy.ca_leginfo is True
    assert policy.ca_courts_recent is True


def test_source_policy_from_values_handles_strings():
    policy = SourcePolicy.from_values(
        firm_authority="false",
        local_corpus="true",
        courtlistener_mode="always_search",
        ca_leginfo="0",
        ca_courts_recent="yes",
    )

    assert policy.firm_authority is False
    assert policy.local_corpus is True
    assert policy.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH
    assert policy.ca_leginfo is False
    assert policy.ca_courts_recent is True


def test_fallback_mode_without_local_sources_becomes_always_search():
    policy = SourcePolicy(
        firm_authority=False,
        local_corpus=False,
        courtlistener_mode=CourtListenerMode.FALLBACK_CURRENT_LAW,
    )

    normalized = normalize_source_policy(policy)

    assert normalized.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH


def test_courtlistener_off_stays_off_without_local_sources():
    policy = SourcePolicy(
        firm_authority=False,
        local_corpus=False,
        courtlistener_mode=CourtListenerMode.OFF,
    )

    normalized = normalize_source_policy(policy)

    assert normalized.courtlistener_mode == CourtListenerMode.OFF


def test_parenthetical_weight_policy_defaults_to_ten_percent_cap():
    policy = ParentheticalWeightPolicy.default()

    assert policy.max_score_contribution == 0.10
    assert policy.allow_parenthetical_as_sole_support is False
    assert policy.duplicate_similarity_threshold == 0.92


def test_treatment_signal_round_trip_preserves_classification():
    signal = TreatmentSignal(
        signal_id="sig-1",
        described_citation="12 Cal.5th 100",
        citing_case_name="Later v. Case",
        citing_citation="15 Cal.5th 200",
        parenthetical_text="holding that the trial court abused its discretion",
        classification=TreatmentClassification.SUPPORTING,
        confidence=0.82,
    )

    restored = TreatmentSignal.from_dict(signal.to_dict())

    assert restored.signal_id == "sig-1"
    assert restored.classification == TreatmentClassification.SUPPORTING
    assert restored.parenthetical_text.startswith("holding")


def test_authority_candidate_exposes_parenthetical_fields():
    candidate = AuthorityCandidate(
        candidate_id="c1",
        case_name="Smith v. Jones",
        citation="12 Cal.5th 100",
        parenthetical_match_score=0.7,
        treatment_signals=[
            TreatmentSignal(
                signal_id="sig-1",
                described_citation="12 Cal.5th 100",
                parenthetical_text="explaining the rule",
            )
        ],
    )

    data = candidate.to_dict()

    assert data["parenthetical_match_score"] == 0.7
    assert data["treatment_signals"][0]["parenthetical_text"] == "explaining the rule"
```

- [ ] **Step 2: Run tests and verify failure**

Run:

```powershell
python -m pytest tests/test_legal_research/test_deep_research/test_models.py -q
```

Expected: FAIL with `ModuleNotFoundError: No module named 'icharlotte_core.legal_research.deep_research'`.

- [ ] **Step 3: Create model implementation**

Create `icharlotte_core/legal_research/deep_research/models.py`:

```python
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
```

Create `icharlotte_core/legal_research/deep_research/__init__.py`:

```python
"""Agentic legal deep research contracts and helpers."""
from .models import (
    AuthorityCandidate,
    CitationAudit,
    CitationAuditItem,
    CitationAuditStatus,
    CourtListenerMode,
    DeepResearchRequest,
    ParentheticalWeightPolicy,
    ResearchPacket,
    ResearchPlan,
    ResearchQuestion,
    ResearchRun,
    ResearchStatus,
    ResearchStep,
    ResearchSurface,
    ResearchTaskType,
    SelectedAuthority,
    SourcePolicy,
    StatutoryMaterial,
    TreatmentClassification,
    TreatmentSignal,
    normalize_source_policy,
)

__all__ = [
    "AuthorityCandidate",
    "CitationAudit",
    "CitationAuditItem",
    "CitationAuditStatus",
    "CourtListenerMode",
    "DeepResearchRequest",
    "ParentheticalWeightPolicy",
    "ResearchPacket",
    "ResearchPlan",
    "ResearchQuestion",
    "ResearchRun",
    "ResearchStatus",
    "ResearchStep",
    "ResearchSurface",
    "ResearchTaskType",
    "SelectedAuthority",
    "SourcePolicy",
    "StatutoryMaterial",
    "TreatmentClassification",
    "TreatmentSignal",
    "normalize_source_policy",
]
```

- [ ] **Step 4: Run model tests and verify pass**

Run:

```powershell
python -m pytest tests/test_legal_research/test_deep_research/test_models.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/legal_research/deep_research/__init__.py icharlotte_core/legal_research/deep_research/models.py tests/test_legal_research/test_deep_research/__init__.py tests/test_legal_research/test_deep_research/test_models.py
git commit -m "feat(legal-research): add deep research contract models"
```

## Task 2: Parenthetical Weighting And Candidate Scoring

**Files:**
- Create: `icharlotte_core/legal_research/deep_research/ranking.py`
- Test: `tests/test_legal_research/test_deep_research/test_ranking.py`

- [ ] **Step 1: Write failing ranking tests**

Create `tests/test_legal_research/test_deep_research/test_ranking.py`:

```python
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
        TreatmentSignal(signal_id="1", parenthetical_text="holding that notice was required", confidence=0.50),
        TreatmentSignal(signal_id="2", parenthetical_text="Holding that notice was required.", confidence=0.90),
        TreatmentSignal(signal_id="3", parenthetical_text="distinguishing cases involving late notice", confidence=0.80),
    ]

    dampened = dampen_duplicate_parentheticals(signals)

    assert [signal.signal_id for signal in dampened] == ["2", "3"]
```

- [ ] **Step 2: Run tests and verify failure**

Run:

```powershell
python -m pytest tests/test_legal_research/test_deep_research/test_ranking.py -q
```

Expected: FAIL with `ModuleNotFoundError: No module named 'icharlotte_core.legal_research.deep_research.ranking'`.

- [ ] **Step 3: Implement ranking helpers**

Create `icharlotte_core/legal_research/deep_research/ranking.py`:

```python
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


def dampen_duplicate_parentheticals(signals: list[TreatmentSignal]) -> list[TreatmentSignal]:
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
```

- [ ] **Step 4: Run ranking tests and verify pass**

Run:

```powershell
python -m pytest tests/test_legal_research/test_deep_research/test_ranking.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/legal_research/deep_research/ranking.py tests/test_legal_research/test_deep_research/test_ranking.py
git commit -m "feat(legal-research): cap parenthetical ranking weight"
```

## Task 3: Prompt Packet Formatting

**Files:**
- Create: `icharlotte_core/legal_research/deep_research/packets.py`
- Modify: `icharlotte_core/legal_research/deep_research/models.py`
- Modify: `icharlotte_core/legal_research/deep_research/__init__.py`
- Test: `tests/test_legal_research/test_deep_research/test_packets.py`

- [ ] **Step 1: Write failing packet tests**

Create `tests/test_legal_research/test_deep_research/test_packets.py`:

```python
from icharlotte_core.legal_research.deep_research import ResearchPacket, SelectedAuthority
from icharlotte_core.legal_research.deep_research.packets import (
    packet_to_prompt_block,
    packet_to_research_basis_markdown,
)


def _packet():
    return ResearchPacket(
        selected_authorities=[
            SelectedAuthority(
                case_name="Smith v. Jones",
                citation="12 Cal.5th 100",
                year="2024",
                court="Cal.",
                supports="Trial courts must consider proportionality.",
                verbatim_quote="The court must weigh the likely benefit against the burden.",
                selection_reason="Recent California Supreme Court case with direct opinion text.",
                parenthetical_summary="holding that discovery burden matters",
                parenthetical_source="Later v. Case (2025) 15 Cal.5th 200",
                verification_status="verified",
            ),
            SelectedAuthority(
                case_name="Unverified Firm Case",
                citation="99 Cal.App.5th 1",
                year="2023",
                source="firm",
                supports="Firm-only cite.",
                verbatim_quote="",
                verification_status="unverified_firm",
            ),
        ],
        searches_run=["local semantic: proportional discovery burden"],
        warnings=["CourtListener was not called."],
    )


def test_prompt_block_includes_verified_authority_quote():
    block = packet_to_prompt_block(_packet())

    assert "[DEEP RESEARCH AUTHORITY]" in block
    assert "Smith v. Jones (2024) 12 Cal.5th 100" in block
    assert "Quote: The court must weigh the likely benefit against the burden." in block


def test_prompt_block_labels_parentheticals_as_research_notes():
    block = packet_to_prompt_block(_packet())

    assert "Parenthetical research note:" in block
    assert "not a quote from the cited opinion" in block


def test_prompt_block_excludes_unverified_firm_authority_from_verified_section():
    block = packet_to_prompt_block(_packet())

    assert "Unverified Firm Case" not in block


def test_research_basis_mentions_searches_and_warnings():
    basis = packet_to_research_basis_markdown(_packet())

    assert "Research Basis" in basis
    assert "local semantic: proportional discovery burden" in basis
    assert "CourtListener was not called." in basis


def test_packet_instance_methods_delegate_to_formatters():
    packet = _packet()

    assert packet.to_prompt_block() == packet_to_prompt_block(packet)
    assert packet.to_research_basis_markdown() == packet_to_research_basis_markdown(packet)
```

- [ ] **Step 2: Run tests and verify failure**

Run:

```powershell
python -m pytest tests/test_legal_research/test_deep_research/test_packets.py -q
```

Expected: FAIL with `ModuleNotFoundError: No module named 'icharlotte_core.legal_research.deep_research.packets'`.

- [ ] **Step 3: Implement packet formatters**

Create `icharlotte_core/legal_research/deep_research/packets.py`:

```python
"""Prompt and display formatting for deep research packets."""
from __future__ import annotations

from .models import ResearchPacket, SelectedAuthority


def _verified_authorities(packet: ResearchPacket) -> list[SelectedAuthority]:
    return [
        authority
        for authority in packet.selected_authorities
        if authority.verification_status != "unverified_firm"
    ]


def packet_to_prompt_block(packet: ResearchPacket) -> str:
    lines = [
        "[DEEP RESEARCH AUTHORITY]",
        "Use only the verified authorities in this block for legal citations.",
        "Do not cite cases from memory or from unverified research notes.",
        "",
    ]
    authorities = _verified_authorities(packet)
    if authorities:
        lines.append("Verified case law:")
        for authority in authorities:
            lines.append(f"- {authority.formatted_citation}")
            if authority.supports:
                lines.append(f"  Supports: {authority.supports}")
            if authority.verbatim_quote:
                lines.append(f"  Quote: {authority.verbatim_quote}")
            if authority.parenthetical_summary:
                lines.append(
                    "  Parenthetical research note: "
                    f"{authority.parenthetical_summary} "
                    f"(source: {authority.parenthetical_source or 'CourtListener bulk'}; "
                    "not a quote from the cited opinion)."
                )
            if authority.limitations:
                lines.append(f"  Limitation: {authority.limitations}")
    else:
        lines.append("Verified case law: none.")
    if packet.statutory_materials:
        lines.append("")
        lines.append("Statutory material:")
        for statute in packet.statutory_materials:
            lines.append(f"- {statute.code} section {statute.section}: {statute.title}")
            if statute.text:
                lines.append(f"  Text: {statute.text}")
    if packet.warnings:
        lines.append("")
        lines.append("Warnings:")
        for warning in packet.warnings:
            lines.append(f"- {warning}")
    lines.append("")
    lines.append("[/DEEP RESEARCH AUTHORITY]")
    return "\n".join(lines)


def packet_to_research_basis_markdown(packet: ResearchPacket) -> str:
    lines = ["### Research Basis"]
    if packet.searches_run:
        lines.append("")
        lines.append("Searches run:")
        for search in packet.searches_run:
            lines.append(f"- {search}")
    verified = _verified_authorities(packet)
    if verified:
        lines.append("")
        lines.append("Authorities selected:")
        for authority in verified:
            reason = authority.selection_reason or authority.supports
            lines.append(f"- {authority.formatted_citation}: {reason}")
    if packet.warnings:
        lines.append("")
        lines.append("Warnings:")
        for warning in packet.warnings:
            lines.append(f"- {warning}")
    return "\n".join(lines)
```

Modify `ResearchPacket` in `models.py` by adding these methods inside the class:

```python
    def to_prompt_block(self) -> str:
        from .packets import packet_to_prompt_block

        return packet_to_prompt_block(self)

    def to_research_basis_markdown(self) -> str:
        from .packets import packet_to_research_basis_markdown

        return packet_to_research_basis_markdown(self)
```

Modify `icharlotte_core/legal_research/deep_research/__init__.py` to export the formatters:

```python
from .packets import packet_to_prompt_block, packet_to_research_basis_markdown
```

Add these names to `__all__`:

```python
    "packet_to_prompt_block",
    "packet_to_research_basis_markdown",
```

- [ ] **Step 4: Run packet tests and verify pass**

Run:

```powershell
python -m pytest tests/test_legal_research/test_deep_research/test_packets.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/legal_research/deep_research/models.py icharlotte_core/legal_research/deep_research/__init__.py icharlotte_core/legal_research/deep_research/packets.py tests/test_legal_research/test_deep_research/test_packets.py
git commit -m "feat(legal-research): format deep research prompt packets"
```

## Task 4: Quote Verification And Citation Audit

**Files:**
- Create: `icharlotte_core/legal_research/deep_research/verification.py`
- Modify: `icharlotte_core/legal_research/deep_research/__init__.py`
- Test: `tests/test_legal_research/test_deep_research/test_verification.py`

- [ ] **Step 1: Write failing verification tests**

Create `tests/test_legal_research/test_deep_research/test_verification.py`:

```python
from icharlotte_core.legal_research.deep_research import (
    CitationAuditStatus,
    ResearchPacket,
    SelectedAuthority,
)
from icharlotte_core.legal_research.deep_research.verification import (
    audit_citations_against_packet,
    contains_verbatim_quote,
)


def test_contains_verbatim_quote_normalizes_whitespace():
    source = "The court must weigh the likely benefit\nagainst the burden."
    quote = "must weigh the likely benefit against the burden"

    assert contains_verbatim_quote(source, quote) is True


def test_contains_verbatim_quote_rejects_changed_words():
    source = "The court must weigh the likely benefit against the burden."
    quote = "must ignore the likely benefit against the burden"

    assert contains_verbatim_quote(source, quote) is False


def test_audit_citations_passes_known_packet_case():
    packet = ResearchPacket(
        selected_authorities=[
            SelectedAuthority(
                case_name="Smith v. Jones",
                citation="12 Cal.5th 100",
                year="2024",
            )
        ]
    )
    text = "Smith v. Jones (2024) 12 Cal.5th 100 controls this issue."

    audit = audit_citations_against_packet(text, packet)

    assert len(audit.items) == 1
    assert audit.items[0].status == CitationAuditStatus.SUPPORTED


def test_audit_citations_flags_off_packet_case():
    packet = ResearchPacket(
        selected_authorities=[
            SelectedAuthority(
                case_name="Smith v. Jones",
                citation="12 Cal.5th 100",
                year="2024",
            )
        ]
    )
    text = "Fake v. Case (2022) 99 Cal.App.5th 1 also applies."

    audit = audit_citations_against_packet(text, packet)

    assert audit.has_off_packet_citations is True
    assert audit.items[0].status == CitationAuditStatus.OFF_PACKET
```

- [ ] **Step 2: Run tests and verify failure**

Run:

```powershell
python -m pytest tests/test_legal_research/test_deep_research/test_verification.py -q
```

Expected: FAIL with `ModuleNotFoundError: No module named 'icharlotte_core.legal_research.deep_research.verification'`.

- [ ] **Step 3: Implement verification helpers**

Create `icharlotte_core/legal_research/deep_research/verification.py`:

```python
"""Verification helpers for deep research output."""
from __future__ import annotations

import re

from .models import (
    CitationAudit,
    CitationAuditItem,
    CitationAuditStatus,
    ResearchPacket,
)


_CASE_CITATION_RE = re.compile(
    r"\b([A-Z][A-Za-z0-9&.,' -]+ v\. [A-Z][A-Za-z0-9&.,' -]+ "
    r"\(\d{4}\) \d+ [A-Za-z. ]+ \d+)\b"
)


def _normalize_ws(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip().lower()


def _normalize_citation(text: str) -> str:
    text = re.sub(r"[*_`]", "", text or "")
    text = re.sub(r"[^a-z0-9]+", "", text.lower())
    return text


def contains_verbatim_quote(source_text: str, quote: str) -> bool:
    source = _normalize_ws(source_text)
    target = _normalize_ws(quote)
    if not source or not target:
        return False
    return target in source


def _known_citation_keys(packet: ResearchPacket) -> set[str]:
    keys: set[str] = set()
    for authority in packet.selected_authorities:
        if authority.verification_status == "unverified_firm":
            continue
        keys.add(_normalize_citation(authority.formatted_citation))
        if authority.citation:
            keys.add(_normalize_citation(authority.citation))
    return {key for key in keys if key}


def audit_citations_against_packet(text: str, packet: ResearchPacket) -> CitationAudit:
    known = _known_citation_keys(packet)
    items: list[CitationAuditItem] = []
    for match in _CASE_CITATION_RE.finditer(text or ""):
        citation = match.group(1).strip()
        key = _normalize_citation(citation)
        reporter_only = ""
        reporter_match = re.search(r"\(\d{4}\)\s+(.+)$", citation)
        if reporter_match:
            reporter_only = _normalize_citation(reporter_match.group(1))
        if key in known or reporter_only in known:
            items.append(
                CitationAuditItem(
                    citation_text=citation,
                    status=CitationAuditStatus.SUPPORTED,
                    detail="Citation appears in the verified research packet.",
                )
            )
        else:
            items.append(
                CitationAuditItem(
                    citation_text=citation,
                    status=CitationAuditStatus.OFF_PACKET,
                    detail="Citation does not appear in the verified research packet.",
                )
            )
    return CitationAudit(items=items)
```

Modify `icharlotte_core/legal_research/deep_research/__init__.py`:

```python
from .verification import audit_citations_against_packet, contains_verbatim_quote
```

Add these names to `__all__`:

```python
    "audit_citations_against_packet",
    "contains_verbatim_quote",
```

- [ ] **Step 4: Run verification tests and verify pass**

Run:

```powershell
python -m pytest tests/test_legal_research/test_deep_research/test_verification.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/legal_research/deep_research/__init__.py icharlotte_core/legal_research/deep_research/verification.py tests/test_legal_research/test_deep_research/test_verification.py
git commit -m "feat(legal-research): audit deep research citations"
```

## Task 5: Fail-Closed Orchestrator Scaffold

**Files:**
- Create: `icharlotte_core/legal_research/deep_research/orchestrator.py`
- Modify: `icharlotte_core/legal_research/deep_research/__init__.py`
- Test: `tests/test_legal_research/test_deep_research/test_orchestrator.py`

- [ ] **Step 1: Write failing orchestrator tests**

Create `tests/test_legal_research/test_deep_research/test_orchestrator.py`:

```python
from icharlotte_core.legal_research.deep_research import (
    CourtListenerMode,
    DeepResearchRequest,
    ResearchStatus,
    SourcePolicy,
)
from icharlotte_core.legal_research.deep_research.orchestrator import run_deep_research


def test_orchestrator_scaffold_fails_closed_without_sources():
    request = DeepResearchRequest(question="What is the summary judgment standard?")

    run = run_deep_research(
        request,
        llm_callback=lambda system, user: "",
        source_registry=None,
    )

    assert run.status == ResearchStatus.FAILED
    assert "No deep-research source adapters were provided." in run.warnings
    assert run.packet.warnings == run.warnings


def test_orchestrator_reports_unusable_source_policy():
    request = DeepResearchRequest(
        question="What is the summary judgment standard?",
        source_policy=SourcePolicy(
            firm_authority=False,
            local_corpus=False,
            courtlistener_mode=CourtListenerMode.OFF,
            ca_leginfo=False,
            ca_courts_recent=False,
        ),
    )

    run = run_deep_research(
        request,
        llm_callback=lambda system, user: "",
        source_registry=None,
    )

    assert run.status == ResearchStatus.FAILED
    assert "No selected research source is usable." in run.warnings
```

- [ ] **Step 2: Run tests and verify failure**

Run:

```powershell
python -m pytest tests/test_legal_research/test_deep_research/test_orchestrator.py -q
```

Expected: FAIL with `ModuleNotFoundError: No module named 'icharlotte_core.legal_research.deep_research.orchestrator'`.

- [ ] **Step 3: Implement fail-closed scaffold**

Create `icharlotte_core/legal_research/deep_research/orchestrator.py`:

```python
"""Phase-1 public API scaffold for agentic legal deep research."""
from __future__ import annotations

from collections.abc import Callable

from .models import (
    CourtListenerMode,
    DeepResearchRequest,
    ResearchPacket,
    ResearchRun,
    ResearchStatus,
    ResearchStep,
    SourcePolicy,
    normalize_source_policy,
)

LLMCallback = Callable[[str, str], str]
StatusCallback = Callable[[str], None]


def _has_selected_source(policy: SourcePolicy) -> bool:
    return any(
        [
            policy.firm_authority,
            policy.local_corpus,
            policy.courtlistener_mode != CourtListenerMode.OFF,
            policy.ca_leginfo,
            policy.ca_courts_recent,
        ]
    )


def run_deep_research(
    request: DeepResearchRequest,
    llm_callback: LLMCallback,
    source_registry,
    status_callback: StatusCallback | None = None,
) -> ResearchRun:
    """Return a fail-closed run until retrieval adapters are implemented."""
    policy = normalize_source_policy(request.source_policy)
    if status_callback:
        status_callback("Preparing deep research request...")
    warnings: list[str] = []
    if not _has_selected_source(policy):
        warnings.append("No selected research source is usable.")
    if source_registry is None:
        warnings.append("No deep-research source adapters were provided.")
    step = ResearchStep(
        phase="initialization",
        input_summary=request.question,
        decision="Research did not run because Phase 1 has no retrieval adapters.",
        warnings=list(warnings),
    )
    packet = ResearchPacket(warnings=list(warnings))
    return ResearchRun(
        request=request,
        status=ResearchStatus.FAILED,
        steps=[step],
        warnings=warnings,
        packet=packet,
        diagnostics={"source_policy": policy},
    )
```

Modify `icharlotte_core/legal_research/deep_research/__init__.py`:

```python
from .orchestrator import run_deep_research
```

Add this name to `__all__`:

```python
    "run_deep_research",
```

- [ ] **Step 4: Run orchestrator tests and verify pass**

Run:

```powershell
python -m pytest tests/test_legal_research/test_deep_research/test_orchestrator.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/legal_research/deep_research/__init__.py icharlotte_core/legal_research/deep_research/orchestrator.py tests/test_legal_research/test_deep_research/test_orchestrator.py
git commit -m "feat(legal-research): add deep research API scaffold"
```

## Task 6: Package Verification

**Files:**
- Verify: `icharlotte_core/legal_research/deep_research/*.py`
- Verify: `tests/test_legal_research/test_deep_research/*.py`

- [ ] **Step 1: Run focused deep-research tests**

Run:

```powershell
python -m pytest tests/test_legal_research/test_deep_research -q
```

Expected: PASS.

- [ ] **Step 2: Run legacy model tests**

Run:

```powershell
python -m pytest tests/test_legal_research/test_models.py -q
```

Expected: PASS.

- [ ] **Step 3: Compile changed package**

Run:

```powershell
python -m py_compile `
  icharlotte_core/legal_research/deep_research/__init__.py `
  icharlotte_core/legal_research/deep_research/models.py `
  icharlotte_core/legal_research/deep_research/ranking.py `
  icharlotte_core/legal_research/deep_research/packets.py `
  icharlotte_core/legal_research/deep_research/verification.py `
  icharlotte_core/legal_research/deep_research/orchestrator.py
```

Expected: no output and exit code 0.

- [ ] **Step 4: Check whitespace**

Run:

```powershell
git diff --check -- `
  icharlotte_core/legal_research/deep_research `
  tests/test_legal_research/test_deep_research
```

Expected: no output and exit code 0.

- [ ] **Step 5: Commit verification-only updates if needed**

If previous tasks already committed all changed files and this task made no edits, do not create an empty commit.

If verification required edits, stage only the edited deep-research files and focused tests:

```powershell
git add icharlotte_core/legal_research/deep_research tests/test_legal_research/test_deep_research
git commit -m "test(legal-research): verify deep research phase one"
```

## Self-Review Checklist

Spec coverage:

- Shared model layer: Task 1.
- Source policy normalization: Task 1.
- Parenthetical treatment model: Task 1.
- Parenthetical cap and duplicate damping: Task 2.
- Prompt packet labels parentheticals separately from opinion text: Task 3.
- Unverified firm authority excluded from verified prompt block: Task 3.
- Quote verification and off-packet citation audit: Task 4.
- Public API scaffold with fail-closed behavior: Task 5.
- Focused verification commands: Task 6.

Intentional Phase 1 gaps for later plans:

- Real source adapters.
- Query planning.
- Iterative retrieval and refinement.
- Conflict/adverse authority pass.
- Statutory analysis pass.
- Chat, Word Assistant, Oppose Motion, and Generate Motion migration.
