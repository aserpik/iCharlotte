# Chat Legal Research Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace only the Chat tab legal research checkbox behavior with a persistent source selector and a Qt-free research service that searches the selected legal authority sources, verifies supporting quotes, and injects a research basis into Chat answers.

**Architecture:** Create `icharlotte_core/chat/legal_research.py` as the Chat-only research orchestrator. It reuses existing firm-brief authority, local California corpus, and CourtListener clients through dependency factories, while `icharlotte_core/ui/tabs.py` owns only the toolbar controls, QSettings persistence, progress messages, prompt injection, and transcript display.

**Tech Stack:** Python 3.12, PySide6 `QSettings`/`QToolButton`/`QMenu`, existing `LocalCaseCorpus`, `FirmAuthorityProvider`, `CourtListenerClient`, pytest, pytest-qt.

---

## Scope

Implement the approved design in `docs/superpowers/specs/2026-06-05-chat-legal-research-design.md`.

This plan does not modify Word assistant legal research, wizard motion drafting, corpus build scripts, or firm-brief ingestion.

The worktree currently contains unrelated dirty files. Every commit command in this plan stages only the files named in that task.

## File Structure

- Create `icharlotte_core/chat/legal_research.py`
  - Defines source settings, CourtListener mode normalization, research packet models, dependency factories, candidate collection, candidate merge, reranking, quote verification, and prompt/display formatting.
  - Has no Qt imports.
- Modify `icharlotte_core/chat/__init__.py`
  - Exports the Chat legal research models used by UI tests and future callers.
- Modify `icharlotte_core/ui/tabs.py`
  - Adds the `Sources` menu next to the existing `Legal Research` checkbox.
  - Persists source choices through `QSettings("iCharlotte", "iCharlotte")`.
  - Replaces the inline `LegalResearchEngine` branch in `send_message` with the new Chat research service.
  - Updates `finalize_response` to use `ChatResearchPacket`.
- Create `tests/test_chat/test_legal_research_service.py`
  - Direct unit tests for the Qt-free service using fake clients/providers.
- Create `tests/test_chat/test_legal_research_ui.py`
  - Qt tests for Chat tab source defaults, persistence, and source mode normalization.
- Modify existing Chat tests only if construction needs a small fixture helper; do not change unrelated tests.

## Task 1: Service Models And Settings Normalization

**Files:**
- Create: `icharlotte_core/chat/legal_research.py`
- Modify: `icharlotte_core/chat/__init__.py`
- Test: `tests/test_chat/test_legal_research_service.py`

- [ ] **Step 1: Write failing settings/model tests**

Create `tests/test_chat/test_legal_research_service.py` with these initial tests:

```python
import pytest

from icharlotte_core.chat.legal_research import (
    ChatResearchSettings,
    CourtListenerMode,
    normalize_settings,
)


def test_default_settings_are_firm_local_and_courtlistener_fallback():
    settings = ChatResearchSettings.default()

    assert settings.firm_authority is True
    assert settings.local_corpus is True
    assert settings.courtlistener_mode == CourtListenerMode.FALLBACK_CURRENT_LAW


def test_courtlistener_fallback_without_local_sources_becomes_always():
    settings = ChatResearchSettings(
        firm_authority=False,
        local_corpus=False,
        courtlistener_mode=CourtListenerMode.FALLBACK_CURRENT_LAW,
    )

    normalized = normalize_settings(settings)

    assert normalized.firm_authority is False
    assert normalized.local_corpus is False
    assert normalized.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH


def test_courtlistener_off_stays_off_without_local_sources():
    settings = ChatResearchSettings(
        firm_authority=False,
        local_corpus=False,
        courtlistener_mode=CourtListenerMode.OFF,
    )

    normalized = normalize_settings(settings)

    assert normalized.courtlistener_mode == CourtListenerMode.OFF


def test_settings_from_values_handles_qsettings_strings():
    settings = ChatResearchSettings.from_values(
        firm_authority="false",
        local_corpus="true",
        courtlistener_mode="always_search",
    )

    assert settings.firm_authority is False
    assert settings.local_corpus is True
    assert settings.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH


def test_unknown_courtlistener_mode_uses_default():
    settings = ChatResearchSettings.from_values(
        firm_authority=True,
        local_corpus=True,
        courtlistener_mode="unknown",
    )

    assert settings.courtlistener_mode == CourtListenerMode.FALLBACK_CURRENT_LAW
```

- [ ] **Step 2: Run the new tests and verify failure**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_service.py -q
```

Expected: FAIL with `ModuleNotFoundError: No module named 'icharlotte_core.chat.legal_research'`.

- [ ] **Step 3: Add service models and normalization**

Create `icharlotte_core/chat/legal_research.py` with this initial content:

```python
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
```

Modify `icharlotte_core/chat/__init__.py` to export the new types:

```python
from .legal_research import (
    ChatResearchError,
    ChatResearchPacket,
    ChatResearchSettings,
    CourtListenerMode,
)
```

Add these names to `__all__`:

```python
    'ChatResearchError',
    'ChatResearchPacket',
    'ChatResearchSettings',
    'CourtListenerMode',
```

- [ ] **Step 4: Run the settings/model tests**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_service.py -q
```

Expected: PASS for the 5 tests in the file.

- [ ] **Step 5: Commit Task 1**

Run:

```powershell
git add icharlotte_core/chat/legal_research.py icharlotte_core/chat/__init__.py tests/test_chat/test_legal_research_service.py
git commit -m "feat(chat): add legal research settings models"
```

## Task 2: Proposition Extraction And Current-Law Detection

**Files:**
- Modify: `icharlotte_core/chat/legal_research.py`
- Test: `tests/test_chat/test_legal_research_service.py`

- [ ] **Step 1: Add failing proposition extraction tests**

Append these tests to `tests/test_chat/test_legal_research_service.py`:

```python
from icharlotte_core.chat.legal_research import (
    ChatLegalResearchService,
    is_current_law_query,
)


def test_extract_propositions_from_json_response():
    calls = []

    def llm(system_prompt, user_prompt):
        calls.append((system_prompt, user_prompt))
        return '{"propositions":["landlord duty to repair stairs","comparative fault open and obvious condition"]}'

    service = ChatLegalResearchService(llm_callback=llm)

    props = service.extract_propositions(
        user_text="Can we defeat summary judgment on premises liability?",
        context_text="Plaintiff fell on stairs after repeated repair requests.",
    )

    assert props == [
        "landlord duty to repair stairs",
        "comparative fault open and obvious condition",
    ]
    assert "Plaintiff fell on stairs" in calls[0][1]


def test_extract_propositions_falls_back_to_user_text_when_llm_returns_bad_json():
    service = ChatLegalResearchService(llm_callback=lambda _system, _user: "not json")

    props = service.extract_propositions(
        user_text="What is the California rule for negligent hiring?",
        context_text="",
    )

    assert props == ["What is the California rule for negligent hiring?"]


def test_extract_propositions_limits_to_five_items_and_drops_blanks():
    service = ChatLegalResearchService(
        llm_callback=lambda _system, _user: (
            '{"propositions":["a"," ","b","c","d","e","f"]}'
        )
    )

    props = service.extract_propositions(user_text="research", context_text="")

    assert props == ["a", "b", "c", "d", "e"]


@pytest.mark.parametrize(
    "query",
    [
        "Find the most recent California cases on discovery sanctions",
        "What is the current law on arbitration unconscionability?",
        "Are there any new cases about negligent hiring?",
        "Use up to date authority on premises liability",
    ],
)
def test_current_law_query_detection_positive(query):
    assert is_current_law_query(query) is True


def test_current_law_query_detection_negative():
    assert is_current_law_query("What is the rule for negligence duty?") is False
```

- [ ] **Step 2: Run tests and verify failure**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_service.py -q
```

Expected: FAIL because `ChatLegalResearchService` and `is_current_law_query` are not defined.

- [ ] **Step 3: Add extraction prompt and service skeleton**

Append this code to `icharlotte_core/chat/legal_research.py`:

```python
PROPOSITION_EXTRACTION_PROMPT = """You are extracting focused California legal research questions for a litigation attorney.

Read the user's request and context. Return 1 to 5 propositions or legal questions that should be researched.

Rules:
- Focus on California legal doctrine, elements, defenses, standards, and procedural rules.
- Do not include party names unless needed to identify a specific case.
- Do not include drafting instructions or formatting instructions.
- Return strict JSON only: {"propositions":["landlord duty to repair stairs"]}.
"""

CURRENT_LAW_RE = re.compile(
    r"\b(most recent|recent|current law|new cases|latest|up to date|updated authority)\b",
    re.I,
)


def is_current_law_query(text: str) -> bool:
    return bool(CURRENT_LAW_RE.search(text or ""))


class ChatLegalResearchService:
    def __init__(
        self,
        *,
        llm_callback: LLMCallback,
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
```

- [ ] **Step 4: Run extraction tests**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_service.py -q
```

Expected: PASS for all tests currently in `tests/test_chat/test_legal_research_service.py`.

- [ ] **Step 5: Commit Task 2**

Run:

```powershell
git add icharlotte_core/chat/legal_research.py tests/test_chat/test_legal_research_service.py
git commit -m "feat(chat): extract legal research propositions"
```

## Task 3: Source Collection And CourtListener Modes

**Files:**
- Modify: `icharlotte_core/chat/legal_research.py`
- Test: `tests/test_chat/test_legal_research_service.py`

- [ ] **Step 1: Add fake source classes and failing source-mode tests**

Append these helpers and tests to `tests/test_chat/test_legal_research_service.py`:

```python
from icharlotte_core.legal_research.models import CaseResult


class FakeCorpusClient:
    def __init__(self, results=None, text_by_id=None, metadata=None):
        self.results = results or []
        self.text_by_id = text_by_id or {}
        self.metadata = metadata or {
            "source_counts": {"cl": 1},
            "max_decision_date": "2026-01-01",
        }
        self.calls = []

    def search_opinions(self, query, *, semantic=False, max_results=15, published_only=True):
        self.calls.append((query, semantic, max_results, published_only))
        return self.results

    def get_opinion_text(self, case_uid):
        return self.text_by_id.get(str(case_uid), "")

    def get_authority_signals(self, case_uid):
        return {"citation_count": 7, "latest_citing_year": "2025"}

    def corpus_metadata(self):
        return self.metadata


class FakeFirmProvider:
    def __init__(self, candidates=None):
        self.candidates = candidates or []
        self.calls = []

    def candidates_for(self, proposition, *, motion_type, side, limit=6):
        self.calls.append((proposition, motion_type, side, limit))
        return self.candidates


def _case(name="Duty v. Care", cite="30 Cal. 4th 43", uid="cap:1", text="The duty rule controls."):
    return CaseResult(
        name=name,
        citation=cite,
        date="2020-01-01",
        court="Cal.",
        snippet=text,
        url="https://example.test/case",
        cluster_id=uid,
    )


def _service_for_sources(*, local=None, firm=None, courtlistener=None):
    return ChatLegalResearchService(
        llm_callback=lambda _system, _user: '{"propositions":["duty rule"]}',
        local_corpus=local,
        firm_provider=firm,
        courtlistener_client=courtlistener,
        courtlistener_token="token" if courtlistener else "",
    )


def test_collect_local_only_searches_local_and_not_courtlistener():
    local = FakeCorpusClient(results=[_case()], text_by_id={"cap:1": "The duty rule controls."})
    cl = FakeCorpusClient(results=[_case(name="Live v. Case", uid="cl:1")])
    service = _service_for_sources(local=local, courtlistener=cl)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["duty rule"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=True,
            courtlistener_mode=CourtListenerMode.OFF,
        ),
        original_query="duty rule",
    )

    assert len(candidates) == 1
    assert candidates[0].case_name == "Duty v. Care"
    assert len(local.calls) == 2
    assert cl.calls == []
    assert warnings == []
    assert any("Local California corpus" in item for item in searches)


def test_collect_firm_only_uses_firm_provider():
    firm = FakeFirmProvider(
        candidates=[
            {
                "cluster_id": "cap:firm",
                "case_name": "Townsend v. Superior Court",
                "citation": "61 Cal.App.4th 1431",
                "year": "1998",
                "text": "The court required a good faith effort.",
                "source": "firm",
                "verification": "local",
                "source_brief": "sample.pdf",
                "passage": "good faith effort",
                "proposition": "meet and confer required",
            }
        ]
    )
    service = _service_for_sources(firm=firm)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["meet and confer required"],
        settings=ChatResearchSettings(
            firm_authority=True,
            local_corpus=False,
            courtlistener_mode=CourtListenerMode.OFF,
        ),
        original_query="meet and confer required",
    )

    assert len(candidates) == 1
    assert candidates[0].sources[0].kind == "firm"
    assert candidates[0].sources[0].reference == "sample.pdf"
    assert warnings == []
    assert searches == ["Firm/sample-motion authority: meet and confer required"]


def test_courtlistener_off_never_calls_live_client_even_with_thin_local_results():
    local = FakeCorpusClient(results=[])
    cl = FakeCorpusClient(results=[_case(name="Live v. Case", uid="cl:1")])
    service = _service_for_sources(local=local, courtlistener=cl)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["thin issue"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=True,
            courtlistener_mode=CourtListenerMode.OFF,
        ),
        original_query="thin issue",
    )

    assert candidates == []
    assert cl.calls == []
    assert any("Local corpus returned thin results" in warning for warning in warnings)


def test_courtlistener_fallback_calls_live_when_local_results_are_thin():
    local = FakeCorpusClient(results=[])
    cl = FakeCorpusClient(
        results=[_case(name="Live v. Case", cite="55 Cal.App.5th 10", uid="cl:1")],
        text_by_id={"cl:1": "Live authority supports the rule."},
    )
    service = _service_for_sources(local=local, courtlistener=cl)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["thin issue"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=True,
            courtlistener_mode=CourtListenerMode.FALLBACK_CURRENT_LAW,
        ),
        original_query="thin issue",
    )

    assert any(candidate.case_name == "Live v. Case" for candidate in candidates)
    assert cl.calls
    assert any("CourtListener API" in item for item in searches)


def test_courtlistener_always_search_calls_live_even_when_local_has_results():
    local = FakeCorpusClient(results=[_case()], text_by_id={"cap:1": "The duty rule controls."})
    cl = FakeCorpusClient(
        results=[_case(name="Live v. Case", cite="55 Cal.App.5th 10", uid="cl:1")],
        text_by_id={"cl:1": "Live authority supports the rule."},
    )
    service = _service_for_sources(local=local, courtlistener=cl)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["duty rule"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=True,
            courtlistener_mode=CourtListenerMode.ALWAYS_SEARCH,
        ),
        original_query="duty rule",
    )

    assert len(candidates) == 2
    assert cl.calls
    assert warnings == []


def test_courtlistener_selected_without_token_warns_and_skips_live():
    service = ChatLegalResearchService(
        llm_callback=lambda _system, _user: '{"propositions":["duty rule"]}',
        courtlistener_client=None,
        courtlistener_token="",
    )

    candidates, warnings, searches = service.collect_candidates(
        propositions=["duty rule"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=False,
            courtlistener_mode=CourtListenerMode.ALWAYS_SEARCH,
        ),
        original_query="duty rule",
    )

    assert candidates == []
    assert searches == []
    assert warnings == ["CourtListener API selected but COURTLISTENER_API_TOKEN is not set."]
```

- [ ] **Step 2: Run tests and verify failure**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_service.py -q
```

Expected: FAIL because `collect_candidates` is not defined.

- [ ] **Step 3: Add source collection helpers**

Append this code to `icharlotte_core/chat/legal_research.py`:

```python
FRESHNESS_MAX_AGE_DAYS = 548
THIN_RESULT_THRESHOLD = 2


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


def _case_result_candidate(
    case: CaseResult,
    *,
    proposition: str,
    source_kind: str,
    source_label: str,
    text: str,
) -> ChatAuthorityCandidate:
    cluster_id = str(getattr(case, "cluster_id", "") or "")
    year = _year(getattr(case, "date", "") or "")
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
        citation_count=getattr(case, "citation_count", None),
        latest_citing_year=getattr(case, "latest_citing_year", "") or "",
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
        url=str(row.get("opinion_url") or ""),
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
```

Add these methods inside `ChatLegalResearchService`:

```python
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
        local_warning = _local_freshness_warning(self.local_corpus)

        for proposition in propositions:
            local_count = 0
            if settings.firm_authority:
                firm_candidates = self._collect_firm(proposition)
                all_candidates.extend(firm_candidates)
                searches.append(f"Firm/sample-motion authority: {proposition}")
                if self.firm_provider is None:
                    warnings.append("Firm/sample-motion authority selected but the firm authority index is unavailable.")

            if settings.local_corpus:
                local_candidates = self._collect_case_client(
                    self.local_corpus,
                    proposition=proposition,
                    source_kind="local_corpus",
                    source_label="Local California corpus",
                )
                local_count = len(local_candidates)
                all_candidates.extend(local_candidates)
                searches.append(f"Local California corpus: {proposition}")
                if self.local_corpus is None:
                    warnings.append("Local California corpus selected but the local corpus is unavailable.")
                elif local_count < THIN_RESULT_THRESHOLD:
                    warnings.append(f"Local corpus returned thin results for: {proposition}")
                if local_warning:
                    warnings.append(local_warning)

            call_cl = self._should_call_courtlistener(
                settings=settings,
                local_count=local_count,
                local_warning=local_warning,
                current_law=current_law,
            )
            if call_cl:
                if not self.courtlistener_token or self.courtlistener_client is None:
                    warnings.append("CourtListener API selected but COURTLISTENER_API_TOKEN is not set.")
                else:
                    cl_candidates = self._collect_case_client(
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

    def _collect_firm(self, proposition: str) -> list[ChatAuthorityCandidate]:
        if self.firm_provider is None:
            return []
        try:
            rows = self.firm_provider.candidates_for(
                proposition,
                motion_type="",
                side="",
                limit=self.max_results_per_source,
            ) or []
        except Exception:
            return []
        return [_firm_candidate(row, proposition=proposition) for row in rows]

    def _collect_case_client(
        self,
        client: Any,
        *,
        proposition: str,
        source_kind: str,
        source_label: str,
    ) -> list[ChatAuthorityCandidate]:
        if client is None:
            return []
        results: list[CaseResult] = []
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
            results.extend(batch)
        candidates: list[ChatAuthorityCandidate] = []
        seen: set[str] = set()
        for case in results:
            cluster_id = str(getattr(case, "cluster_id", "") or "")
            if cluster_id in seen:
                continue
            seen.add(cluster_id)
            text = ""
            if cluster_id:
                try:
                    text = client.get_opinion_text(cluster_id) or ""
                except Exception:
                    text = ""
            candidates.append(
                _case_result_candidate(
                    case,
                    proposition=proposition,
                    source_kind=source_kind,
                    source_label=source_label,
                    text=text,
                )
            )
        return candidates

    @staticmethod
    def _should_call_courtlistener(
        *,
        settings: ChatResearchSettings,
        local_count: int,
        local_warning: str,
        current_law: bool,
    ) -> bool:
        if settings.courtlistener_mode == CourtListenerMode.OFF:
            return False
        if settings.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH:
            return True
        return local_count < THIN_RESULT_THRESHOLD or bool(local_warning) or current_law
```

- [ ] **Step 4: Run source-mode tests**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_service.py -q
```

Expected: PASS for all service tests.

- [ ] **Step 5: Commit Task 3**

Run:

```powershell
git add icharlotte_core/chat/legal_research.py tests/test_chat/test_legal_research_service.py
git commit -m "feat(chat): collect selected legal research sources"
```

## Task 4: Reranking, Quote Verification, Research Packet, And Prompt Injection Text

**Files:**
- Modify: `icharlotte_core/chat/legal_research.py`
- Test: `tests/test_chat/test_legal_research_service.py`

- [ ] **Step 1: Add failing rerank and packet tests**

Append these tests to `tests/test_chat/test_legal_research_service.py`:

```python
from icharlotte_core.chat.legal_research import ChatAuthorityCandidate, ChatResearchSource


def test_select_authorities_keeps_only_verbatim_quotes():
    candidates = [
        ChatAuthorityCandidate(
            id="cap:1",
            proposition="duty rule",
            case_name="Duty v. Care",
            citation="30 Cal. 4th 43",
            year="2020",
            text="The duty rule controls the negligence analysis.",
            sources=[ChatResearchSource(kind="local_corpus", label="Local California corpus")],
        ),
        ChatAuthorityCandidate(
            id="cap:2",
            proposition="duty rule",
            case_name="Bad v. Quote",
            citation="10 Cal.App.5th 1",
            year="2021",
            text="This text does not contain the selected phrase.",
            sources=[ChatResearchSource(kind="local_corpus", label="Local California corpus")],
        ),
    ]
    service = ChatLegalResearchService(
        llm_callback=lambda _system, _user: (
            '{"selections":['
            '{"id":"cap:1","reason":"Direct duty rule.","supports":"Duty controls negligence.",'
            '"quote":"The duty rule controls the negligence analysis.","caveat":""},'
            '{"id":"cap:2","reason":"Bad quote.","supports":"Bad support.",'
            '"quote":"fabricated quote","caveat":""}'
            ']}'
        )
    )

    selected = service.select_authorities("duty rule", candidates)

    assert len(selected) == 1
    assert selected[0].case_name == "Duty v. Care"
    assert selected[0].quote == "The duty rule controls the negligence analysis."
    assert selected[0].reason == "Direct duty rule."


def test_research_raises_when_selected_sources_produce_no_verified_authority():
    service = ChatLegalResearchService(
        llm_callback=lambda _system, _user: '{"propositions":["unsupported issue"]}',
        local_corpus=FakeCorpusClient(results=[]),
    )

    with pytest.raises(ChatResearchError) as exc:
        service.research(
            user_text="unsupported issue",
            context_text="",
            settings=ChatResearchSettings(
                firm_authority=False,
                local_corpus=True,
                courtlistener_mode=CourtListenerMode.OFF,
            ),
        )

    assert "No verified legal authorities were found" in str(exc.value)


def test_research_packet_contains_authority_block_and_research_basis():
    local = FakeCorpusClient(
        results=[_case(text="The duty rule controls the negligence analysis.")],
        text_by_id={"cap:1": "The duty rule controls the negligence analysis."},
    )

    def llm(system_prompt, user_prompt):
        if "extracting focused" in system_prompt:
            return '{"propositions":["duty rule"]}'
        return (
            '{"selections":[{"id":"cap:1","reason":"It states the governing duty rule.",'
            '"supports":"Duty controls negligence analysis.",'
            '"quote":"The duty rule controls the negligence analysis.","caveat":""}]}'
        )

    service = ChatLegalResearchService(llm_callback=llm, local_corpus=local)

    packet = service.research(
        user_text="research duty rule",
        context_text="",
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=True,
            courtlistener_mode=CourtListenerMode.OFF,
        ),
    )

    block = packet.format_authority_block()
    html_lines = packet.format_research_basis_html()

    assert "Selected authorities:" in block
    assert "Duty v. Care (2020) 30 Cal. 4th 43" in block
    assert "The duty rule controls the negligence analysis." in block
    assert any("Legal Research Basis" in line for line in html_lines)


def test_build_augmented_chat_prompt_requires_research_basis():
    packet = ChatResearchPacket(
        query="duty rule",
        settings=ChatResearchSettings.default(),
        propositions=["duty rule"],
        selected_authorities=[
            ChatSelectedAuthority(
                id="cap:1",
                proposition="duty rule",
                case_name="Duty v. Care",
                citation="30 Cal. 4th 43",
                year="2020",
                reason="It states the rule.",
                supports="Duty controls negligence.",
                quote="The duty rule controls the negligence analysis.",
                sources=[ChatResearchSource(kind="local_corpus", label="Local California corpus")],
            )
        ],
    )

    prompt = packet.build_augmented_system_prompt("Base prompt.")

    assert "Base prompt." in prompt
    assert "Research Basis" in prompt
    assert "cite only authorities" in prompt.lower()
    assert "Duty v. Care" in prompt
```

- [ ] **Step 2: Run tests and verify failure**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_service.py -q
```

Expected: FAIL because `select_authorities`, `research`, and `build_augmented_system_prompt` are not implemented.

- [ ] **Step 3: Add excerpt formatting, rerank, packet prompt, and research orchestration**

Append this code to `icharlotte_core/chat/legal_research.py`:

```python
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
        score = bisect.bisect_right(hits, window_end) - bisect.bisect_left(hits, window_start)
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
        blocks.append(
            f"[{candidate.id}] {candidate.case_name}, {candidate.citation}\n"
            f"Sources: {source_labels}\n"
            f"Excerpt:\n{excerpt}"
        )
    return "\n\n".join(blocks)
```

Add this method to `ChatResearchPacket`:

```python
    def build_augmented_system_prompt(self, base_system_prompt: str) -> str:
        return "\n\n".join(
            [
                base_system_prompt,
                RESEARCH_PROMPT_INSTRUCTION,
                self.format_authority_block(),
            ]
        )
```

Add these methods inside `ChatLegalResearchService`:

```python
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
            source_text = candidate.text or candidate.snippet
            if not quote or _normalize_ws(quote) not in _normalize_ws(source_text):
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
        propositions = self.extract_propositions(user_text=user_text, context_text=context_text)
        status("Searching selected legal research sources")
        candidates, warnings, searches = self.collect_candidates(
            propositions=propositions,
            settings=settings,
            original_query=query,
        )
        selected: list[ChatSelectedAuthority] = []
        for proposition in propositions:
            prop_candidates = [candidate for candidate in candidates if candidate.proposition == proposition]
            status(f"Selecting authorities for: {proposition}")
            selected.extend(self.select_authorities(proposition, prop_candidates))
        selected = self._dedupe_selected(selected)
        if not selected:
            if warnings:
                raise ChatResearchError(
                    "No verified legal authorities were found. " + " ".join(warnings)
                )
            raise ChatResearchError("No verified legal authorities were found for the selected research sources.")
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
    def _dedupe_selected(authorities: list[ChatSelectedAuthority]) -> list[ChatSelectedAuthority]:
        seen: set[str] = set()
        out: list[ChatSelectedAuthority] = []
        for authority in authorities:
            key = re.sub(r"[^a-z0-9]+", "", authority.citation.lower()) or authority.id
            if key in seen:
                continue
            seen.add(key)
            out.append(authority)
        return out
```

- [ ] **Step 4: Run rerank and packet tests**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_service.py -q
```

Expected: PASS for all service tests.

- [ ] **Step 5: Commit Task 4**

Run:

```powershell
git add icharlotte_core/chat/legal_research.py tests/test_chat/test_legal_research_service.py
git commit -m "feat(chat): build verified legal research packets"
```

## Task 5: Dependency Factories For Local Corpus, Firm Authority, And CourtListener

**Files:**
- Modify: `icharlotte_core/chat/legal_research.py`
- Test: `tests/test_chat/test_legal_research_service.py`

- [ ] **Step 1: Add failing factory tests**

Append these tests to `tests/test_chat/test_legal_research_service.py`:

```python
def test_make_service_from_environment_uses_factories(monkeypatch):
    created = {}

    class FakeLocal:
        pass

    class FakeFirm:
        pass

    class FakeCL:
        def __init__(self, token):
            created["token"] = token

    monkeypatch.setattr(
        "icharlotte_core.chat.legal_research.make_local_corpus",
        lambda: FakeLocal(),
    )
    monkeypatch.setattr(
        "icharlotte_core.chat.legal_research.make_firm_provider",
        lambda corpus, token: FakeFirm(),
    )
    monkeypatch.setattr(
        "icharlotte_core.chat.legal_research.CourtListenerClient",
        FakeCL,
    )

    service = ChatLegalResearchService.from_environment(
        llm_callback=lambda _system, _user: "{}",
        courtlistener_token="tok",
    )

    assert isinstance(service.local_corpus, FakeLocal)
    assert isinstance(service.firm_provider, FakeFirm)
    assert isinstance(service.courtlistener_client, FakeCL)
    assert created["token"] == "tok"


def test_make_service_from_environment_handles_missing_token(monkeypatch):
    monkeypatch.setattr(
        "icharlotte_core.chat.legal_research.make_local_corpus",
        lambda: None,
    )
    monkeypatch.setattr(
        "icharlotte_core.chat.legal_research.make_firm_provider",
        lambda corpus, token: None,
    )

    service = ChatLegalResearchService.from_environment(
        llm_callback=lambda _system, _user: "{}",
        courtlistener_token="",
    )

    assert service.local_corpus is None
    assert service.firm_provider is None
    assert service.courtlistener_client is None
```

- [ ] **Step 2: Run tests and verify failure**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_service.py -q
```

Expected: FAIL because `from_environment`, `make_local_corpus`, `make_firm_provider`, and `CourtListenerClient` are not exposed in the module.

- [ ] **Step 3: Add factories**

Near the top of `icharlotte_core/chat/legal_research.py`, add:

```python
try:
    from icharlotte_core.legal_research.sources.courtlistener import CourtListenerClient
except Exception:
    CourtListenerClient = None
```

Append these functions before `class ChatLegalResearchService`:

```python
def make_local_corpus() -> Any:
    try:
        from icharlotte_core.config import CASELAW_DATA_DIR
        from icharlotte_core.legal_research.local_corpus.corpus import LocalCaseCorpus
        from icharlotte_core.legal_research.local_corpus.embedder import OnnxEmbedder

        db_path = os.path.join(CASELAW_DATA_DIR, "corpus.db")
        vectors_path = os.path.join(CASELAW_DATA_DIR, "vectors.f16")
        if not (os.path.exists(db_path) and os.path.exists(vectors_path)):
            return None
        return LocalCaseCorpus(db_path=db_path, vectors_path=vectors_path, embedder=OnnxEmbedder())
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
            cl_client = CourtListenerClient(token)
        return FirmAuthorityProvider(index, corpus, cl_client=cl_client)
    except Exception:
        return None
```

Add this classmethod inside `ChatLegalResearchService`:

```python
    @classmethod
    def from_environment(
        cls,
        *,
        llm_callback: LLMCallback,
        courtlistener_token: str | None = None,
    ) -> "ChatLegalResearchService":
        token = courtlistener_token if courtlistener_token is not None else os.environ.get("COURTLISTENER_API_TOKEN", "")
        token = (token or "").strip()
        local_corpus = make_local_corpus()
        firm_provider = make_firm_provider(local_corpus, token)
        courtlistener_client = CourtListenerClient(token) if token and CourtListenerClient is not None else None
        return cls(
            llm_callback=llm_callback,
            local_corpus=local_corpus,
            firm_provider=firm_provider,
            courtlistener_client=courtlistener_client,
            courtlistener_token=token,
        )
```

- [ ] **Step 4: Run service tests**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_service.py -q
```

Expected: PASS for all service tests.

- [ ] **Step 5: Commit Task 5**

Run:

```powershell
git add icharlotte_core/chat/legal_research.py tests/test_chat/test_legal_research_service.py
git commit -m "feat(chat): add legal research source factories"
```

## Task 6: Chat Tab Source Selector UI And Persistence

**Files:**
- Modify: `icharlotte_core/ui/tabs.py`
- Test: `tests/test_chat/test_legal_research_ui.py`

- [ ] **Step 1: Write failing UI persistence tests**

Create `tests/test_chat/test_legal_research_ui.py`:

```python
import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QApplication

from icharlotte_core.chat.legal_research import CourtListenerMode


def _app():
    return QApplication.instance() or QApplication([])


def _clear_chat_research_settings():
    settings = QSettings("iCharlotte", "iCharlotte")
    for key in (
        "chat_tab/legal_research_firm_authority",
        "chat_tab/legal_research_local_corpus",
        "chat_tab/legal_research_courtlistener_mode",
    ):
        settings.remove(key)
    settings.sync()


def test_chat_research_source_defaults(qtbot):
    _app()
    _clear_chat_research_settings()
    from icharlotte_core.ui.tabs import ChatTab

    tab = ChatTab()
    qtbot.addWidget(tab)

    settings = tab._current_chat_research_settings()

    assert settings.firm_authority is True
    assert settings.local_corpus is True
    assert settings.courtlistener_mode == CourtListenerMode.FALLBACK_CURRENT_LAW
    assert tab.research_sources_btn.text() == "Sources: Firm + Local + CL Fallback"


def test_chat_research_source_choices_persist(qtbot):
    _app()
    _clear_chat_research_settings()
    from icharlotte_core.ui.tabs import ChatTab

    tab = ChatTab()
    qtbot.addWidget(tab)
    tab.firm_authority_action.setChecked(False)
    tab.local_corpus_action.setChecked(False)
    tab.courtlistener_always_action.setChecked(True)
    tab._on_research_source_changed()

    tab2 = ChatTab()
    qtbot.addWidget(tab2)
    settings = tab2._current_chat_research_settings()

    assert settings.firm_authority is False
    assert settings.local_corpus is False
    assert settings.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH
    assert tab2.courtlistener_always_action.isChecked() is True


def test_chat_research_source_menu_normalizes_fallback_when_no_local_sources(qtbot):
    _app()
    _clear_chat_research_settings()
    from icharlotte_core.ui.tabs import ChatTab

    tab = ChatTab()
    qtbot.addWidget(tab)
    tab.firm_authority_action.setChecked(False)
    tab.local_corpus_action.setChecked(False)
    tab.courtlistener_fallback_action.setChecked(True)
    tab._on_research_source_changed()

    settings = tab._current_chat_research_settings()

    assert settings.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH
    assert tab.courtlistener_always_action.isChecked() is True
```

- [ ] **Step 2: Run UI tests and verify failure**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_ui.py -q
```

Expected: FAIL because the ChatTab has no `research_sources_btn` or source action helpers.

- [ ] **Step 3: Add imports**

Modify `icharlotte_core/ui/tabs.py` imports:

Replace:

```python
from PySide6.QtGui import QTextCursor, QDragEnterEvent, QDropEvent, QAction, QPixmap, QBrush
```

with:

```python
from PySide6.QtGui import QTextCursor, QDragEnterEvent, QDropEvent, QAction, QActionGroup, QPixmap, QBrush
```

Replace:

```python
from ..chat import ChatPersistence, TokenCounter, Message, Conversation, BUILTIN_PROMPTS, TRANSCRIBE_PROMPT
```

with:

```python
from ..chat import (
    ChatPersistence,
    TokenCounter,
    Message,
    Conversation,
    BUILTIN_PROMPTS,
    TRANSCRIBE_PROMPT,
    ChatResearchSettings,
    CourtListenerMode,
)
```

- [ ] **Step 4: Add source menu next to the checkbox**

In `ChatTab.setup_ui`, after:

```python
        toolbar_layout.addWidget(self.legal_research_check)
```

insert:

```python
        self.research_sources_btn = QToolButton()
        self.research_sources_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.research_sources_btn.setToolTip("Choose legal research sources")
        self._build_research_sources_menu()
        toolbar_layout.addWidget(self.research_sources_btn)
```

- [ ] **Step 5: Add QSettings source helpers**

Add these methods to `ChatTab` before `# --- Splitter Persistence ---`:

```python
    # --- Chat Legal Research Source Persistence ---

    def _load_chat_research_settings(self):
        settings = QSettings("iCharlotte", "iCharlotte")
        return ChatResearchSettings.from_values(
            firm_authority=settings.value("chat_tab/legal_research_firm_authority", True),
            local_corpus=settings.value("chat_tab/legal_research_local_corpus", True),
            courtlistener_mode=settings.value(
                "chat_tab/legal_research_courtlistener_mode",
                CourtListenerMode.FALLBACK_CURRENT_LAW.value,
            ),
        )

    def _save_chat_research_settings(self, research_settings):
        settings = QSettings("iCharlotte", "iCharlotte")
        settings.setValue("chat_tab/legal_research_firm_authority", research_settings.firm_authority)
        settings.setValue("chat_tab/legal_research_local_corpus", research_settings.local_corpus)
        settings.setValue("chat_tab/legal_research_courtlistener_mode", research_settings.courtlistener_mode.value)

    def _build_research_sources_menu(self):
        menu = QMenu(self.research_sources_btn)
        current = self._load_chat_research_settings()

        self.firm_authority_action = QAction("Firm/sample-motion authority", menu)
        self.firm_authority_action.setCheckable(True)
        self.firm_authority_action.setChecked(current.firm_authority)

        self.local_corpus_action = QAction("Local California corpus", menu)
        self.local_corpus_action.setCheckable(True)
        self.local_corpus_action.setChecked(current.local_corpus)

        menu.addAction(self.firm_authority_action)
        menu.addAction(self.local_corpus_action)
        menu.addSeparator()

        mode_group = QActionGroup(menu)
        mode_group.setExclusive(True)
        self.courtlistener_off_action = QAction("CourtListener API: Off", menu)
        self.courtlistener_fallback_action = QAction("CourtListener API: Fallback/current-law", menu)
        self.courtlistener_always_action = QAction("CourtListener API: Always search", menu)
        for action in (
            self.courtlistener_off_action,
            self.courtlistener_fallback_action,
            self.courtlistener_always_action,
        ):
            action.setCheckable(True)
            mode_group.addAction(action)
            menu.addAction(action)

        if current.courtlistener_mode == CourtListenerMode.OFF:
            self.courtlistener_off_action.setChecked(True)
        elif current.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH:
            self.courtlistener_always_action.setChecked(True)
        else:
            self.courtlistener_fallback_action.setChecked(True)

        for action in (
            self.firm_authority_action,
            self.local_corpus_action,
            self.courtlistener_off_action,
            self.courtlistener_fallback_action,
            self.courtlistener_always_action,
        ):
            action.triggered.connect(self._on_research_source_changed)

        self.research_sources_btn.setMenu(menu)
        self._refresh_research_sources_label()

    def _current_chat_research_settings(self):
        if self.courtlistener_off_action.isChecked():
            mode = CourtListenerMode.OFF
        elif self.courtlistener_always_action.isChecked():
            mode = CourtListenerMode.ALWAYS_SEARCH
        else:
            mode = CourtListenerMode.FALLBACK_CURRENT_LAW
        return ChatResearchSettings.from_values(
            firm_authority=self.firm_authority_action.isChecked(),
            local_corpus=self.local_corpus_action.isChecked(),
            courtlistener_mode=mode.value,
        )

    def _on_research_source_changed(self):
        current = self._current_chat_research_settings()
        if current.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH:
            self.courtlistener_always_action.setChecked(True)
        elif current.courtlistener_mode == CourtListenerMode.OFF:
            self.courtlistener_off_action.setChecked(True)
        else:
            self.courtlistener_fallback_action.setChecked(True)
        self._save_chat_research_settings(current)
        self._refresh_research_sources_label()

    def _refresh_research_sources_label(self):
        current = self._current_chat_research_settings()
        parts = []
        if current.firm_authority:
            parts.append("Firm")
        if current.local_corpus:
            parts.append("Local")
        if current.courtlistener_mode == CourtListenerMode.FALLBACK_CURRENT_LAW:
            parts.append("CL Fallback")
        elif current.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH:
            parts.append("CL Always")
        else:
            parts.append("CL Off")
        self.research_sources_btn.setText("Sources: " + " + ".join(parts))
```

- [ ] **Step 6: Run UI source tests**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_ui.py -q
```

Expected: PASS for all tests in `tests/test_chat/test_legal_research_ui.py`.

- [ ] **Step 7: Commit Task 6**

Run:

```powershell
git add icharlotte_core/ui/tabs.py tests/test_chat/test_legal_research_ui.py
git commit -m "feat(chat): persist legal research source selector"
```

## Task 7: Replace ChatTab Inline Research Branch With New Service

**Files:**
- Modify: `icharlotte_core/ui/tabs.py`
- Test: `tests/test_chat/test_legal_research_ui.py`

- [ ] **Step 1: Add failing ChatTab wiring tests**

Append these tests to `tests/test_chat/test_legal_research_ui.py`:

```python
from types import SimpleNamespace


def test_run_chat_legal_research_passes_selected_settings(qtbot, monkeypatch):
    _app()
    _clear_chat_research_settings()
    from icharlotte_core.ui import tabs
    from icharlotte_core.ui.tabs import ChatTab

    captured = {}

    class FakeService:
        @classmethod
        def from_environment(cls, *, llm_callback):
            captured["has_llm_callback"] = callable(llm_callback)
            return cls()

        def research(self, *, user_text, context_text, settings, status_callback):
            captured["user_text"] = user_text
            captured["context_text"] = context_text
            captured["settings"] = settings
            status_callback("fake progress")
            return SimpleNamespace(
                selected_authorities=[],
                get_known_case_names=lambda: [],
                build_augmented_system_prompt=lambda base: base + "\nAUGMENTED",
                format_research_basis_html=lambda: ["<b>Legal Research Basis</b>"],
            )

    monkeypatch.setattr(tabs, "ChatLegalResearchService", FakeService)
    tab = ChatTab()
    qtbot.addWidget(tab)
    tab.firm_authority_action.setChecked(False)
    tab.local_corpus_action.setChecked(False)
    tab.courtlistener_always_action.setChecked(True)
    tab._on_research_source_changed()

    packet = tab._run_chat_legal_research("research this", "context text")

    assert captured["user_text"] == "research this"
    assert captured["context_text"] == "context text"
    assert captured["settings"].firm_authority is False
    assert captured["settings"].local_corpus is False
    assert captured["settings"].courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH
    assert "fake progress" in tab.chat_history.toPlainText()
    assert packet is not None


def test_run_chat_legal_research_fail_closed_restores_buttons(qtbot, monkeypatch):
    _app()
    _clear_chat_research_settings()
    from icharlotte_core.ui import tabs
    from icharlotte_core.ui.tabs import ChatTab
    from icharlotte_core.chat.legal_research import ChatResearchError

    class FakeService:
        @classmethod
        def from_environment(cls, *, llm_callback):
            return cls()

        def research(self, *, user_text, context_text, settings, status_callback):
            raise ChatResearchError("No verified legal authorities were found.")

    monkeypatch.setattr(tabs, "ChatLegalResearchService", FakeService)
    tab = ChatTab()
    qtbot.addWidget(tab)
    tab.send_btn.setEnabled(False)
    tab.stop_btn.setEnabled(True)

    packet = tab._run_chat_legal_research("research this", "")

    assert packet is None
    assert "No verified legal authorities were found." in tab.chat_history.toPlainText()
    assert tab.send_btn.isEnabled() is True
    assert tab.stop_btn.isEnabled() is False
```

- [ ] **Step 2: Run wiring tests and verify failure**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_ui.py -q
```

Expected: FAIL because `ChatLegalResearchService` is not imported in `tabs.py` and `_run_chat_legal_research` is not defined.

- [ ] **Step 3: Add imports for the service**

In `icharlotte_core/ui/tabs.py`, extend the Chat import block from Task 6 with:

```python
    ChatResearchError,
```

Add after the Chat imports:

```python
from ..chat.legal_research import ChatLegalResearchService
```

- [ ] **Step 4: Add the ChatTab service helper**

Add this method to `ChatTab` near the other Chat helper methods:

```python
    def _run_chat_legal_research(self, user_text, file_content):
        from icharlotte_core.llm import LLMHandler

        def llm_for_research(system_prompt, user_prompt):
            return LLMHandler.generate(
                provider=self.provider_combo.currentText(),
                model=self.model_combo.currentText(),
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                file_contents="",
                settings={**self.settings, 'stream': False, 'temperature': 0.2},
            )

        def status(message):
            self.chat_history.append(f"<i>  {message}</i>")
            QApplication.processEvents()

        self.chat_history.append("<i>Researching selected legal authority</i>")
        QApplication.processEvents()
        try:
            service = ChatLegalResearchService.from_environment(llm_callback=llm_for_research)
            return service.research(
                user_text=user_text,
                context_text=file_content[:100000] if file_content else "",
                settings=self._current_chat_research_settings(),
                status_callback=status,
            )
        except ChatResearchError as exc:
            self.chat_history.append(f"<font color='orange'>Legal research stopped: {exc}</font>")
            self.send_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            return None
        except Exception as exc:
            self.chat_history.append(f"<font color='orange'>Legal research error: {exc}</font>")
            self.send_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            return None
```

- [ ] **Step 5: Replace the inline research block in `send_message`**

In `ChatTab.send_message`, remove the current inline `LegalResearchEngine` block beginning with:

```python
        # Legal Research: if checked, run research before LLM call
```

and ending with:

```python
        self._pending_research = research_result
```

with:

```python
        research_packet = None
        if self.legal_research_check.isChecked():
            research_packet = self._run_chat_legal_research(user_text, file_content)
            if research_packet is None:
                return

        self._pending_research = research_packet
```

- [ ] **Step 6: Replace prompt augmentation**

In `ChatTab.send_message`, replace:

```python
        effective_system_prompt = self.system_prompt
        if research_result:
            from icharlotte_core.legal_research.prompts import build_augmented_system_prompt
            authority = research_result.format_authority_block()
            memo = research_result.memo or ""
            effective_system_prompt = build_augmented_system_prompt(
                self.system_prompt, authority, research_memo=memo
            )
```

with:

```python
        effective_system_prompt = self.system_prompt
        if research_packet:
            effective_system_prompt = research_packet.build_augmented_system_prompt(self.system_prompt)
```

- [ ] **Step 7: Run wiring tests**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_ui.py -q
```

Expected: PASS for all UI tests.

- [ ] **Step 8: Commit Task 7**

Run:

```powershell
git add icharlotte_core/ui/tabs.py tests/test_chat/test_legal_research_ui.py
git commit -m "feat(chat): run selected legal research before answers"
```

## Task 8: Finalize Response Display And Deterministic Citation Backstop

**Files:**
- Modify: `icharlotte_core/ui/tabs.py`
- Test: `tests/test_chat/test_legal_research_ui.py`

- [ ] **Step 1: Add failing finalize display test**

Append this test to `tests/test_chat/test_legal_research_ui.py`:

```python
def test_finalize_response_appends_research_basis_for_packet(qtbot):
    _app()
    _clear_chat_research_settings()
    from icharlotte_core.ui.tabs import ChatTab
    from icharlotte_core.chat.legal_research import (
        ChatResearchPacket,
        ChatSelectedAuthority,
        ChatResearchSource,
        ChatResearchSettings,
    )

    tab = ChatTab()
    qtbot.addWidget(tab)
    tab.stream_start_time = 1.0
    tab.stream_start_pos = tab.chat_history.textCursor().position()
    tab._pending_research = ChatResearchPacket(
        query="duty rule",
        settings=ChatResearchSettings.default(),
        selected_authorities=[
            ChatSelectedAuthority(
                id="cap:1",
                proposition="duty rule",
                case_name="Duty v. Care",
                citation="30 Cal. 4th 43",
                year="2020",
                reason="It states the governing duty rule.",
                supports="Duty controls negligence.",
                quote="The duty rule controls the negligence analysis.",
                sources=[ChatResearchSource(kind="local_corpus", label="Local California corpus")],
            )
        ],
    )

    tab.finalize_response("Duty is governed by Duty v. Care (2020) 30 Cal. 4th 43.")

    plain = tab.chat_history.toPlainText()
    assert "Legal Research Basis" in plain
    assert "It states the governing duty rule." in plain
    assert "The duty rule controls the negligence analysis." in plain
    assert tab._pending_research is None
```

- [ ] **Step 2: Run finalize test and verify failure**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_ui.py::test_finalize_response_appends_research_basis_for_packet -q
```

Expected: FAIL because `finalize_response` still expects the old `ResearchResult` shape.

- [ ] **Step 3: Update `finalize_response` citation backstop**

In `ChatTab.finalize_response`, replace the deterministic citation check block with:

```python
        if hasattr(self, '_pending_research') and self._pending_research:
            packet = self._pending_research
            try:
                known_names = packet.get_known_case_names()
                if known_names:
                    from icharlotte_core.legal_research.engine import LegalResearchEngine
                    text = LegalResearchEngine._deterministic_citation_check(
                        text, known_names
                    )
            except Exception as e:
                print(f"[ChatTab] Deterministic citation check failed: {e}")
```

- [ ] **Step 4: Update `finalize_response` research basis display**

In `ChatTab.finalize_response`, replace the old `Legal Sources Found` block with:

```python
        if hasattr(self, '_pending_research') and self._pending_research:
            packet = self._pending_research
            try:
                for line in packet.format_research_basis_html():
                    self.chat_history.append(line)
            except Exception as e:
                print(f"[ChatTab] Research basis display failed: {e}")
            self._pending_research = None
```

- [ ] **Step 5: Run finalize display tests**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_ui.py -q
```

Expected: PASS for all UI tests.

- [ ] **Step 6: Commit Task 8**

Run:

```powershell
git add icharlotte_core/ui/tabs.py tests/test_chat/test_legal_research_ui.py
git commit -m "feat(chat): display legal research basis"
```

## Task 9: Focused Verification And Regression Sweep

**Files:**
- Verify only; no planned edits unless a focused test failure identifies a defect in files from Tasks 1-8.

- [ ] **Step 1: Run Chat service and UI tests**

Run:

```powershell
python -m pytest tests/test_chat/test_legal_research_service.py tests/test_chat/test_legal_research_ui.py -q
```

Expected: PASS.

- [ ] **Step 2: Run existing Chat tests**

Run:

```powershell
python -m pytest tests/test_chat tests/test_chat_scroll_follow.py tests/test_chat_doc_extraction.py -q
```

Expected: PASS. If a failure is unrelated to this task, record the exact failing test and reason before deciding whether to expand scope.

- [ ] **Step 3: Run legal source regression tests**

Run:

```powershell
python -m pytest tests/test_legal_research/test_courtlistener.py tests/test_legal_research/test_courtlistener_semantic.py tests/test_legal_research/test_local_corpus tests/test_firm_briefs/test_provider.py tests/test_firm_briefs/test_research_injection.py -q
```

Expected: PASS. Avoid broad `tests/test_legal_research` unless optional dependencies for unrelated LegInfo/CA courts tests are installed.

- [ ] **Step 4: Compile changed Python files**

Run:

```powershell
python -m py_compile icharlotte_core/chat/legal_research.py icharlotte_core/chat/__init__.py icharlotte_core/ui/tabs.py
```

Expected: no output and exit code 0.

- [ ] **Step 5: Check staged cleanliness before final commit**

Run:

```powershell
git status --short
git diff -- icharlotte_core/chat/legal_research.py icharlotte_core/chat/__init__.py icharlotte_core/ui/tabs.py tests/test_chat/test_legal_research_service.py tests/test_chat/test_legal_research_ui.py
git diff --check -- icharlotte_core/chat/legal_research.py icharlotte_core/chat/__init__.py icharlotte_core/ui/tabs.py tests/test_chat/test_legal_research_service.py tests/test_chat/test_legal_research_ui.py
```

Expected: the named files contain only the feature changes; `git diff --check` is clean.

- [ ] **Step 6: Final feature commit if any fixes were made after Task 8**

Run this only if Task 9 required additional edits:

```powershell
git add icharlotte_core/chat/legal_research.py icharlotte_core/chat/__init__.py icharlotte_core/ui/tabs.py tests/test_chat/test_legal_research_service.py tests/test_chat/test_legal_research_ui.py
git commit -m "test(chat): verify legal research source selection"
```

## Self-Review

Spec coverage:

- Chat tab only: covered by Tasks 6-8; no Word assistant or wizard files are modified.
- New Chat-specific orchestrator: covered by Tasks 1-5.
- Persistent source selection: covered by Task 6.
- User-selectable source combinations: covered by Task 6 UI and Task 3 source collection tests.
- CourtListener off/fallback/always: covered by Tasks 1, 3, and 6.
- No CourtListener fallback when off: covered by Task 3.
- Verbatim quotes and rationale: covered by Task 4.
- Research Basis in answer/transcript: covered by Tasks 4 and 8.
- Fail-closed behavior: covered by Tasks 4 and 7.
- Focused verification: covered by Task 9.

Placeholder scan:

- The plan uses concrete paths, command lines, expected outcomes, and code blocks.
- Execution-time defect repair is constrained to the files named in Task 9, and unrelated failures must be recorded with the exact failing test and reason.

Type consistency:

- `CourtListenerMode` values match the persisted strings and tests.
- `ChatResearchSettings` flows from service tests to UI helpers to `_run_chat_legal_research`.
- `ChatResearchPacket` provides `build_augmented_system_prompt`, `get_known_case_names`, and `format_research_basis_html`, which are the only packet methods used by `ChatTab`.
