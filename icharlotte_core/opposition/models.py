"""Serializable data models for the opposition workflow."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


def _string_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        stripped = value.strip()
        return [stripped] if stripped else []
    if isinstance(value, (list, tuple)):
        return [str(item).strip() for item in value if str(item).strip()]
    return []


def _candidate_list(value: Any) -> list[Any]:
    if value is None:
        return []
    if isinstance(value, str):
        stripped = value.strip()
        return [stripped] if stripped else []
    if isinstance(value, (list, tuple)):
        candidates: list[Any] = []
        for item in value:
            if isinstance(item, dict):
                candidates.append(dict(item))
            elif isinstance(item, str) and item.strip():
                candidates.append(item.strip())
            elif item is not None and not isinstance(item, str):
                candidates.append(item)
        return candidates
    return []


@dataclass
class MotionMetadata:
    motion_type: str = ""
    moving_party: str = ""
    opposing_party: str = ""
    relief_requested: str = ""
    hearing_date: str = ""
    opposition_due_date: str = ""
    procedural_posture: str = ""
    principal_arguments: list[str] = field(default_factory=list)
    opposition_posture: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any] | None) -> "MotionMetadata":
        data = data or {}
        return cls(
            motion_type=data.get("motion_type", ""),
            moving_party=data.get("moving_party", ""),
            opposing_party=data.get("opposing_party", ""),
            relief_requested=data.get("relief_requested", ""),
            hearing_date=data.get("hearing_date", ""),
            opposition_due_date=data.get("opposition_due_date", ""),
            procedural_posture=data.get("procedural_posture", ""),
            principal_arguments=_string_list(data.get("principal_arguments")),
            opposition_posture=data.get("opposition_posture", ""),
        )

    def required_missing(self) -> list[str]:
        missing = []
        if not self.motion_type.strip():
            missing.append("Motion Type")
        if not self.relief_requested.strip():
            missing.append("Relief Requested")
        if not self.principal_arguments:
            missing.append("Principal Arguments")
        return missing


@dataclass
class OutlineNode:
    id: str = ""
    text: str = ""
    selected: bool = True
    children: list["OutlineNode"] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "text": self.text,
            "selected": self.selected,
            "children": [child.to_dict() for child in self.children],
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any] | None) -> "OutlineNode":
        data = data or {}
        return cls(
            id=data.get("id", ""),
            text=data.get("text", ""),
            selected=data.get("selected", True),
            children=[
                cls.from_dict(child) for child in data.get("children", []) or []
            ],
        )


@dataclass
class SectionPlanItem:
    id: str = ""
    path: list[str] = field(default_factory=list)
    text: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any] | None) -> "SectionPlanItem":
        data = data or {}
        return cls(
            id=data.get("id", ""),
            path=_string_list(data.get("path")),
            text=data.get("text", ""),
        )


@dataclass
class RetrievedAuthority:
    argument_id: str = ""          # outline node id / argument index
    argument_text: str = ""        # heading or proposition researched
    cluster_id: str = ""
    case_name: str = ""
    citation: str = ""             # from CourtListener metadata, not generated
    year: str = ""                 # decision year, for "Name (year) cite" form
    supports: str = ""             # one-sentence proposition this case supports
    passage: str = ""              # verbatim opinion quote
    opinion_url: str = ""
    citation_count: int | None = None
    latest_citing_year: str = ""

    # Provenance (firm-brief authority reuse). Defaults keep corpus-only behavior.
    source: str = "corpus"            # "firm" | "corpus"
    verification: str = "local"       # "local" | "courtlistener" | "unverified_firm"
    source_brief: str = ""            # path/label of the firm brief this came from
    alternatives: list = field(default_factory=list)  # corpus options for same point

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any] | None) -> "RetrievedAuthority":
        data = data or {}
        return cls(
            argument_id=data.get("argument_id", ""),
            argument_text=data.get("argument_text", ""),
            cluster_id=str(data.get("cluster_id", "") or ""),
            case_name=data.get("case_name", ""),
            citation=data.get("citation", ""),
            year=data.get("year", ""),
            supports=data.get("supports", ""),
            passage=data.get("passage", ""),
            opinion_url=data.get("opinion_url", ""),
            citation_count=data.get("citation_count"),
            latest_citing_year=data.get("latest_citing_year", ""),
        )


@dataclass
class CitationVerification:
    citation_text: str = ""
    normalized_citation: str = ""
    status: str = ""

    # Verdict from the new verifier. Empty string means "not yet verified".
    # Values: "SUPPORTED" | "PARTIAL" | "NOT_SUPPORTED" | "NOT_FOUND" | "UNVERIFIED".
    verdict: str = ""

    # Citation kind for the new pipeline: "case" | "statute" | "rule" | "unknown".
    kind: str = ""

    # The brief's surrounding proposition (1-2 sentences of context).
    proposition: str = ""

    # 1-2 verbatim sentences from the authority text cited as evidence.
    evidence: str = ""

    # Short attorney-facing explanation of the verdict.
    note: str = ""

    # Character offset of the cite in body_text (for UI underline placement).
    body_offset: int | None = None

    # Case-specific fields (existing).
    case_name: str = ""
    court: str = ""
    date: str = ""
    opinion_url: str = ""
    cluster_id: str = ""

    # Statute-specific fields.
    law_code: str = ""
    section_num: str = ""

    # Soft good-law hint (not a Shepard's/KeyCite check).
    citation_count: int | None = None
    latest_citing_year: str = ""

    # Legacy fields retained for back-compat with the old verifier.
    supporting_passage: str = ""
    support_start: int | None = None
    support_end: int | None = None
    warning: str = ""

    replacement_candidates: list[Any] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any] | None) -> "CitationVerification":
        data = data or {}
        return cls(
            citation_text=data.get("citation_text", ""),
            normalized_citation=data.get("normalized_citation", ""),
            status=data.get("status", ""),
            verdict=data.get("verdict", ""),
            kind=data.get("kind", ""),
            proposition=data.get("proposition", ""),
            evidence=data.get("evidence", ""),
            note=data.get("note", ""),
            body_offset=data.get("body_offset"),
            case_name=data.get("case_name", ""),
            court=data.get("court", ""),
            date=data.get("date", ""),
            opinion_url=data.get("opinion_url", ""),
            cluster_id=data.get("cluster_id", ""),
            law_code=data.get("law_code", ""),
            section_num=data.get("section_num", ""),
            citation_count=data.get("citation_count"),
            latest_citing_year=data.get("latest_citing_year", ""),
            supporting_passage=data.get("supporting_passage", ""),
            support_start=data.get("support_start"),
            support_end=data.get("support_end"),
            warning=data.get("warning", ""),
            replacement_candidates=_candidate_list(data.get("replacement_candidates")),
        )


@dataclass
class DraftDocument:
    title: str = ""
    body_text: str = ""
    citations: list[CitationVerification] = field(default_factory=list)
    preview_path: str = ""
    # Populated by the drafter when it rejects a response; lets the wizard
    # surface a specific reason instead of a generic "no usable body" error.
    rejection_reason: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "title": self.title,
            "body_text": self.body_text,
            "citations": [citation.to_dict() for citation in self.citations],
            "preview_path": self.preview_path,
            "rejection_reason": self.rejection_reason,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any] | None) -> "DraftDocument":
        data = data or {}
        return cls(
            title=data.get("title", ""),
            body_text=data.get("body_text", ""),
            citations=[
                citation
                if isinstance(citation, CitationVerification)
                else CitationVerification.from_dict(citation)
                for citation in data.get("citations", []) or []
            ],
            preview_path=data.get("preview_path", ""),
            rejection_reason=data.get("rejection_reason", ""),
        )
