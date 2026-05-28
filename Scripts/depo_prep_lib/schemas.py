"""Dataclasses + JSON validators for Depo Prep session artifacts."""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import List, Optional


@dataclass
class DeponentStatement:
    text: str
    location: str
    context: str = ""

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "DeponentStatement":
        return cls(text=d["text"], location=d["location"], context=d.get("context", ""))


@dataclass
class FactualAnchor:
    fact: str
    location: str
    topic_tags: List[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "FactualAnchor":
        return cls(fact=d["fact"], location=d["location"], topic_tags=list(d.get("topic_tags", [])))


@dataclass
class Inconsistency:
    claim_a: str
    claim_a_source: str
    claim_b: str
    claim_b_source: str
    topic_tags: List[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "Inconsistency":
        return cls(
            claim_a=d["claim_a"], claim_a_source=d["claim_a_source"],
            claim_b=d["claim_b"], claim_b_source=d["claim_b_source"],
            topic_tags=list(d.get("topic_tags", [])),
        )


@dataclass
class SourceDigest:
    source_id: str
    source_kind: str
    deponent_statements: List[DeponentStatement] = field(default_factory=list)
    factual_anchors: List[FactualAnchor] = field(default_factory=list)
    inconsistencies: List[Inconsistency] = field(default_factory=list)
    summary: str = ""

    def to_dict(self) -> dict:
        return {
            "source_id": self.source_id,
            "source_kind": self.source_kind,
            "deponent_statements": [s.to_dict() for s in self.deponent_statements],
            "factual_anchors": [a.to_dict() for a in self.factual_anchors],
            "inconsistencies": [i.to_dict() for i in self.inconsistencies],
            "summary": self.summary,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "SourceDigest":
        return cls(
            source_id=d["source_id"],
            source_kind=d["source_kind"],
            deponent_statements=[DeponentStatement.from_dict(s) for s in d.get("deponent_statements", [])],
            factual_anchors=[FactualAnchor.from_dict(a) for a in d.get("factual_anchors", [])],
            inconsistencies=[Inconsistency.from_dict(i) for i in d.get("inconsistencies", [])],
            summary=d.get("summary", ""),
        )


@dataclass
class Topic:
    id: str
    title: str
    strategic_note: str
    relevant_digest_refs: List[str] = field(default_factory=list)
    default_checked: bool = True
    lawyer_added: bool = False

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "Topic":
        return cls(
            id=d["id"], title=d["title"], strategic_note=d.get("strategic_note", ""),
            relevant_digest_refs=list(d.get("relevant_digest_refs", [])),
            default_checked=bool(d.get("default_checked", True)),
            lawyer_added=bool(d.get("lawyer_added", False)),
        )


@dataclass
class Question:
    n: int
    text: str
    purpose: Optional[str] = None
    source_facts: Optional[List[str]] = None
    impeachment_hook: Optional[str] = None
    objection_alts: Optional[List[str]] = None

    def to_dict(self) -> dict:
        d = {"n": self.n, "text": self.text}
        if self.purpose is not None: d["purpose"] = self.purpose
        if self.source_facts is not None: d["source_facts"] = list(self.source_facts)
        if self.impeachment_hook is not None: d["impeachment_hook"] = self.impeachment_hook
        if self.objection_alts is not None: d["objection_alts"] = list(self.objection_alts)
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "Question":
        return cls(
            n=int(d["n"]), text=d["text"],
            purpose=d.get("purpose"),
            source_facts=list(d["source_facts"]) if "source_facts" in d else None,
            impeachment_hook=d.get("impeachment_hook"),
            objection_alts=list(d["objection_alts"]) if "objection_alts" in d else None,
        )


@dataclass
class TopicQuestions:
    topic_id: str
    questions: List[Question] = field(default_factory=list)
    error: Optional[str] = None

    def to_dict(self) -> dict:
        d = {"topic_id": self.topic_id, "questions": [q.to_dict() for q in self.questions]}
        if self.error is not None: d["error"] = self.error
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "TopicQuestions":
        return cls(
            topic_id=d["topic_id"],
            questions=[Question.from_dict(q) for q in d.get("questions", [])],
            error=d.get("error"),
        )


_DIGEST_REQUIRED = ("source_id", "source_kind", "deponent_statements", "factual_anchors", "inconsistencies")


def validate_source_digest_dict(d: dict) -> None:
    """Raise ValueError if d is not a valid digest dict."""
    if not isinstance(d, dict):
        raise ValueError("source digest must be a dict")
    missing = [k for k in _DIGEST_REQUIRED if k not in d]
    if missing:
        raise ValueError(f"source digest missing keys: {missing}")
    for list_key in ("deponent_statements", "factual_anchors", "inconsistencies"):
        if not isinstance(d[list_key], list):
            raise ValueError(f"{list_key} must be a list")


def validate_topics_dict(d: dict) -> None:
    """Raise ValueError if d is not a valid topics dict."""
    if not isinstance(d, dict) or "topics" not in d:
        raise ValueError("topics payload must have 'topics' key")
    if not isinstance(d["topics"], list):
        raise ValueError("topics must be a list")
    for t in d["topics"]:
        for k in ("id", "title"):
            if k not in t:
                raise ValueError(f"topic missing '{k}': {t}")
