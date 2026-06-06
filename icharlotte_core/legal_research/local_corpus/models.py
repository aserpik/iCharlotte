"""Normalized, source-agnostic records for the local case-law corpus."""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any


@dataclass
class CaseRecord:
    case_uid: str                       # source-prefixed, e.g. "cap:269732", "cl:4408734"
    source: str                         # "cap" | "cl"
    name: str = ""                      # full caption
    name_abbreviation: str = ""         # short name for display
    citation: str = ""                  # preferred CA reporter cite
    parallel_citations: list[str] = field(default_factory=list)
    court: str = ""
    decision_date: str = ""             # ISO yyyy-mm-dd
    year: str = ""
    docket_number: str = ""
    url: str = ""
    full_text: str = ""
    citation_count: int | None = None   # inbound count (good-law soft signal)
    latest_citing_year: str = ""
    cites_to: list[str] = field(default_factory=list)  # outbound reporter cites
    published_status: str = ""
    citable: bool = True

    def to_row(self) -> dict[str, Any]:
        return {
            "case_uid": self.case_uid,
            "source": self.source,
            "name": self.name,
            "name_abbreviation": self.name_abbreviation,
            "citation": self.citation,
            "parallel_citations": json.dumps(self.parallel_citations),
            "court": self.court,
            "decision_date": self.decision_date,
            "year": self.year,
            "docket_number": self.docket_number,
            "url": self.url,
            "full_text": self.full_text,
            "citation_count": self.citation_count,
            "latest_citing_year": self.latest_citing_year,
            "cites_to": json.dumps(self.cites_to),
            "published_status": self.published_status,
            "citable": 1 if self.citable else 0,
        }

    @classmethod
    def from_row(cls, row: dict[str, Any]) -> "CaseRecord":
        def _loads(v: Any) -> list[str]:
            if not v:
                return []
            try:
                out = json.loads(v)
                return [str(x) for x in out] if isinstance(out, list) else []
            except (TypeError, ValueError):
                return []
        return cls(
            case_uid=row["case_uid"],
            source=row.get("source", ""),
            name=row.get("name", ""),
            name_abbreviation=row.get("name_abbreviation", ""),
            citation=row.get("citation", ""),
            parallel_citations=_loads(row.get("parallel_citations")),
            court=row.get("court", ""),
            decision_date=row.get("decision_date", ""),
            year=row.get("year", ""),
            docket_number=row.get("docket_number", ""),
            url=row.get("url", ""),
            full_text=row.get("full_text", ""),
            citation_count=row.get("citation_count"),
            latest_citing_year=row.get("latest_citing_year", ""),
            cites_to=_loads(row.get("cites_to")),
            published_status=row.get("published_status", ""),
            citable=bool(row.get("citable", 1)),
        )


@dataclass
class PassageRecord:
    passage_uid: str          # f"{case_uid}#{ordinal}" or f"{case_uid}#parenthetical:{id}"
    case_uid: str
    ordinal: int
    text: str
    page_label: str = ""      # reporter page this passage starts on (pin-cite)
    vec_row: int | None = None  # row index into vectors.f16 (set by indexer)
    passage_type: str = "opinion"
    source: str = ""
    parenthetical_id: str = ""
    parenthetical_score: float | None = None
    described_opinion_id: str = ""
    describing_opinion_id: str = ""
    describing_cluster_id: str = ""
