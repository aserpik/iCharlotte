"""
Data models for the deposition testimony extraction pipeline.
"""

import json
import os
from dataclasses import dataclass, field, asdict
from typing import List, Optional


@dataclass
class DeponentInfo:
    """Information about the deponent extracted from transcript header."""
    full_name: str = ""
    last_name: str = ""
    deposition_date: str = ""
    case_name: str = ""
    case_number: str = ""
    volume: str = ""

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> 'DeponentInfo':
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})


@dataclass
class QAExchange:
    """A single question-and-answer exchange from a deposition transcript."""
    id: int
    question: str          # Verbatim question text (cleaned of line numbers/timestamps)
    answer: str            # Verbatim answer text (cleaned)
    page_start: int        # Transcript page number where Q begins
    line_start: int        # Line number where Q begins
    page_end: int          # Transcript page number where A ends
    line_end: int          # Line number where A ends

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> 'QAExchange':
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})

    def citation_range(self) -> str:
        """Format the page:line range for citations."""
        if self.page_start == self.page_end:
            return f"{self.page_start}:{self.line_start}-{self.line_end}"
        return f"{self.page_start}:{self.line_start}-{self.page_end}:{self.line_end}"


@dataclass
class TranscriptIndex:
    """Complete parsed index of a deposition transcript."""
    source_pdf: str                      # Path to the original PDF
    deponent: DeponentInfo = field(default_factory=DeponentInfo)
    exchanges: List[QAExchange] = field(default_factory=list)
    total_transcript_pages: int = 0      # Number of transcript pages (not PDF pages)
    is_condensed: bool = False           # Whether transcript is condensed (4-per-page)
    parse_timestamp: str = ""            # When this index was created
    raw_text: str = ""                   # Full cleaned transcript text for viewer display

    def to_dict(self) -> dict:
        return {
            "source_pdf": self.source_pdf,
            "deponent": self.deponent.to_dict(),
            "exchanges": [e.to_dict() for e in self.exchanges],
            "total_transcript_pages": self.total_transcript_pages,
            "is_condensed": self.is_condensed,
            "parse_timestamp": self.parse_timestamp,
            # raw_text excluded from JSON — too large; regenerated on load
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'TranscriptIndex':
        return cls(
            source_pdf=data.get("source_pdf", ""),
            deponent=DeponentInfo.from_dict(data.get("deponent", {})),
            exchanges=[QAExchange.from_dict(e) for e in data.get("exchanges", [])],
            total_transcript_pages=data.get("total_transcript_pages", 0),
            is_condensed=data.get("is_condensed", False),
            parse_timestamp=data.get("parse_timestamp", ""),
        )

    def save(self, path: str = None):
        """Save index as JSON. Default path is alongside the source PDF."""
        if path is None:
            base = os.path.splitext(self.source_pdf)[0]
            path = base + ".depo_index.json"
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)
        return path

    @classmethod
    def load(cls, path: str) -> 'TranscriptIndex':
        """Load index from JSON."""
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return cls.from_dict(data)

    def get_cache_path(self) -> str:
        """Get the expected cache file path for this transcript."""
        base = os.path.splitext(self.source_pdf)[0]
        return base + ".depo_index.json"

    @staticmethod
    def cache_path_for(pdf_path: str) -> str:
        """Get the expected cache file path for a given PDF."""
        base = os.path.splitext(pdf_path)[0]
        return base + ".depo_index.json"

    @staticmethod
    def is_cache_valid(pdf_path: str) -> bool:
        """Check if a cached index exists and is newer than the PDF."""
        cache_path = TranscriptIndex.cache_path_for(pdf_path)
        if not os.path.exists(cache_path):
            return False
        pdf_mtime = os.path.getmtime(pdf_path)
        cache_mtime = os.path.getmtime(cache_path)
        return cache_mtime > pdf_mtime


@dataclass
class ExtractionResult:
    """Result of a single extraction prompt."""
    prompt: str                              # The user's extraction prompt
    selected_ids: List[int] = field(default_factory=list)  # IDs of selected exchanges
    groups: List[List[int]] = field(default_factory=list)   # Consecutive groups for citations

    def group_consecutive(self, selected_ids: List[int]):
        """Group selected IDs into consecutive runs for citation blocks."""
        self.selected_ids = sorted(selected_ids)
        self.groups = []
        if not self.selected_ids:
            return

        current_group = [self.selected_ids[0]]
        for i in range(1, len(self.selected_ids)):
            if self.selected_ids[i] == self.selected_ids[i - 1] + 1:
                current_group.append(self.selected_ids[i])
            else:
                self.groups.append(current_group)
                current_group = [self.selected_ids[i]]
        self.groups.append(current_group)
