from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class MemberFile:
    """One source document inside a library entry."""
    source_path: str
    source_name: str
    blob: Optional[str]          # "<sha1>.txt" in blobs/, or None if extraction failed
    char_count: int = 0
    est_tokens: int = 0
    extract_method: str = ""
    error: Optional[str] = None

    def to_dict(self) -> dict:
        return {
            "source_path": self.source_path,
            "source_name": self.source_name,
            "blob": self.blob,
            "char_count": self.char_count,
            "est_tokens": self.est_tokens,
            "extract_method": self.extract_method,
            "error": self.error,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "MemberFile":
        return cls(
            source_path=d.get("source_path", ""),
            source_name=d.get("source_name", ""),
            blob=d.get("blob"),
            char_count=int(d.get("char_count", 0)),
            est_tokens=int(d.get("est_tokens", 0)),
            extract_method=d.get("extract_method", ""),
            error=d.get("error"),
        )


@dataclass
class LibraryEntry:
    """One task run (or manual add); the expandable unit in the UI."""
    id: str
    label: str
    auto_label: str
    task_type: str
    created_at: str
    members: list = field(default_factory=list)  # list[MemberFile]

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "label": self.label,
            "auto_label": self.auto_label,
            "task_type": self.task_type,
            "created_at": self.created_at,
            "members": [m.to_dict() for m in self.members],
        }

    @classmethod
    def from_dict(cls, d: dict) -> "LibraryEntry":
        return cls(
            id=d.get("id", ""),
            label=d.get("label", ""),
            auto_label=d.get("auto_label", d.get("label", "")),
            task_type=d.get("task_type", ""),
            created_at=d.get("created_at", ""),
            members=[MemberFile.from_dict(m) for m in d.get("members", [])],
        )
