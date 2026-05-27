"""Manage workbench-configured style exemplars for the oppose_motion pipeline.

The registry persists to ``Scripts/prompts/oppose_motion/style_examples.json``.
Each exemplar has a path to a .docx file, a free-form label, motion-type tags,
and an active flag. At draft time, ``matches_for_motion_type`` returns up to
N active exemplars whose tags appear as substrings of the current motion type
(plus any tag-less "universal" exemplars).
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
from dataclasses import dataclass, field, asdict
from typing import Any

logger = logging.getLogger(__name__)


@dataclass
class StyleExample:
    id: str = ""
    label: str = ""
    path: str = ""
    motion_types: list[str] = field(default_factory=list)
    active: bool = True
    added_at: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "StyleExample":
        return cls(
            id=str(data.get("id", "")),
            label=str(data.get("label", "")),
            path=str(data.get("path", "")),
            motion_types=[str(t).strip().lower() for t in (data.get("motion_types") or []) if str(t).strip()],
            active=bool(data.get("active", True)),
            added_at=str(data.get("added_at", "")),
        )


class StyleExampleRegistry:
    def __init__(self, *, path: str, examples: list[StyleExample] | None = None) -> None:
        self.path = path
        self.examples: list[StyleExample] = list(examples or [])

    @classmethod
    def load(cls, path: str) -> "StyleExampleRegistry":
        if not os.path.exists(path):
            return cls(path=path, examples=[])
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, ValueError, json.JSONDecodeError):
            logger.warning("Could not read style_examples.json", exc_info=True)
            return cls(path=path, examples=[])
        examples = [StyleExample.from_dict(d) for d in (data.get("examples") or [])]
        return cls(path=path, examples=examples)

    def save(self) -> None:
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump({"examples": [e.to_dict() for e in self.examples]}, f, indent=2)

    def add(self, example: StyleExample) -> None:
        # Replace if id already present.
        for i, e in enumerate(self.examples):
            if e.id == example.id:
                self.examples[i] = example
                return
        self.examples.append(example)

    def update(self, example_id: str, **fields: Any) -> bool:
        for e in self.examples:
            if e.id == example_id:
                if "motion_types" in fields:
                    fields["motion_types"] = [t.strip().lower() for t in fields["motion_types"] if t.strip()]
                for k, v in fields.items():
                    setattr(e, k, v)
                return True
        return False

    def remove(self, example_id: str) -> bool:
        for i, e in enumerate(self.examples):
            if e.id == example_id:
                del self.examples[i]
                return True
        return False

    def matches_for_motion_type(
        self,
        motion_type: str,
        *,
        max_results: int = 3,
    ) -> list[StyleExample]:
        needle = (motion_type or "").strip().lower()
        matches: list[StyleExample] = []
        for e in self.examples:
            if not e.active:
                continue
            if not e.motion_types:
                # Universal exemplar.
                matches.append(e)
                continue
            if any(tag in needle for tag in e.motion_types):
                matches.append(e)
            if len(matches) >= max_results:
                break
        return matches[:max_results]


def _cache_key(path: str) -> str:
    try:
        mtime = os.path.getmtime(path)
    except OSError:
        mtime = 0.0
    digest = hashlib.sha1(f"{os.path.abspath(path)}|{mtime}".encode("utf-8")).hexdigest()
    return digest


def extract_exemplar_text(path: str, *, cache_dir: str) -> str:
    """Extract plain text from a .docx file, caching by path+mtime."""
    if not path or not os.path.isfile(path):
        return ""
    os.makedirs(cache_dir, exist_ok=True)
    cache_path = os.path.join(cache_dir, f"{_cache_key(path)}.txt")
    if os.path.exists(cache_path):
        try:
            with open(cache_path, "r", encoding="utf-8") as f:
                return f.read()
        except OSError:
            logger.warning("Could not read exemplar cache: %s", cache_path, exc_info=True)

    try:
        from icharlotte_core.document_processor import extract_docx_text
        text = extract_docx_text(path) or ""
    except Exception:
        # Fallback: plain-paragraph reader.
        try:
            from docx import Document
            doc = Document(path)
            text = "\n".join(p.text for p in doc.paragraphs if p.text)
        except Exception:
            logger.warning("Could not extract exemplar text from %s", path, exc_info=True)
            text = ""

    if text:
        try:
            with open(cache_path, "w", encoding="utf-8") as f:
                f.write(text)
        except OSError:
            logger.warning("Could not write exemplar cache: %s", cache_path, exc_info=True)
    return text
