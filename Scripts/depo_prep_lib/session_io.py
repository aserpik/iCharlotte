"""Session folder layout + JSON helpers for Depo Prep."""
from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Union


# Windows-illegal filename characters (matches summarize_deposition.py convention).
_UNSAFE_FILENAME_CHARS = re.compile(r'[\\/*?:"<>|]')


@dataclass(frozen=True)
class SessionPaths:
    session_dir: Path
    session_json: Path
    topics_json: Path
    digests_dir: Path
    raw_dir: Path
    outline_docx: Path
    outline_md: Path
    trace_log: Path


def build_session_folder_name(deponent_name: str, when_iso: str | None = None) -> str:
    """Return a Windows-safe folder name like 'Depo Prep - Jane Doe - 2026-05-27 1432'."""
    when_iso = (when_iso or datetime.now().isoformat(timespec="minutes"))[:16]
    # Strip illegal chars; collapse whitespace.
    safe_name = _UNSAFE_FILENAME_CHARS.sub("", deponent_name or "Unknown").strip()
    safe_name = re.sub(r"\s+", " ", safe_name) or "Unknown"
    safe_name = safe_name[:60].rstrip()
    # "2026-05-27T14:32" → "2026-05-27 1432"
    when = when_iso.replace("T", " ").replace(":", "")
    return f"Depo Prep - {safe_name} - {when}"


def compute_session_paths(case_root: str, deponent_name: str, when_iso: str | None = None) -> SessionPaths:
    """Compute the canonical layout for a Depo Prep session.

    Layout: {case_root}/NOTES/AI Output/{folder_name}/
              session.json
              topics.json
              digests/
                raw/<source>.txt
                <source>.json
              outline.docx
              outline.md
              trace.log
    """
    folder = build_session_folder_name(deponent_name, when_iso=when_iso)
    session_dir = Path(case_root) / "NOTES" / "AI Output" / folder
    digests = session_dir / "digests"
    return SessionPaths(
        session_dir=session_dir,
        session_json=session_dir / "session.json",
        topics_json=session_dir / "topics.json",
        digests_dir=digests,
        raw_dir=digests / "raw",
        outline_docx=session_dir / "outline.docx",
        outline_md=session_dir / "outline.md",
        trace_log=session_dir / "trace.log",
    )


def write_json(path: Union[str, Path], data) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def read_json(path: Union[str, Path]):
    with Path(path).open("r", encoding="utf-8") as f:
        return json.load(f)


def file_sha256(path: Union[str, Path]) -> str:
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()
