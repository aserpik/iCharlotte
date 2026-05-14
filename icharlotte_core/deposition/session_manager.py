"""Sidecar JSON session management for the interactive deposition summary flow.

Phase 1 of the deposition agent writes a session file describing the
proposed topics and pauses. The UI loads it, lets the user edit, writes
back a user_config block, and launches phase 2. Phase 2 reads the
session, generates the summary, and cleans up.
"""

import hashlib
import json
import os
import secrets
import time
from dataclasses import dataclass
from pathlib import Path

# Resolved relative to the iCharlotte project root (icharlotte_core/.. == project)
_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
SESSION_DIR = _PROJECT_ROOT / "logs" / "depo_sessions"


@dataclass(frozen=True)
class SessionPaths:
    session_path: Path
    cached_text_path: Path


def compute_session_paths(input_path: str) -> SessionPaths:
    """Build a unique (session_json, cached_text) path pair for an input file."""
    digest = hashlib.sha1(os.fspath(input_path).encode("utf-8")).hexdigest()[:12]
    ts = time.strftime("%Y%m%d_%H%M%S")
    nonce = secrets.token_hex(3)
    base = f"{digest}_{ts}_{nonce}"
    SESSION_DIR.mkdir(parents=True, exist_ok=True)
    return SessionPaths(
        session_path=SESSION_DIR / f"{base}.json",
        cached_text_path=SESSION_DIR / f"{base}.txt",
    )


def write_session(session_path: Path, data: dict) -> None:
    """Atomically write the session JSON via tmp file + os.replace."""
    session_path = Path(session_path)
    session_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = session_path.with_suffix(session_path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
    os.replace(tmp, session_path)


def read_session(session_path: Path) -> dict:
    return json.loads(Path(session_path).read_text(encoding="utf-8"))


def update_user_config(session_path: Path, user_config: dict) -> None:
    """Load, set user_config, flip phase to 'ready_for_summary', write atomically."""
    data = read_session(session_path)
    data["user_config"] = user_config
    data["phase"] = "ready_for_summary"
    write_session(session_path, data)


def cleanup_session(session_path: Path) -> None:
    """Delete the session JSON and its cached transcript. Tolerant of missing files."""
    session_path = Path(session_path)
    try:
        data = read_session(session_path)
        cached = data.get("cached_text_path")
        if cached:
            cached_path = Path(cached)
            if cached_path.exists():
                cached_path.unlink()
    except (FileNotFoundError, json.JSONDecodeError):
        pass
    if session_path.exists():
        session_path.unlink()
