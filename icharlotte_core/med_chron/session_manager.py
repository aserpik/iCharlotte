"""Session JSON management for the two-phase Med-Cron agent.

Phase 1 (prep) writes the session and pauses. The wizard UI fills in
``user_config`` and flips ``phase`` to ``"ready_to_run"``. Phase 2 (run)
reads the session and produces output.

Cache layout, scoped to the case's output directory:

    <output_dir>/.med_chron/<file_hash>/
        narrative.txt
        full.txt
        session.json

``<file_hash>`` is sha1(abspath + mtime_ns), truncated to 12 hex chars,
so touching the source file invalidates the cache.
"""

import hashlib
import json
import os
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class SessionPaths:
    cache_dir: Path
    session_path: Path
    narrative_text_path: Path
    full_text_path: Path


def _file_hash(input_path: str) -> str:
    abspath = os.path.abspath(input_path)
    try:
        mtime_ns = os.stat(input_path).st_mtime_ns
    except FileNotFoundError:
        mtime_ns = 0
    raw = f"{abspath}|{mtime_ns}".encode("utf-8")
    return hashlib.sha1(raw).hexdigest()[:12]


def compute_session_paths(input_path: str, output_dir: str) -> SessionPaths:
    """Return the cache + session paths for an input file under output_dir."""
    cache_dir = Path(output_dir) / ".med_chron" / _file_hash(input_path)
    return SessionPaths(
        cache_dir=cache_dir,
        session_path=cache_dir / "session.json",
        narrative_text_path=cache_dir / "narrative.txt",
        full_text_path=cache_dir / "full.txt",
    )


def write_session(session_path: str | Path, data: dict) -> None:
    """Atomically write the session JSON via tmp file + os.replace."""
    session_path = Path(session_path)
    session_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = session_path.with_suffix(session_path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
    os.replace(tmp, session_path)


def read_session(session_path: str | Path) -> dict:
    """Read and return the session JSON. Raises FileNotFoundError if the file is absent."""
    return json.loads(Path(session_path).read_text(encoding="utf-8"))


def update_user_config(session_path: str | Path, user_config: dict) -> None:
    """Load, set user_config, flip phase to 'ready_to_run', write atomically."""
    data = read_session(session_path)
    data["user_config"] = user_config
    data["phase"] = "ready_to_run"
    write_session(session_path, data)
