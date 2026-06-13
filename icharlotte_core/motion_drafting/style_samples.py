"""Style samples loaded from the selected Motion Database taxonomy folder."""

from __future__ import annotations

import hashlib
import logging
import os
import re
from pathlib import Path
from typing import Callable

logger = logging.getLogger(__name__)

SUPPORTED_SAMPLE_EXTENSIONS = {".docx", ".pdf", ".txt", ".md"}

_ARG_HEADING_RE = re.compile(
    r"(?im)^\s*(?:[IVXLC]+\.\s*)?"
    r"(ARGUMENT|MEMORANDUM OF POINTS|LEGAL ARGUMENT|DISCUSSION)\b"
)


def motion_database_style_cache_dir() -> str:
    repo_root = Path(__file__).resolve().parents[2]
    return str(repo_root / "Scripts" / "prompts" / "motion_drafting" / ".cache" / "style_samples")


def load_motion_database_style_samples(
    source_path: str | os.PathLike[str] | None,
    *,
    max_samples: int = 3,
    max_chars: int = 8000,
    cache_dir: str | os.PathLike[str] | None = None,
    extract_fn: Callable[[str], str] | None = None,
) -> list[str]:
    """Return trimmed style excerpts from the selected Motion Database folder."""
    source = Path(source_path or "")
    if max_samples <= 0 or not source.is_dir():
        return []

    cache_root = str(cache_dir or motion_database_style_cache_dir())
    extractor = extract_fn or _default_extract_text
    out: list[str] = []
    for path in _iter_sample_files(source):
        text = _cached_excerpt(
            path,
            cache_dir=cache_root,
            extract_fn=extractor,
            max_chars=max_chars,
        )
        if text.strip():
            out.append(text)
        if len(out) >= max_samples:
            break
    return out


def _iter_sample_files(source: Path):
    files = [
        path
        for path in source.rglob("*")
        if path.is_file() and path.suffix.lower() in SUPPORTED_SAMPLE_EXTENSIONS
    ]
    base_parts = len(source.parts)
    yield from sorted(files, key=lambda path: (len(path.parts) - base_parts, path.name.lower(), str(path).lower()))


def _cached_excerpt(
    path: Path,
    *,
    cache_dir: str,
    extract_fn: Callable[[str], str],
    max_chars: int,
) -> str:
    os.makedirs(cache_dir, exist_ok=True)
    cache_path = os.path.join(cache_dir, f"{_cache_key(path)}.txt")
    if os.path.exists(cache_path):
        try:
            with open(cache_path, "r", encoding="utf-8") as f:
                return f.read()
        except OSError:
            logger.warning("Could not read Motion Database style cache: %s", cache_path, exc_info=True)

    try:
        raw = extract_fn(str(path)) or ""
    except Exception:
        logger.warning("Could not extract Motion Database style sample: %s", path, exc_info=True)
        raw = ""

    excerpt = _trim_to_argument(raw, max_chars)
    if excerpt:
        try:
            with open(cache_path, "w", encoding="utf-8") as f:
                f.write(excerpt)
        except OSError:
            logger.warning("Could not write Motion Database style cache: %s", cache_path, exc_info=True)
    return excerpt


def _default_extract_text(path: str) -> str:
    suffix = Path(path).suffix.lower()
    if suffix in {".txt", ".md"}:
        try:
            return Path(path).read_text(encoding="utf-8", errors="ignore")
        except OSError:
            return ""
    try:
        from icharlotte_core.document_processor import DocumentProcessor

        return DocumentProcessor().extract_text(path, ocr_enabled=False).text or ""
    except Exception:
        logger.warning("Motion Database sample extraction failed: %s", path, exc_info=True)
        return ""


def _trim_to_argument(text: str, max_chars: int) -> str:
    text = text or ""
    match = _ARG_HEADING_RE.search(text)
    if match:
        text = text[match.start():]
    return text[:max_chars].strip()


def _cache_key(path: Path) -> str:
    try:
        stat = path.stat()
        stamp = f"{stat.st_mtime_ns}|{stat.st_size}"
    except OSError:
        stamp = "missing"
    return hashlib.sha1(f"{path.resolve()}|{stamp}".encode("utf-8")).hexdigest()
