"""Draft-time style-exemplar selection from the firm brief library.

Embeds the current motion's issue profile, finds the most similar past briefs of
the same motion type AND side, and returns trimmed argument-section excerpts to
feed the drafter as style/voice models. Degrades to [] when the index is absent.
"""
from __future__ import annotations

import hashlib
import logging
import os
import re
from typing import Callable, List, Optional

logger = logging.getLogger(__name__)

_ARG_HEADING_RE = re.compile(r"(?im)^\s*(?:[IVXLC]+\.\s*)?(ARGUMENT|MEMORANDUM OF POINTS|LEGAL ARGUMENT|DISCUSSION)\b")


def _trim_to_argument(text: str, max_chars: int) -> str:
    """Start at the ARGUMENT/Memorandum heading if present (skips caption/TOC),
    then cap to max_chars."""
    text = text or ""
    m = _ARG_HEADING_RE.search(text)
    if m:
        text = text[m.start():]
    return text[:max_chars].strip()


def _default_excerpt_extract(path: str) -> str:
    from icharlotte_core.document_processor import DocumentProcessor
    try:
        return DocumentProcessor().extract_text(path, ocr_enabled=False).text or ""
    except Exception:
        logger.warning("style excerpt extract failed: %s", path, exc_info=True)
        return ""


def _cache_key(path: str) -> str:
    try:
        mtime = os.path.getmtime(path)
    except OSError:
        mtime = 0.0
    return hashlib.sha1(f"{os.path.abspath(path)}|{mtime}".encode("utf-8")).hexdigest()


def _excerpt(path: str, *, cache_dir: str, extract_fn: Callable[[str], str],
             max_chars: int, stored_text: str = "") -> str:
    os.makedirs(cache_dir, exist_ok=True)
    cache_path = os.path.join(cache_dir, f"{_cache_key(path)}.txt")
    if os.path.exists(cache_path):
        try:
            with open(cache_path, "r", encoding="utf-8") as f:
                return f.read()
        except OSError:
            pass
    raw = stored_text if (stored_text and stored_text.strip()) else extract_fn(path)
    excerpt = _trim_to_argument(raw, max_chars)
    if excerpt:
        try:
            with open(cache_path, "w", encoding="utf-8") as f:
                f.write(excerpt)
        except OSError:
            logger.warning("could not cache style excerpt: %s", cache_path, exc_info=True)
    return excerpt


def select_exemplars(
    motion_type: str,
    side: str,
    motion_metadata,
    *,
    k: int = 3,
    max_chars: int = 8000,
    index=None,
    embedder=None,
    extract_fn: Optional[Callable[[str], str]] = None,
    cache_dir: Optional[str] = None,
) -> List[str]:
    """Return up to k trimmed style excerpts most similar to the current motion."""
    from .factory import make_index, DATA_DIR
    from .profile import profile_from_metadata
    from .embedding import get_embedder

    if index is None:
        index = make_index()
    if index is None:
        return []
    prof = profile_from_metadata(motion_metadata)
    if not prof.strip():
        return []
    try:
        embedder = embedder or get_embedder()
        qv = embedder.encode([prof])[0]
        cands = index.style_candidates(qv, motion_type=motion_type, side=side, k=k)
    except Exception:
        logger.warning("style candidate selection failed", exc_info=True)
        return []
    extract_fn = extract_fn or _default_excerpt_extract
    cache_dir = cache_dir or os.path.join(DATA_DIR, ".cache", "style")
    out: List[str] = []
    for c in cands:
        stored = ""
        try:
            stored = index.get_full_text(c["path"])
        except Exception:
            stored = ""
        txt = _excerpt(c["path"], cache_dir=cache_dir, extract_fn=extract_fn,
                       max_chars=max_chars, stored_text=stored)
        if txt.strip():
            out.append(txt)
    return out
