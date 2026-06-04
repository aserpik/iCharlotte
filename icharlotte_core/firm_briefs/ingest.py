# icharlotte_core/firm_briefs/ingest.py
"""Incrementally ingest a sorted firm-brief library root into a FirmBriefIndex."""
from __future__ import annotations

import hashlib
import os
from typing import Callable, Optional

from .citation_harvest import harvest_cites
from .path_meta import meta_for_path
from .profile import extract_headings, compose_profile, profile_from_text


def content_hash(path: str) -> str:
    try:
        st = os.stat(path)
    except OSError:
        return ""
    h = hashlib.sha1()
    h.update(f"{os.path.abspath(path)}|{st.st_mtime_ns}|{st.st_size}".encode("utf-8"))
    return h.hexdigest()


def _default_extract(path: str) -> str:
    from icharlotte_core.document_processor import DocumentProcessor
    try:
        # extract_text returns an ExtractResult; .text is the extracted string.
        return DocumentProcessor().extract_text(path).text or ""
    except Exception:
        return ""


def _ocr_ratio(text: str) -> float:
    # Crude noise signal: share of non-ASCII / replacement chars in the text.
    if not text:
        return 1.0
    bad = sum(1 for ch in text if ord(ch) > 0x2122 or ch == "�")
    return bad / len(text)


def ingest_root(root: str, index, embedder, *,
                extract_fn: Optional[Callable[[str], str]] = None,
                on_progress: Optional[Callable[[str], None]] = None) -> dict:
    extract_fn = extract_fn or _default_extract
    added = updated = skipped = failed = 0
    seen_paths: set[str] = set()

    for dirpath, _dirs, files in os.walk(root):
        for name in files:
            if not name.lower().endswith(".pdf"):
                continue
            path = os.path.join(dirpath, name)
            meta = meta_for_path(path, root)
            if meta is None:
                continue  # _Support / _Other / unrecognized
            motion_type, side = meta
            seen_paths.add(os.path.abspath(path))
            h = content_hash(path)
            if index.has_current(path, h):
                skipped += 1
                continue
            text = extract_fn(path)
            if not text.strip():
                failed += 1
                continue
            cites = harvest_cites(text)
            headings = extract_headings(text)
            profile = compose_profile("", headings, [c.proposition for c in cites]) \
                or profile_from_text(text)
            vec = embedder.encode([profile])[0]
            # Corrected: check if path already exists in DB (any hash) to distinguish add vs update
            existed = bool(index._conn().execute(
                "SELECT 1 FROM briefs WHERE path=?", (path,)).fetchone())
            index.upsert_brief(
                path=path, content_hash=h, motion_type=motion_type, side=side,
                heading=headings[0] if headings else "", profile=profile,
                profile_vec=vec, char_len=len(text), ocr_ratio=_ocr_ratio(text),
                cites=cites,
            )
            updated += 1 if existed else 0
            added += 0 if existed else 1
            if on_progress:
                on_progress(f"  indexed {name} ({motion_type}/{side}, {len(cites)} cites)")

    # Mark DB briefs under this root that no longer exist on disk as stale.
    staled = 0
    con = index._conn()
    rows = con.execute(
        "SELECT path FROM briefs WHERE status='ok' AND path LIKE ?",
        (os.path.join(os.path.abspath(root), "") + "%",),
    ).fetchall()
    for r in rows:
        if os.path.abspath(r["path"]) not in seen_paths and not os.path.exists(r["path"]):
            index.mark_stale(r["path"])
            staled += 1

    return {"added": added, "updated": updated, "skipped": skipped,
            "failed": failed, "staled": staled}
