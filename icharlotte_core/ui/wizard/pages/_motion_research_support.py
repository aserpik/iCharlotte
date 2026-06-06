"""Shared research/style helpers for the Oppose- and Generate-a-Motion pages.

Extracted so both pages depend on one module instead of generate importing
oppose's private helpers. Behavior is identical to the previous oppose_motion_page
definitions.
"""
from __future__ import annotations

from datetime import date, datetime
import json
import os
import re

import os as _os_corpus
from icharlotte_core.config import CASELAW_DATA_DIR


_FRESHNESS_MAX_AGE_DAYS = 548


_STRUCTURAL_TARGETS = {
    "argument",
    "legal argument",
    "legal standard",
    "standard of review",
    "introduction",
    "conclusion",
    "statement of facts",
    "factual background",
    "preliminary statement",
    "prayer",
}

_TARGET_STOPWORDS = {
    "the", "and", "for", "with", "without", "into", "from", "that", "this",
    "cause", "causes", "action", "actions", "claim", "claims", "complaint",
    "amended", "first", "second", "third", "fourth", "fifth", "sixth",
    "seventh", "eighth", "ninth", "tenth",
}


def _corpus_paths() -> tuple[str, str]:
    return (_os_corpus.path.join(CASELAW_DATA_DIR, "corpus.db"),
            _os_corpus.path.join(CASELAW_DATA_DIR, "vectors.f16"))


def _corpus_available() -> bool:
    # Require BOTH files: corpus.db is created (empty) at the start of a build,
    # but vectors.f16 is only written at finalize(). Requiring both means an
    # in-progress or partial build safely falls back to the live API instead of
    # using an empty corpus with a missing vector sidecar.
    db, vec = _corpus_paths()
    return _os_corpus.path.exists(db) and _os_corpus.path.exists(vec)


def _corpus_embedder():
    from icharlotte_core.legal_research.local_corpus.embedder import OnnxEmbedder
    return OnnxEmbedder()


def _target_heading_key(text: str) -> str:
    text = re.sub(r"^\s*(?:[A-Z]\.|[IVXLC]+\.)\s*", "", text or "", flags=re.I)
    text = re.sub(r"^\s*\d+\.\s*", "", text)
    text = re.sub(r"[^a-z0-9]+", " ", text.lower())
    return re.sub(r"\s+", " ", text).strip()


def _target_tokens(text: str) -> set[str]:
    key = _target_heading_key(text)
    return {
        token
        for token in key.split()
        if len(token) > 2 and token not in _TARGET_STOPWORDS
    }


def _same_research_point(left: set[str], right: set[str]) -> bool:
    if not left or not right:
        return False
    overlap = len(left & right)
    return (overlap / min(len(left), len(right))) >= 0.72


def research_targets(metadata, plan) -> list[str]:
    """Propositions to research, deduped: the union of the principal arguments
    and every selected section-plan leaf.

    The drafter expands the brief into one subsection per section-plan leaf, and
    each subsection makes its own legal proposition (meet-and-confer, discovery
    cutoff, cumulative discovery, ...). Researching only the top-level arguments
    left those sub-points ungrounded — the drafter then emitted "[no case
    authority retrieved for this point]". Researching each leaf gives every
    subsection its own on-point authority. Purely structural sections are
    skipped; the count is capped to bound LLM calls under provider rate limits.
    """
    targets: list[str] = []
    target_token_sets: list[set[str]] = []
    seen: set[str] = set()

    def _add(text: str) -> None:
        t = (text or "").replace("\x00", "").strip()
        if not t:
            return
        key = _target_heading_key(t)
        if key in _STRUCTURAL_TARGETS:
            return
        if key in seen:
            return
        tokens = _target_tokens(t)
        for idx, existing in enumerate(target_token_sets):
            if _same_research_point(tokens, existing):
                if len(tokens) > len(existing) + 2:
                    seen.discard(_target_heading_key(targets[idx]))
                    targets[idx] = t
                    target_token_sets[idx] = tokens
                    seen.add(key)
                return
        seen.add(key)
        targets.append(t)
        target_token_sets.append(tokens)

    for arg in (getattr(metadata, "principal_arguments", None) or []):
        _add(arg)
    # Structural sections that argue no legal point and need no case authority.
    for item in (plan or []):
        text = (getattr(item, "text", "") or "").strip()
        if not text:
            continue
        _add(text)
    return targets[:24]


def make_local_corpus():
    if not _corpus_available():
        return None
    from icharlotte_core.legal_research.local_corpus.corpus import LocalCaseCorpus
    db, vec = _corpus_paths()
    return LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=_corpus_embedder())


def _parse_date(value) -> date | None:
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    if not text:
        return None
    try:
        return datetime.strptime(text[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def corpus_freshness_status(corpus, *, today=None, max_age_days: int = _FRESHNESS_MAX_AGE_DAYS) -> dict:
    """Return whether local case-law data is current enough to be primary.

    Unknown test doubles are treated as fresh so unit tests that use simple
    objects do not accidentally exercise the live fallback path.
    """
    if corpus is None:
        return {"fresh": False, "reason": "No local corpus.", "metadata": {}}
    if not hasattr(corpus, "corpus_metadata"):
        return {"fresh": True, "reason": "Corpus metadata unavailable.", "metadata": {}}
    try:
        metadata = corpus.corpus_metadata() or {}
    except Exception:
        return {"fresh": True, "reason": "Corpus metadata unavailable.", "metadata": {}}
    source_counts = metadata.get("source_counts") or {}
    if isinstance(source_counts, str):
        try:
            source_counts = json.loads(source_counts)
        except ValueError:
            source_counts = {}
    cl_count = int(source_counts.get("cl") or 0)
    if cl_count <= 0:
        return {
            "fresh": False,
            "reason": "Local corpus has no CourtListener recent slice.",
            "metadata": metadata,
        }
    max_date = _parse_date(metadata.get("max_decision_date"))
    today_date = _parse_date(today) or date.today()
    if not max_date:
        return {
            "fresh": False,
            "reason": "Local corpus has no max decision date metadata.",
            "metadata": metadata,
        }
    age_days = (today_date - max_date).days
    if age_days > max_age_days:
        return {
            "fresh": False,
            "reason": f"Local corpus is stale; newest decision is {max_date.isoformat()}.",
            "metadata": metadata,
        }
    return {"fresh": True, "reason": "", "metadata": metadata}


def select_research_client(corpus, token: str, *, on_progress=None):
    status = corpus_freshness_status(corpus)
    if corpus is not None and status.get("fresh"):
        return corpus, "local_corpus", status
    if corpus is not None and not status.get("fresh"):
        reason = status.get("reason") or "Local corpus is stale."
        if token:
            if on_progress:
                on_progress(f"WARNING: {reason} Using CourtListener API for current case law.")
            from icharlotte_core.legal_research.sources.courtlistener import CourtListenerClient
            return CourtListenerClient(token), "courtlistener", status
        if on_progress:
            on_progress(f"WARNING: {reason} No COURTLISTENER_API_TOKEN; using stale local corpus.")
        return corpus, "local_corpus", status
    if token:
        from icharlotte_core.legal_research.sources.courtlistener import CourtListenerClient
        return CourtListenerClient(token), "courtlistener", status
    return None, "none", status


def firm_style_exemplars(motion_type, side, metadata):
    """Firm-library style excerpts most similar to this motion; [] if no index."""
    try:
        from icharlotte_core.firm_briefs import style
        return style.select_exemplars(motion_type, side, metadata) or []
    except Exception:
        return []


def make_firm_provider(corpus):
    """Build a FirmAuthorityProvider if the firm-brief index is built, else None.

    cl_client is the live CourtListener fallback for firm cites not in the local
    corpus; reuse the same token the research path uses.
    """
    try:
        from icharlotte_core.firm_briefs import factory
        index = factory.make_index()
        if index is None:
            return None
        from icharlotte_core.firm_briefs.provider import FirmAuthorityProvider
        cl = None
        token = os.environ.get("COURTLISTENER_API_TOKEN", "").strip()
        if token:
            from icharlotte_core.legal_research.sources.courtlistener import CourtListenerClient
            cl = CourtListenerClient(token)
        return FirmAuthorityProvider(index, corpus, cl_client=cl)
    except Exception:
        return None
