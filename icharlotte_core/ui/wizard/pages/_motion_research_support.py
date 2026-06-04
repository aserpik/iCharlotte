"""Shared research/style helpers for the Oppose- and Generate-a-Motion pages.

Extracted so both pages depend on one module instead of generate importing
oppose's private helpers. Behavior is identical to the previous oppose_motion_page
definitions.
"""
from __future__ import annotations

import os
import re

import os as _os_corpus
from icharlotte_core.config import CASELAW_DATA_DIR


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
    seen: set[str] = set()

    def _add(text: str) -> None:
        t = (text or "").strip()
        if not t:
            return
        key = re.sub(r"\s+", " ", t.lower())
        if key in seen:
            return
        seen.add(key)
        targets.append(t)

    for arg in (getattr(metadata, "principal_arguments", None) or []):
        _add(arg)
    # Structural sections that argue no legal point and need no case authority.
    _skip = ("introduction", "conclusion", "statement of facts",
             "factual background", "preliminary statement", "prayer")
    for item in (plan or []):
        text = (getattr(item, "text", "") or "").strip()
        if not text or any(s in text.lower() for s in _skip):
            continue
        _add(text)
    return targets[:24]


def make_local_corpus():
    if not _corpus_available():
        return None
    from icharlotte_core.legal_research.local_corpus.corpus import LocalCaseCorpus
    db, vec = _corpus_paths()
    return LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=_corpus_embedder())


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
