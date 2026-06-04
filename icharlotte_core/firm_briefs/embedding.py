"""Reuse the local-corpus fastembed embedder (BGE-small, 384-dim, no torch)."""
from __future__ import annotations

from icharlotte_core.legal_research.local_corpus.embedder import (
    FakeEmbedder,
    OnnxEmbedder,
    cosine_topk,
)

EMBED_DIM = 384


def get_embedder(*, fake: bool = False):
    """Return an embedder. ``fake=True`` for tests (deterministic, no model)."""
    if fake:
        return FakeEmbedder(dim=EMBED_DIM)
    return OnnxEmbedder(dim=EMBED_DIM)


__all__ = ["get_embedder", "cosine_topk", "FakeEmbedder", "EMBED_DIM"]
