import numpy as np
from icharlotte_core.firm_briefs.embedding import get_embedder, EMBED_DIM
from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder


def test_fake_embedder_shape():
    emb = FakeEmbedder(dim=EMBED_DIM)
    vecs = emb.encode(["meet and confer", "discovery cutoff"])
    assert vecs.shape == (2, EMBED_DIM)
    assert vecs.dtype == np.float32


def test_get_embedder_returns_object_with_encode():
    emb = get_embedder(fake=True)
    assert hasattr(emb, "encode")
    assert emb.dim == EMBED_DIM
