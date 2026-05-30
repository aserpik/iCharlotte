import numpy as np
from icharlotte_core.legal_research.local_corpus.embedder import (
    Embedder, FakeEmbedder, cosine_topk,
)


def test_fake_embedder_is_deterministic_and_normalized():
    emb = FakeEmbedder(dim=16)
    a = emb.encode(["duty of care"])
    b = emb.encode(["duty of care"])
    assert a.shape == (1, 16)
    np.testing.assert_allclose(a, b)               # deterministic
    np.testing.assert_allclose(np.linalg.norm(a[0]), 1.0, atol=1e-5)  # unit norm


def test_fake_embedder_satisfies_protocol():
    assert isinstance(FakeEmbedder(dim=8), Embedder)


def test_cosine_topk_ranks_by_similarity():
    mat = np.array([[1.0, 0.0], [0.0, 1.0], [0.7, 0.7]], dtype=np.float32)
    q = np.array([1.0, 0.0], dtype=np.float32)
    idx, scores = cosine_topk(q, mat, k=2)
    assert idx[0] == 0           # identical vector ranks first
    assert idx[1] == 2           # the 45-degree one beats the orthogonal one
