"""Swappable text embedder. Default = ONNX BGE-small via fastembed (no torch).

`Embedder` is a runtime-checkable Protocol so a stronger model can be dropped
in later without touching the corpus/index code. `FakeEmbedder` gives CI a
deterministic, dependency-free embedder.
"""
from __future__ import annotations

import hashlib
import os
from typing import Protocol, runtime_checkable

import numpy as np


@runtime_checkable
class Embedder(Protocol):
    dim: int
    def encode(self, texts: list[str]) -> np.ndarray:  # (n, dim) float32, unit-norm
        ...


class FakeEmbedder:
    """Deterministic hash-based embedder for tests (no model download)."""
    def __init__(self, dim: int = 384) -> None:
        self.dim = dim

    def encode(self, texts: list[str]) -> np.ndarray:
        out = np.zeros((len(texts), self.dim), dtype=np.float32)
        for i, t in enumerate(texts):
            for tok in (t or "").lower().split():
                h = int(hashlib.md5(tok.encode()).hexdigest(), 16)
                out[i, h % self.dim] += 1.0
            n = np.linalg.norm(out[i])
            if n > 0:
                out[i] /= n
        return out


class OnnxEmbedder:
    """fastembed BGE-small (ONNX, CPU). Lazy-loads the model on first encode.

    ``threads`` sets onnxruntime intra-op parallelism; defaults to all logical
    cores. Building a ~150k-case corpus on the default (~1-2 threads) was ~30
    passages/sec; using all cores is several times faster. We deliberately do
    NOT use fastembed's multiprocessing ``parallel=`` — it deadlocks on Windows
    without an ``if __name__ == '__main__'`` guard.
    """
    def __init__(self, model_name: str = "BAAI/bge-small-en-v1.5", dim: int = 384,
                 threads: int | None = None, batch_size: int = 256) -> None:
        self.model_name = model_name
        self.dim = dim
        self.threads = threads if threads is not None else (os.cpu_count() or 4)
        self.batch_size = batch_size
        self._model = None

    def _ensure(self):
        if self._model is None:
            from fastembed import TextEmbedding  # lazy import; heavy
            self._model = TextEmbedding(model_name=self.model_name, threads=self.threads)
        return self._model

    def encode(self, texts: list[str]) -> np.ndarray:
        if not texts:
            return np.zeros((0, self.dim), dtype=np.float32)
        model = self._ensure()
        vecs = np.array(list(model.embed(texts, batch_size=self.batch_size)), dtype=np.float32)
        norms = np.linalg.norm(vecs, axis=1, keepdims=True)
        norms[norms == 0] = 1.0
        return vecs / norms


def cosine_topk(query: np.ndarray, matrix: np.ndarray, k: int) -> tuple[np.ndarray, np.ndarray]:
    """Top-k rows of `matrix` by cosine similarity to `query` (both assumed
    finite; query need not be unit-norm). Returns (indices, scores) desc."""
    if matrix.shape[0] == 0:
        return np.array([], dtype=int), np.array([], dtype=np.float32)
    q = query.astype(np.float32)
    qn = np.linalg.norm(q)
    if qn > 0:
        q = q / qn
    sims = matrix @ q                       # matrix rows assumed unit-norm
    k = min(k, sims.shape[0])
    idx = np.argpartition(-sims, k - 1)[:k]
    idx = idx[np.argsort(-sims[idx])]
    return idx, sims[idx]
