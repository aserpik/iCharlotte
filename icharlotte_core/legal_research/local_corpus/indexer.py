"""Write normalized records into the corpus DB + vectors.f16 memmap.

Usage: create, .add(case, passages) per case, then .finalize(). Embeddings are
batched and appended to a float16 sidecar; each passage row records its vec_row.
"""
from __future__ import annotations

import sqlite3
from typing import Iterable

import numpy as np

from icharlotte_core.legal_research.local_corpus.embedder import Embedder
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord

_BATCH = 256


class CorpusIndexer:
    def __init__(self, con: sqlite3.Connection, *, vectors_path: str, embedder: Embedder) -> None:
        self.con = con
        self.vectors_path = vectors_path
        self.embedder = embedder
        self._pending: list[PassageRecord] = []
        self._vec_blocks: list[np.ndarray] = []
        self._next_vec_row = 0
        self._seen_citation: set[str] = set()

    def add(self, case: CaseRecord, passages: Iterable[PassageRecord]) -> bool:
        """Insert one case + its passages. Returns False if deduped (skipped)."""
        # Cross-source dedup by normalized citation (first writer wins).
        norm = (case.citation or "").replace(" ", "").lower()
        if norm and norm in self._seen_citation:
            return False
        if norm:
            self._seen_citation.add(norm)
        row = case.to_row()
        self.con.execute(
            "INSERT OR REPLACE INTO cases (%s) VALUES (%s)" % (
                ",".join(row.keys()),
                ",".join(["?"] * len(row)),
            ),
            list(row.values()),
        )
        for p in passages:
            self._pending.append(p)
            if len(self._pending) >= _BATCH:
                self._flush()
        return True

    def _flush(self) -> None:
        if not self._pending:
            return
        vecs = self.embedder.encode([p.text for p in self._pending]).astype(np.float16)
        self._vec_blocks.append(vecs)
        for p in self._pending:
            vec_row = self._next_vec_row
            self._next_vec_row += 1
            self.con.execute(
                "INSERT OR REPLACE INTO passages (passage_uid, case_uid, ordinal, text, page_label, vec_row) "
                "VALUES (?,?,?,?,?,?)",
                (p.passage_uid, p.case_uid, p.ordinal, p.text, p.page_label, vec_row),
            )
            self.con.execute(
                "INSERT INTO passages_fts (rowid, text) VALUES (?, ?)",
                (vec_row + 1, p.text),   # fts rowid aligned to vec_row+1 (1-based)
            )
        self._pending.clear()

    def finalize(self) -> None:
        self._flush()
        self.con.commit()
        if self._vec_blocks:
            allvecs = np.concatenate(self._vec_blocks, axis=0).astype(np.float16)
        else:
            allvecs = np.zeros((0, self.embedder.dim), dtype=np.float16)
        mm = np.memmap(self.vectors_path, dtype=np.float16, mode="w+", shape=allvecs.shape)
        mm[:] = allvecs[:]
        mm.flush()
        del mm
