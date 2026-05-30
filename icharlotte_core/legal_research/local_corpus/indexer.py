"""Write normalized records into the corpus DB + vectors.f16 sidecar.

Resumable, volume-checkpointed build (see commit_volume/abort_volume). On
restart, construct with ``resume=True`` and already-ingested volumes are skipped.

``embed_year_cutoff``: cases decided before this year are still fully indexed
for keyword (FTS5) search, but get a ZERO placeholder vector instead of a real
embedding — skipping the expensive transformer call for old, rarely-cited cases
while keeping vec_row<->passage alignment intact. Semantic search never returns
them (zero vector => zero cosine); keyword search still does.

Crash-safety ordering: each batch writes + fsyncs vectors BEFORE the DB commit,
so committed passages always have their vector row on disk; on resume the
sidecar is truncated to the committed passage count.
"""
from __future__ import annotations

import os
import sqlite3
from typing import Iterable

import numpy as np

from icharlotte_core.legal_research.local_corpus.embedder import Embedder
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord

_BATCH = 256


class CorpusIndexer:
    def __init__(self, con: sqlite3.Connection, *, vectors_path: str, embedder: Embedder,
                 embed: bool = True, embed_year_cutoff: int | None = None,
                 resume: bool = False) -> None:
        self.con = con
        self.vectors_path = vectors_path
        self.embedder = embedder
        # embed=False => FTS5-only (no vectors at all). embed=True with a cutoff
        # => vectors for cases >= cutoff, zero placeholders for older ones.
        self.embed = embed
        self.embed_year_cutoff = embed_year_cutoff
        self._pending: list[tuple[PassageRecord, bool]] = []  # (passage, do_embed)
        self._next_vec_row = 0
        self._seen_citation: set[str] = set()
        self._done_volumes: set[str] = set()
        self._bytes_per_vec = self.embedder.dim * 2  # float16

        if resume:
            self._resume()
        else:
            self._vec_fh = open(vectors_path, "wb")

    # ------------------------------------------------------------------
    def _resume(self) -> None:
        n_committed = self.con.execute("SELECT COUNT(*) FROM passages").fetchone()[0]
        self._next_vec_row = n_committed
        self._done_volumes = {
            r[0] for r in self.con.execute("SELECT name FROM ingested_volumes")
        }
        self._seen_citation = {
            (r[0] or "").replace(" ", "").lower()
            for r in self.con.execute("SELECT citation FROM cases")
            if r[0]
        }
        if self.embed:
            # Truncate the sidecar to exactly the committed passage count, then append.
            want_bytes = n_committed * self._bytes_per_vec
            if not os.path.exists(self.vectors_path):
                open(self.vectors_path, "wb").close()
            with open(self.vectors_path, "r+b") as f:
                f.truncate(want_bytes)
            self._vec_fh = open(self.vectors_path, "ab")
        else:
            # FTS-only: no vector sidecar in use.
            self._vec_fh = open(self.vectors_path, "ab")

    # ------------------------------------------------------------------
    def is_volume_done(self, name: str) -> bool:
        return name in self._done_volumes

    def _should_embed_case(self, case: CaseRecord) -> bool:
        if not self.embed:
            return False
        if self.embed_year_cutoff is None:
            return True
        y = (case.year or "")[:4]
        return y.isdigit() and int(y) >= self.embed_year_cutoff

    def add(self, case: CaseRecord, passages: Iterable[PassageRecord]) -> bool:
        """Buffer one case + its passages (uncommitted until commit_volume).

        Returns False if deduped (citation already seen)."""
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
        do_embed = self._should_embed_case(case)
        for p in passages:
            self._pending.append((p, do_embed))
            if len(self._pending) >= _BATCH:
                self._flush()
        return True

    def _flush(self) -> None:
        if not self._pending:
            return
        if self.embed:
            # Build the batch's vector block: real embeddings for to-embed
            # passages, zero placeholders for the rest. Keeps one vector row per
            # passage so vec_row <-> passage alignment is exact.
            out = np.zeros((len(self._pending), self.embedder.dim), dtype=np.float16)
            embed_idx = [i for i, (_p, e) in enumerate(self._pending) if e]
            if embed_idx:
                vecs = self.embedder.encode(
                    [self._pending[i][0].text for i in embed_idx]
                ).astype(np.float16)
                for j, i in enumerate(embed_idx):
                    out[i] = vecs[j]
            self._vec_fh.write(np.ascontiguousarray(out, dtype=np.float16).tobytes())
        for p, _e in self._pending:
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

    def commit_volume(self, name: str) -> None:
        """Checkpoint: flush the buffer, fsync vectors, then commit the volume."""
        self._flush()
        self._vec_fh.flush()
        os.fsync(self._vec_fh.fileno())
        self.con.execute("INSERT OR IGNORE INTO ingested_volumes (name) VALUES (?)", (name,))
        self.con.commit()
        self._done_volumes.add(name)

    def abort_volume(self) -> None:
        """Roll back the current uncommitted volume and reconcile state."""
        self.con.rollback()
        self._pending.clear()
        n_committed = self.con.execute("SELECT COUNT(*) FROM passages").fetchone()[0]
        self._next_vec_row = n_committed
        self._vec_fh.flush()
        self._vec_fh.close()
        if self.embed:
            with open(self.vectors_path, "r+b") as f:
                f.truncate(n_committed * self._bytes_per_vec)
        self._vec_fh = open(self.vectors_path, "ab")
        self._seen_citation = {
            (r[0] or "").replace(" ", "").lower()
            for r in self.con.execute("SELECT citation FROM cases")
            if r[0]
        }

    def finalize(self) -> None:
        self._flush()
        self.con.commit()
        self._vec_fh.flush()
        self._vec_fh.close()
