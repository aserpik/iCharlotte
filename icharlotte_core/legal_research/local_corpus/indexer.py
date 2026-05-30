"""Write normalized records into the corpus DB + vectors.f16 sidecar.

Resumable, volume-checkpointed build: ``add()`` cases for a volume, then
``commit_volume(name)`` to make that volume durable. On restart, construct with
``resume=True`` against the same DB + vectors file and already-ingested volumes
are skipped — so a crash/reboot mid-build loses at most the current volume,
never the whole run.

Crash-safety ordering: each batch writes + fsyncs its vectors to the sidecar
file BEFORE the DB rows are committed, so every committed passage always has its
vector on disk. On resume we truncate the sidecar back to the committed passage
count, dropping any vectors written for an uncommitted (partial) volume.
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
                 embed: bool = True, resume: bool = False) -> None:
        self.con = con
        self.vectors_path = vectors_path
        self.embedder = embedder
        # FTS5-only mode (embed=False) skips transformer embedding entirely.
        # Vectors stay empty; BM25 carries retrieval and a semantic layer can be
        # added later without a rebuild.
        self.embed = embed
        self._pending: list[PassageRecord] = []
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
        """Reconcile in-memory state + the vectors file with the committed DB."""
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
        # Truncate the sidecar to exactly the committed passage count (drop any
        # vectors written for an uncommitted partial volume), then append.
        want_bytes = n_committed * self._bytes_per_vec
        if not os.path.exists(self.vectors_path):
            open(self.vectors_path, "wb").close()
        with open(self.vectors_path, "r+b") as f:
            f.truncate(want_bytes)
        self._vec_fh = open(self.vectors_path, "ab")

    # ------------------------------------------------------------------
    def is_volume_done(self, name: str) -> bool:
        return name in self._done_volumes

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
        for p in passages:
            self._pending.append(p)
            if len(self._pending) >= _BATCH:
                self._flush()
        return True

    def _flush(self) -> None:
        if not self._pending:
            return
        if self.embed:
            vecs = self.embedder.encode([p.text for p in self._pending]).astype(np.float16)
            self._vec_fh.write(np.ascontiguousarray(vecs, dtype=np.float16).tobytes())
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

    def commit_volume(self, name: str) -> None:
        """Checkpoint: flush the buffer, fsync vectors, then commit the volume."""
        self._flush()
        # fsync vectors FIRST so committed passages always have vectors on disk.
        self._vec_fh.flush()
        os.fsync(self._vec_fh.fileno())
        self.con.execute("INSERT OR IGNORE INTO ingested_volumes (name) VALUES (?)", (name,))
        self.con.commit()
        self._done_volumes.add(name)

    def abort_volume(self) -> None:
        """Roll back the current uncommitted volume's partial work and reconcile
        the vectors file + in-memory state to the last committed checkpoint, so a
        failed volume leaves no orphan rows for the next volume to commit."""
        self.con.rollback()
        self._pending.clear()
        n_committed = self.con.execute("SELECT COUNT(*) FROM passages").fetchone()[0]
        self._next_vec_row = n_committed
        self._vec_fh.flush()
        self._vec_fh.close()
        with open(self.vectors_path, "r+b") as f:
            f.truncate(n_committed * self._bytes_per_vec)
        self._vec_fh = open(self.vectors_path, "ab")
        # Drop citations added in-memory for the rolled-back (uncommitted) cases.
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
