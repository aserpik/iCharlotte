# icharlotte_core/firm_briefs/index.py
"""SQLite + float16 vector sidecar index over the firm brief library.

Connections are thread-local (WAL); the vector sidecar is an append-only
float16 file, one row per brief, addressed by briefs.vec_row.

FTS5 external-content table note:
  citations_fts is an external-content table (content='citations',
  content_rowid='id').  Inserts MUST explicitly pass (rowid, proposition).
  When citation rows are deleted, the corresponding FTS rows must also be
  removed via the 'delete' command ('INSERT INTO citations_fts(citations_fts,
  rowid, proposition) VALUES(''delete'', old_id, old_prop)') or the old
  propositions linger in MATCH results.
"""
from __future__ import annotations

import os
import sqlite3
import threading
from typing import Any, List, Optional

import numpy as np

from .embedding import EMBED_DIM

_SCHEMA = """
CREATE TABLE IF NOT EXISTS briefs(
  id INTEGER PRIMARY KEY,
  path TEXT UNIQUE, content_hash TEXT,
  motion_type TEXT, side TEXT,
  heading TEXT, profile TEXT,
  vec_row INTEGER DEFAULT -1,
  char_len INTEGER DEFAULT 0, ocr_ratio REAL DEFAULT 0.0,
  ingested_at TEXT DEFAULT (datetime('now')),
  status TEXT DEFAULT 'ok'
);
CREATE TABLE IF NOT EXISTS citations(
  id INTEGER PRIMARY KEY,
  brief_id INTEGER REFERENCES briefs(id) ON DELETE CASCADE,
  case_name TEXT, reporter_cite TEXT, year TEXT, norm_cite TEXT,
  proposition TEXT, quoted_passage TEXT
);
CREATE INDEX IF NOT EXISTS ix_cit_norm ON citations(norm_cite);
CREATE INDEX IF NOT EXISTS ix_cit_brief ON citations(brief_id);
CREATE VIRTUAL TABLE IF NOT EXISTS citations_fts USING fts5(
  proposition, content='citations', content_rowid='id'
);
"""


class FirmBriefIndex:
    def __init__(self, *, db_path: str, vectors_path: str, embedder=None) -> None:
        self.db_path = db_path
        self.vectors_path = vectors_path
        self.embedder = embedder
        self._local = threading.local()
        os.makedirs(os.path.dirname(db_path) or ".", exist_ok=True)

    # -- connection -------------------------------------------------------
    def _conn(self) -> sqlite3.Connection:
        con = getattr(self._local, "con", None)
        if con is None:
            con = sqlite3.connect(self.db_path)
            con.row_factory = sqlite3.Row
            con.execute("PRAGMA journal_mode=WAL")
            con.execute("PRAGMA foreign_keys=ON")
            self._local.con = con
        return con

    def create_schema(self) -> None:
        con = self._conn()
        con.executescript(_SCHEMA)
        con.commit()

    # -- vector sidecar ---------------------------------------------------
    def _append_vector(self, vec: np.ndarray) -> int:
        v = np.asarray(vec, dtype=np.float16).reshape(EMBED_DIM)
        row = 0
        if os.path.exists(self.vectors_path):
            row = os.path.getsize(self.vectors_path) // (EMBED_DIM * 2)
        with open(self.vectors_path, "ab") as f:
            f.write(v.tobytes())
            f.flush()
            os.fsync(f.fileno())
        return int(row)

    def load_vectors(self) -> np.ndarray:
        if not os.path.exists(self.vectors_path) or os.path.getsize(self.vectors_path) == 0:
            return np.zeros((0, EMBED_DIM), dtype=np.float16)
        return np.memmap(self.vectors_path, dtype=np.float16, mode="r").reshape(-1, EMBED_DIM)

    # -- writes -----------------------------------------------------------
    def _delete_citations_for_brief(self, con: sqlite3.Connection, bid: int) -> None:
        """Delete citations rows AND remove them from the FTS index.

        For an external-content FTS5 table the FTS shadow tables are NOT
        updated automatically when the content table rows are deleted.  We
        must issue an explicit 'delete' command for each row so the old
        propositions stop appearing in MATCH results.
        """
        old_rows = con.execute(
            "SELECT id, proposition FROM citations WHERE brief_id=?", (bid,)
        ).fetchall()
        for row in old_rows:
            con.execute(
                "INSERT INTO citations_fts(citations_fts, rowid, proposition) "
                "VALUES('delete', ?, ?)",
                (int(row["id"]), row["proposition"] or ""),
            )
        con.execute("DELETE FROM citations WHERE brief_id=?", (bid,))

    def upsert_brief(self, *, path: str, content_hash: str, motion_type: str, side: str,
                     heading: str, profile: str, profile_vec, char_len: int, ocr_ratio: float,
                     cites: List[Any]) -> int:
        con = self._conn()
        # Append a fresh vector row (sidecar is append-only; old rows orphaned
        # until --compact). fsync sidecar BEFORE the DB commit (crash ordering).
        vec_row = self._append_vector(profile_vec)
        existing = con.execute("SELECT id FROM briefs WHERE path=?", (path,)).fetchone()
        if existing:
            bid = int(existing["id"])
            self._delete_citations_for_brief(con, bid)
            con.execute(
                "UPDATE briefs SET content_hash=?, motion_type=?, side=?, heading=?, "
                "profile=?, vec_row=?, char_len=?, ocr_ratio=?, status='ok', "
                "ingested_at=datetime('now') WHERE id=?",
                (content_hash, motion_type, side, heading, profile, vec_row,
                 char_len, ocr_ratio, bid),
            )
        else:
            cur = con.execute(
                "INSERT INTO briefs(path, content_hash, motion_type, side, heading, "
                "profile, vec_row, char_len, ocr_ratio) VALUES(?,?,?,?,?,?,?,?,?)",
                (path, content_hash, motion_type, side, heading, profile, vec_row,
                 char_len, ocr_ratio),
            )
            bid = int(cur.lastrowid)
        for c in cites:
            cur = con.execute(
                "INSERT INTO citations(brief_id, case_name, reporter_cite, year, "
                "norm_cite, proposition, quoted_passage) VALUES(?,?,?,?,?,?,?)",
                (bid, getattr(c, "case_name", ""), getattr(c, "reporter_citation", ""),
                 getattr(c, "year", ""), getattr(c, "norm_cite", ""),
                 getattr(c, "proposition", ""), getattr(c, "quoted_passage", "")),
            )
            con.execute("INSERT INTO citations_fts(rowid, proposition) VALUES(?,?)",
                        (cur.lastrowid, getattr(c, "proposition", "")))
        con.commit()
        return bid

    def mark_stale(self, path: str) -> None:
        con = self._conn()
        row = con.execute("SELECT id FROM briefs WHERE path=?", (path,)).fetchone()
        if row:
            self._delete_citations_for_brief(con, int(row["id"]))
            con.execute("UPDATE briefs SET status='stale' WHERE id=?", (int(row["id"]),))
            con.commit()

    # -- reads ------------------------------------------------------------
    def has_current(self, path: str, content_hash: str) -> bool:
        con = self._conn()
        row = con.execute(
            "SELECT 1 FROM briefs WHERE path=? AND content_hash=? AND status='ok'",
            (path, content_hash),
        ).fetchone()
        return row is not None

    def authority_candidates(self, proposition: str, *, motion_type: str,
                             limit: int = 8) -> List[dict]:
        con = self._conn()
        q = _fts_query(proposition)
        if not q:
            return []
        try:
            rows = con.execute(
                "SELECT c.case_name, c.reporter_cite, c.year, c.norm_cite, c.proposition, "
                "c.quoted_passage, b.path AS source_brief "
                "FROM citations_fts f "
                "JOIN citations c ON c.id = f.rowid "
                "JOIN briefs b ON b.id = c.brief_id "
                "WHERE citations_fts MATCH ? AND b.status='ok' AND b.motion_type=? "
                "ORDER BY bm25(citations_fts) LIMIT ?",
                (q, motion_type, limit),
            ).fetchall()
        except sqlite3.OperationalError:
            # bm25 not available — fall back to rowid ordering
            rows = con.execute(
                "SELECT c.case_name, c.reporter_cite, c.year, c.norm_cite, c.proposition, "
                "c.quoted_passage, b.path AS source_brief "
                "FROM citations_fts f "
                "JOIN citations c ON c.id = f.rowid "
                "JOIN briefs b ON b.id = c.brief_id "
                "WHERE citations_fts MATCH ? AND b.status='ok' AND b.motion_type=? "
                "ORDER BY f.rowid LIMIT ?",
                (q, motion_type, limit),
            ).fetchall()
        return [dict(r) for r in rows]

    def stats(self) -> dict:
        con = self._conn()
        b = con.execute("SELECT COUNT(*) n FROM briefs WHERE status='ok'").fetchone()["n"]
        c = con.execute("SELECT COUNT(*) n FROM citations").fetchone()["n"]
        return {"briefs": b, "citations": c}


def _fts_query(text: str) -> str:
    import re
    toks = re.findall(r"[A-Za-z0-9]+", (text or "").lower())
    toks = [t for t in toks if len(t) > 1][:12]
    return " OR ".join(toks)
