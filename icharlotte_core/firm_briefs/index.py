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
  status TEXT DEFAULT 'ok',
  full_text TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS citations(
  id INTEGER PRIMARY KEY,
  brief_id INTEGER REFERENCES briefs(id) ON DELETE CASCADE,
  case_name TEXT, reporter_cite TEXT, year TEXT, norm_cite TEXT,
  proposition TEXT, quoted_passage TEXT,
  prop_vec_row INTEGER DEFAULT -1
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
        self.prop_vectors_path = vectors_path + ".prop"
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
        # Migrations: add columns that may not exist in pre-existing tables.
        try:
            con.execute("ALTER TABLE briefs ADD COLUMN full_text TEXT DEFAULT ''")
        except sqlite3.OperationalError:
            pass  # column already exists
        try:
            con.execute("ALTER TABLE citations ADD COLUMN prop_vec_row INTEGER DEFAULT -1")
        except sqlite3.OperationalError:
            pass  # column already exists
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

    # -- prop vector sidecar (per-citation) --------------------------------
    def _append_prop_vector(self, vec: np.ndarray) -> int:
        v = np.asarray(vec, dtype=np.float16).reshape(EMBED_DIM)
        row = 0
        if os.path.exists(self.prop_vectors_path):
            row = os.path.getsize(self.prop_vectors_path) // (EMBED_DIM * 2)
        with open(self.prop_vectors_path, "ab") as f:
            f.write(v.tobytes())
            f.flush()
            os.fsync(f.fileno())
        return int(row)

    def load_prop_vectors(self) -> np.ndarray:
        if not os.path.exists(self.prop_vectors_path) or os.path.getsize(self.prop_vectors_path) == 0:
            return np.zeros((0, EMBED_DIM), dtype=np.float16)
        return np.memmap(self.prop_vectors_path, dtype=np.float16, mode="r").reshape(-1, EMBED_DIM)

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
                     cites: List[Any], full_text: str = "",
                     prop_vecs: Optional[List] = None) -> int:
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
                "profile=?, vec_row=?, char_len=?, ocr_ratio=?, full_text=?, status='ok', "
                "ingested_at=datetime('now') WHERE id=?",
                (content_hash, motion_type, side, heading, profile, vec_row,
                 char_len, ocr_ratio, full_text or "", bid),
            )
        else:
            cur = con.execute(
                "INSERT INTO briefs(path, content_hash, motion_type, side, heading, "
                "profile, vec_row, char_len, ocr_ratio, full_text) VALUES(?,?,?,?,?,?,?,?,?,?)",
                (path, content_hash, motion_type, side, heading, profile, vec_row,
                 char_len, ocr_ratio, full_text or ""),
            )
            bid = int(cur.lastrowid)
        # Determine whether we have aligned prop vectors for citations.
        _has_prop_vecs = (prop_vecs is not None and len(prop_vecs) == len(cites))
        for i, c in enumerate(cites):
            pvr = -1
            if _has_prop_vecs:
                pvr = self._append_prop_vector(prop_vecs[i])
            cur = con.execute(
                "INSERT INTO citations(brief_id, case_name, reporter_cite, year, "
                "norm_cite, proposition, quoted_passage, prop_vec_row) VALUES(?,?,?,?,?,?,?,?)",
                (bid, getattr(c, "case_name", ""), getattr(c, "reporter_citation", ""),
                 getattr(c, "year", ""), getattr(c, "norm_cite", ""),
                 getattr(c, "proposition", ""), getattr(c, "quoted_passage", ""), pvr),
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

    def get_full_text(self, path: str) -> str:
        row = self._conn().execute(
            "SELECT full_text FROM briefs WHERE path=?", (path,)).fetchone()
        return (row["full_text"] if row else "") or ""

    def get_full_text_by_id(self, brief_id: int) -> str:
        row = self._conn().execute(
            "SELECT full_text FROM briefs WHERE id=?", (brief_id,)).fetchone()
        return (row["full_text"] if row else "") or ""

    # -- reads ------------------------------------------------------------
    def has_current(self, path: str, content_hash: str) -> bool:
        con = self._conn()
        row = con.execute(
            "SELECT 1 FROM briefs WHERE path=? AND content_hash=? AND status='ok'",
            (path, content_hash),
        ).fetchone()
        return row is not None

    def authority_candidates(self, proposition: str, *, motion_type: str,
                             limit: int = 8, query_vec=None) -> List[dict]:
        con = self._conn()
        q = _fts_query(proposition)
        if not q:
            return []
        try:
            rows = con.execute(
                "SELECT c.case_name, c.reporter_cite, c.year, c.norm_cite, c.proposition, "
                "c.quoted_passage, b.path AS source_brief, c.prop_vec_row "
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
                "c.quoted_passage, b.path AS source_brief, c.prop_vec_row "
                "FROM citations_fts f "
                "JOIN citations c ON c.id = f.rowid "
                "JOIN briefs b ON b.id = c.brief_id "
                "WHERE citations_fts MATCH ? AND b.status='ok' AND b.motion_type=? "
                "ORDER BY f.rowid LIMIT ?",
                (q, motion_type, limit),
            ).fetchall()
        # When query_vec is None, return in FTS order (byte-identical to old behavior).
        if query_vec is None:
            return [
                {k: r[k] for k in r.keys() if k != "prop_vec_row"}
                for r in rows
            ]
        # Semantic rerank: reciprocal-rank fusion of FTS position and cosine(query_vec, prop_vec).
        pvecs = self.load_prop_vectors()
        qv = np.asarray(query_vec, dtype=np.float32).reshape(-1)
        qvn = float(np.linalg.norm(qv)) or 1.0
        scored: List[tuple] = []
        for fts_pos, r in enumerate(rows):
            pvr = r["prop_vec_row"] if "prop_vec_row" in r.keys() else -1
            fts_score = 1.0 / (60 + fts_pos)
            sem_score = 0.0
            if pvr >= 0 and pvecs.shape[0] > pvr:
                pv = np.asarray(pvecs[pvr], dtype=np.float32)
                pvn = float(np.linalg.norm(pv)) or 1.0
                sem_score = float(np.dot(qv, pv) / (qvn * pvn))
            scored.append((fts_score + sem_score, r))
        scored.sort(key=lambda x: x[0], reverse=True)
        return [
            {k: r[k] for k in r.keys() if k != "prop_vec_row"}
            for _, r in scored
        ]

    def style_candidates(self, query_vec, *, motion_type: str, side: str,
                         k: int = 3) -> list[dict]:
        """Top-k briefs of this motion_type AND side by cosine similarity of the
        stored profile vector to query_vec, with a quality penalty (short/noisy
        briefs make poor style models) and version dedup."""
        import os
        import re as _re
        con = self._conn()
        rows = con.execute(
            "SELECT id, path, vec_row, char_len, ocr_ratio FROM briefs "
            "WHERE status='ok' AND motion_type=? AND side=? AND vec_row>=0",
            (motion_type, side),
        ).fetchall()
        if not rows:
            return []
        vecs = self.load_vectors()
        if getattr(vecs, "shape", (0,))[0] == 0:
            return []
        q = np.asarray(query_vec, dtype=np.float32).reshape(-1)
        qn = float(np.linalg.norm(q)) or 1.0
        scored: list[tuple[float, dict]] = []
        for r in rows:
            vr = int(r["vec_row"])
            if vr < 0 or vr >= vecs.shape[0]:
                continue
            v = np.asarray(vecs[vr], dtype=np.float32)
            vn = float(np.linalg.norm(v)) or 1.0
            cos = float(np.dot(q, v) / (qn * vn))
            penalty = 0.0
            if (r["char_len"] or 0) < 1500:
                penalty += 0.15
            if (r["ocr_ratio"] or 0.0) > 0.10:
                penalty += 0.15
            scored.append((cos - penalty, dict(r)))
        scored.sort(key=lambda x: x[0], reverse=True)
        out: list[dict] = []
        seen: set[str] = set()
        for score, r in scored:
            base = _re.sub(r"_\d+$", "", os.path.splitext(os.path.basename(r["path"]))[0])
            if base in seen:
                continue
            seen.add(base)
            r["score"] = score
            out.append(r)
            if len(out) >= k:
                break
        return out

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
