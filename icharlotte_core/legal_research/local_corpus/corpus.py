"""LocalCaseCorpus: offline retrieval mirroring CourtListenerClient's interface.

search_opinions runs FTS5 BM25 + exact-cosine semantic over memmap'd vectors,
fused by Reciprocal Rank Fusion. get_opinion_text / get_authority_signals /
lookup_by_citation serve the drafter + verifier without any network call.
"""
from __future__ import annotations

import logging
import re
import sqlite3
from typing import Any

import numpy as np

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.embedder import Embedder, OnnxEmbedder
from icharlotte_core.legal_research.models import CaseResult

logger = logging.getLogger(__name__)

_RRF_K = 60          # standard reciprocal-rank-fusion constant
_CANDIDATES = 100    # passages pulled per retrieval arm before fusion


def _fts_query(q: str) -> str:
    # OR the bare terms so partial overlaps still match; quote to neutralize FTS syntax.
    terms = re.findall(r"[A-Za-z0-9]+", q or "")
    return " OR ".join(f'"{t}"' for t in terms) if terms else '""'


class LocalCaseCorpus:
    def __init__(self, *, db_path: str, vectors_path: str, embedder: Embedder | None = None) -> None:
        self.db_path = db_path
        self.vectors_path = vectors_path
        self.embedder = embedder or OnnxEmbedder()
        self._con: sqlite3.Connection | None = None
        self._vectors: np.ndarray | None = None

    # ---- lazy resources -------------------------------------------------
    def _conn(self) -> sqlite3.Connection:
        if self._con is None:
            self._con = schema.connect(self.db_path)
        return self._con

    def _vecs(self) -> np.ndarray:
        if self._vectors is None:
            self._vectors = np.memmap(
                self.vectors_path, dtype=np.float16, mode="r"
            ).reshape(-1, self.embedder.dim)
        return self._vectors

    # ---- retrieval arms -------------------------------------------------
    def _bm25_case_ranking(self, query: str, limit: int) -> list[str]:
        con = self._conn()
        rows = con.execute(
            "SELECT p.case_uid AS uid, bm25(passages_fts) AS score "
            "FROM passages_fts JOIN passages p ON p.vec_row = passages_fts.rowid - 1 "
            "WHERE passages_fts MATCH ? ORDER BY score LIMIT ?",
            (_fts_query(query), limit),
        ).fetchall()
        seen, order = set(), []
        for r in rows:                       # bm25() ascending = best first
            if r["uid"] not in seen:
                seen.add(r["uid"]); order.append(r["uid"])
        return order

    def _semantic_case_ranking(self, query: str, limit: int) -> list[str]:
        vecs = self._vecs()
        if vecs.shape[0] == 0:
            return []
        from icharlotte_core.legal_research.local_corpus.embedder import cosine_topk
        qv = self.embedder.encode([query])[0].astype(np.float32)
        idx, _scores = cosine_topk(qv, vecs.astype(np.float32), k=limit)
        con = self._conn()
        order, seen = [], set()
        for vec_row in idx.tolist():
            row = con.execute("SELECT case_uid FROM passages WHERE vec_row=?", (int(vec_row),)).fetchone()
            if row and row["case_uid"] not in seen:
                seen.add(row["case_uid"]); order.append(row["case_uid"])
        return order

    @staticmethod
    def _rrf(*rankings: list[str]) -> list[str]:
        scores: dict[str, float] = {}
        for ranking in rankings:
            for rank, uid in enumerate(ranking):
                scores[uid] = scores.get(uid, 0.0) + 1.0 / (_RRF_K + rank + 1)
        return [uid for uid, _ in sorted(scores.items(), key=lambda kv: -kv[1])]

    # ---- public interface (mirrors CourtListenerClient) -----------------
    def search_opinions(self, query: str, *, semantic: bool = False,
                        max_results: int = 15, published_only: bool = True) -> list[CaseResult]:
        bm25 = self._bm25_case_ranking(query, _CANDIDATES)
        rankings = [bm25]
        if semantic:
            try:
                rankings.append(self._semantic_case_ranking(query, _CANDIDATES))
            except Exception:
                logger.warning("semantic ranking failed; BM25 only", exc_info=True)
        fused = self._rrf(*rankings)[:max_results]
        return [self._case_result(uid, query) for uid in fused]

    def _case_result(self, case_uid: str, query: str) -> CaseResult:
        con = self._conn()
        c = con.execute("SELECT * FROM cases WHERE case_uid=?", (case_uid,)).fetchone()
        snippet = ""
        if c:
            p = con.execute(
                "SELECT text FROM passages WHERE case_uid=? ORDER BY ordinal LIMIT 1", (case_uid,)
            ).fetchone()
            snippet = (p["text"][:400] if p else (c["full_text"] or "")[:400])
        return CaseResult(
            name=c["name"] if c else "", citation=c["citation"] if c else "",
            date=c["decision_date"] if c else "", court=c["court"] if c else "",
            snippet=snippet, url=c["url"] if c else "", cluster_id=case_uid,
        )

    def get_opinion_text(self, case_uid: str | int) -> str | None:
        row = self._conn().execute(
            "SELECT full_text FROM cases WHERE case_uid=?", (str(case_uid),)
        ).fetchone()
        return (row["full_text"] if row else None) or None

    def get_authority_signals(self, case_uid: str | int) -> dict[str, Any]:
        row = self._conn().execute(
            "SELECT citation_count, latest_citing_year FROM cases WHERE case_uid=?", (str(case_uid),)
        ).fetchone()
        if not row:
            return {"citation_count": None, "latest_citing_year": ""}
        return {"citation_count": row["citation_count"], "latest_citing_year": row["latest_citing_year"] or ""}

    def lookup_by_citation(self, citation: str) -> dict[str, Any] | None:
        norm = (citation or "").replace(" ", "").lower()
        if not norm:
            return None
        for row in self._conn().execute("SELECT * FROM cases"):
            if (row["citation"] or "").replace(" ", "").lower() == norm:
                return dict(row)
        return None
