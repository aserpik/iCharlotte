"""LocalCaseCorpus: offline retrieval mirroring CourtListenerClient's interface.

search_opinions runs FTS5 BM25 + exact-cosine semantic over memmap'd vectors,
fused by Reciprocal Rank Fusion. get_opinion_text / get_authority_signals /
lookup_by_citation serve the drafter + verifier without any network call.
"""
from __future__ import annotations

import json
import logging
import os
import re
import sqlite3
import threading
from typing import Any

import numpy as np

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.embedder import Embedder, OnnxEmbedder
from icharlotte_core.legal_research.models import CaseResult

logger = logging.getLogger(__name__)

_RRF_K = 60          # standard reciprocal-rank-fusion constant
_CANDIDATES = 100    # passages pulled per retrieval arm before fusion
_QUALITY_WEIGHT = 0.02

_METADATA_STOP = {
    "the", "and", "for", "with", "from", "that", "this", "motion", "party",
    "parties", "burden", "standard", "issue", "issues", "judgment", "summary",
}


def _fts_query(q: str) -> str:
    # OR the bare terms so partial overlaps still match; quote to neutralize FTS syntax.
    terms = re.findall(r"[A-Za-z0-9]+", q or "")
    return " OR ".join(f'"{t}"' for t in terms) if terms else '""'


class LocalCaseCorpus:
    def __init__(self, *, db_path: str, vectors_path: str, embedder: Embedder | None = None) -> None:
        self.db_path = db_path
        self.vectors_path = vectors_path
        self.embedder = embedder or OnnxEmbedder()
        # SQLite connections are not shareable across threads. The opposition
        # research/verify steps fan arguments out over a ThreadPoolExecutor, so
        # each thread must open its own connection — a single shared one raises
        # "SQLite objects created in a thread can only be used in that same
        # thread" and silently sinks every search to zero results. WAL mode
        # (set in schema.connect) makes concurrent readers safe.
        self._local = threading.local()
        self._vectors: np.ndarray | None = None

    # ---- lazy resources -------------------------------------------------
    def _conn(self) -> sqlite3.Connection:
        con = getattr(self._local, "con", None)
        if con is None:
            con = schema.connect(self.db_path)
            schema.create_schema(con)
            self._local.con = con
        return con

    def _vecs(self) -> np.ndarray:
        if self._vectors is None:
            # np.memmap raises on a 0-byte file; an empty corpus has no vectors.
            try:
                if os.path.getsize(self.vectors_path) == 0:
                    self._vectors = np.zeros((0, self.embedder.dim), dtype=np.float16)
                else:
                    self._vectors = np.memmap(
                        self.vectors_path, dtype=np.float16, mode="r"
                    ).reshape(-1, self.embedder.dim)
            except OSError:
                self._vectors = np.zeros((0, self.embedder.dim), dtype=np.float16)
        return self._vectors

    # ---- retrieval arms -------------------------------------------------
    def _bm25_case_ranking(self, query: str, limit: int) -> list[str]:
        con = self._conn()
        try:
            rows = con.execute(
                "SELECT p.case_uid AS uid, bm25(passages_fts) AS score "
                "FROM passages_fts JOIN passages p ON p.vec_row = passages_fts.rowid - 1 "
                "WHERE p.passage_type='opinion' AND passages_fts MATCH ? "
                "ORDER BY score LIMIT ?",
                (_fts_query(query), limit),
            ).fetchall()
        except sqlite3.OperationalError:
            logger.debug("opinion-only BM25 failed; falling back to all passages", exc_info=True)
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

    def _parenthetical_case_ranking(self, query: str, limit: int) -> list[str]:
        con = self._conn()
        try:
            rows = con.execute(
                "SELECT p.case_uid AS uid, bm25(passages_fts) AS score "
                "FROM passages_fts JOIN passages p ON p.vec_row = passages_fts.rowid - 1 "
                "WHERE p.passage_type='parenthetical' AND passages_fts MATCH ? "
                "ORDER BY score LIMIT ?",
                (_fts_query(query), limit),
            ).fetchall()
        except sqlite3.OperationalError:
            logger.debug("parenthetical ranking skipped; schema lacks passage_type", exc_info=True)
            return []
        seen, order = set(), []
        for r in rows:
            if r["uid"] not in seen:
                seen.add(r["uid"]); order.append(r["uid"])
        return order

    @staticmethod
    def _norm_citation(citation: str) -> str:
        return re.sub(r"[^a-z0-9]+", "", (citation or "").lower())

    @staticmethod
    def _is_citable_row(row) -> bool:
        if row is None:
            return False
        citation = (row["citation"] or "").lower()
        status = (row["published_status"] or "").lower() if "published_status" in row.keys() else ""
        citable = row["citable"] if "citable" in row.keys() else 1
        if "unrep" in citation or "unpublished" in citation or "unpublished" in status:
            return False
        return bool(citable)

    def _case_allowed(self, case_uid: str, published_only: bool) -> bool:
        if not published_only:
            return True
        row = self._conn().execute(
            "SELECT citation, published_status, citable FROM cases WHERE case_uid=?",
            (case_uid,),
        ).fetchone()
        return self._is_citable_row(row)

    def _metadata_case_ranking(self, query: str, limit: int, *, published_only: bool) -> list[str]:
        terms = [
            t.lower()
            for t in re.findall(r"[A-Za-z0-9]+", query or "")
            if len(t) > 2 and t.lower() not in _METADATA_STOP
        ]
        norm_query = self._norm_citation(query)
        if not terms and not norm_query:
            return []

        clauses: list[str] = []
        params: list[str] = []
        for term in terms[:8]:
            like = f"%{term}%"
            clauses.extend([
                "lower(name) LIKE ?",
                "lower(name_abbreviation) LIKE ?",
                "lower(citation) LIKE ?",
                "lower(parallel_citations) LIKE ?",
            ])
            params.extend([like, like, like, like])
        if not clauses:
            return []

        sql = (
            "SELECT case_uid, name, name_abbreviation, citation, parallel_citations, "
            "court, decision_date, citation_count, latest_citing_year, published_status, citable "
            "FROM cases WHERE " + " OR ".join(clauses)
        )
        rows = self._conn().execute(sql, params).fetchall()
        scored: list[tuple[float, str]] = []
        for row in rows:
            if published_only and not self._is_citable_row(row):
                continue
            hay_name = f"{row['name'] or ''} {row['name_abbreviation'] or ''}".lower()
            hay_cite = f"{row['citation'] or ''} {row['parallel_citations'] or ''}".lower()
            score = 0.0
            for term in terms:
                if term in hay_name:
                    score += 5.0
                if term in hay_cite:
                    score += 8.0
            row_cites = [row["citation"] or ""]
            try:
                row_cites.extend(json.loads(row["parallel_citations"] or "[]"))
            except (TypeError, ValueError):
                pass
            if any(self._norm_citation(cite) and self._norm_citation(cite) in norm_query for cite in row_cites):
                score += 30.0
            if score > 0:
                score += self._quality_score(row)
                scored.append((score, row["case_uid"]))
        scored.sort(key=lambda item: (-item[0], item[1]))
        return [uid for _score, uid in scored[:limit]]

    def _semantic_case_ranking(self, query: str, limit: int) -> list[str]:
        vecs = self._vecs()
        if vecs.shape[0] == 0:
            return []
        from icharlotte_core.legal_research.local_corpus.embedder import cosine_topk
        qv = self.embedder.encode([query])[0].astype(np.float32)
        con = self._conn()
        rows = con.execute(
            "SELECT vec_row, case_uid FROM passages "
            "WHERE passage_type='opinion' AND vec_row IS NOT NULL ORDER BY vec_row"
        ).fetchall()
        opinion_rows: list[int] = []
        case_by_vec_row: dict[int, str] = {}
        for row in rows:
            vec_row = int(row["vec_row"])
            if 0 <= vec_row < vecs.shape[0]:
                opinion_rows.append(vec_row)
                case_by_vec_row[vec_row] = row["case_uid"]
        if not opinion_rows:
            return []
        matrix = vecs[np.asarray(opinion_rows, dtype=np.int64)].astype(np.float32)
        idx, _scores = cosine_topk(qv, matrix, k=limit)
        order, seen = [], set()
        for matrix_row in idx.tolist():
            vec_row = opinion_rows[int(matrix_row)]
            case_uid = case_by_vec_row[vec_row]
            if case_uid not in seen:
                seen.add(case_uid); order.append(case_uid)
        return order

    @staticmethod
    def _rrf_scores(*rankings: list[str]) -> dict[str, float]:
        scores: dict[str, float] = {}
        for ranking in rankings:
            for rank, uid in enumerate(ranking):
                scores[uid] = scores.get(uid, 0.0) + 1.0 / (_RRF_K + rank + 1)
        return scores

    @staticmethod
    def _rrf(*rankings: list[str]) -> list[str]:
        scores = LocalCaseCorpus._rrf_scores(*rankings)
        return [uid for uid, _ in sorted(scores.items(), key=lambda kv: -kv[1])]

    @staticmethod
    def _quality_score(row) -> float:
        court = (row["court"] or "").lower() if row else ""
        score = 0.0
        if court in {"cal.", "cal"} or "supreme" in court:
            score += 0.35
        elif "app" in court:
            score += 0.12
        date = (row["decision_date"] or "") if row else ""
        year = date[:4]
        if year.isdigit():
            y = int(year)
            score += max(0.0, min(0.35, (y - 1980) / 120.0))
        try:
            count = int(row["citation_count"] or 0) if row else 0
        except (TypeError, ValueError):
            count = 0
        if count > 0:
            import math
            score += min(0.25, math.log10(count + 1) / 8.0)
        latest = (row["latest_citing_year"] or "") if row else ""
        if latest.isdigit():
            score += max(0.0, min(0.10, (int(latest) - 2000) / 250.0))
        return score

    def _case_quality_score(self, case_uid: str) -> float:
        row = self._conn().execute(
            "SELECT court, decision_date, citation_count, latest_citing_year FROM cases WHERE case_uid=?",
            (case_uid,),
        ).fetchone()
        return self._quality_score(row)

    @staticmethod
    def _ordered_by_quality(scores: dict[str, float], quality) -> list[str]:
        return sorted(
            scores,
            key=lambda uid: (-(scores[uid] + (_QUALITY_WEIGHT * quality(uid))), uid),
        )

    # ---- public interface (mirrors CourtListenerClient) -----------------
    def search_opinions(self, query: str, *, semantic: bool = False,
                        max_results: int = 15, published_only: bool = True) -> list[CaseResult]:
        metadata = self._metadata_case_ranking(query, _CANDIDATES, published_only=published_only)
        bm25 = self._bm25_case_ranking(query, _CANDIDATES)
        parenthetical = self._parenthetical_case_ranking(query, _CANDIDATES)
        semantic_ranking: list[str] = []
        if semantic:
            try:
                semantic_ranking = self._semantic_case_ranking(query, _CANDIDATES)
            except Exception:
                logger.warning("semantic ranking failed; BM25 only", exc_info=True)

        direct_cases = set(metadata) | set(bm25)
        direct_scores = self._rrf_scores(metadata, bm25)
        direct_hits = [
            uid
            for uid in self._ordered_by_quality(direct_scores, self._case_quality_score)
            if uid in direct_cases
        ]

        parenthetical_recall = [
            uid for uid in parenthetical
            if uid not in direct_cases
        ]
        recall_cases = set(parenthetical_recall)
        semantic_scores = self._rrf_scores(semantic_ranking)
        semantic_only = [
            uid
            for uid in self._ordered_by_quality(semantic_scores, self._case_quality_score)
            if uid not in direct_cases and uid not in recall_cases
        ]

        ordered: list[str] = []
        seen: set[str] = set()
        for case_uid in direct_hits + parenthetical_recall + semantic_only:
            if case_uid in seen:
                continue
            seen.add(case_uid)
            if self._case_allowed(case_uid, published_only):
                ordered.append(case_uid)
            if len(ordered) >= max_results:
                break

        relevance_scores = self._rrf_scores(metadata, bm25, semantic_ranking, parenthetical)
        return [
            self._case_result(uid, query, relevance_score=relevance_scores.get(uid, 0.0))
            for uid in ordered
        ]

    def _best_passage_for_query(self, case_uid: str, query: str):
        con = self._conn()
        try:
            row = con.execute(
                "SELECT p.text, p.passage_type, p.parenthetical_id "
                "FROM passages_fts JOIN passages p ON p.vec_row = passages_fts.rowid - 1 "
                "WHERE p.case_uid=? AND passages_fts MATCH ? "
                "ORDER BY bm25(passages_fts) LIMIT 1",
                (case_uid, _fts_query(query)),
            ).fetchone()
            if row:
                return row
        except sqlite3.Error:
            logger.debug("best passage lookup failed for %s", case_uid, exc_info=True)
        return con.execute(
            "SELECT text, passage_type, parenthetical_id FROM passages "
            "WHERE case_uid=? AND passage_type='opinion' ORDER BY ordinal LIMIT 1",
            (case_uid,),
        ).fetchone()

    def _case_result(self, case_uid: str, query: str, *, relevance_score: float = 0.0) -> CaseResult:
        con = self._conn()
        c = con.execute("SELECT * FROM cases WHERE case_uid=?", (case_uid,)).fetchone()
        snippet = ""
        snippet_source = ""
        snippet_parenthetical_id = ""
        display_name = ""
        if c:
            p = self._best_passage_for_query(case_uid, query)
            snippet = (p["text"][:400] if p else (c["full_text"] or "")[:400])
            if p:
                snippet_source = p["passage_type"] or ""
                snippet_parenthetical_id = p["parenthetical_id"] or ""
            # Prefer the Bluebook short name (CAP name_abbreviation, e.g.
            # "Engalla v. Permanente Medical Group, Inc.") over the full party
            # caption stored in `name`. The caption is hundreds of chars long,
            # reads badly in a brief, and is unparseable by the citation parser
            # (so the output panel can't make it a selectable/verifiable cite).
            display_name = (c["name_abbreviation"] or c["name"] or "")
        return CaseResult(
            name=display_name, citation=c["citation"] if c else "",
            date=c["decision_date"] if c else "", court=c["court"] if c else "",
            snippet=snippet, url=c["url"] if c else "", cluster_id=case_uid,
            relevance_score=relevance_score,
            snippet_source=snippet_source,
            snippet_parenthetical_id=snippet_parenthetical_id,
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
        norm = self._norm_citation(citation)
        if not norm:
            return None
        for row in self._conn().execute("SELECT * FROM cases"):
            cites = [row["citation"] or ""]
            try:
                cites.extend(json.loads(row["parallel_citations"] or "[]"))
            except (TypeError, ValueError):
                pass
            if any(self._norm_citation(cite) == norm for cite in cites):
                return dict(row)
        return None

    def corpus_metadata(self) -> dict[str, Any]:
        con = self._conn()
        meta = schema.get_meta(con)
        source_counts = {
            str(row["source"] or ""): int(row["n"])
            for row in con.execute("SELECT source, COUNT(*) AS n FROM cases GROUP BY source")
        }
        max_date = con.execute(
            "SELECT MAX(decision_date) FROM cases WHERE decision_date IS NOT NULL AND decision_date<>''"
        ).fetchone()[0] or ""
        case_count = con.execute("SELECT COUNT(*) FROM cases").fetchone()[0]
        passage_count = con.execute("SELECT COUNT(*) FROM passages").fetchone()[0]
        vector_bytes = os.path.getsize(self.vectors_path) if os.path.exists(self.vectors_path) else 0
        return {
            **meta,
            "source_counts": source_counts,
            "max_decision_date": max_date,
            "case_count": int(case_count),
            "passage_count": int(passage_count),
            "vector_bytes": int(vector_bytes),
        }
