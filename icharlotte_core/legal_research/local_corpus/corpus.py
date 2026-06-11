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
from dataclasses import dataclass
from typing import Any

import numpy as np

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.embedder import Embedder, OnnxEmbedder
from icharlotte_core.legal_research.models import CaseResult

logger = logging.getLogger(__name__)

_RRF_K = 60          # standard reciprocal-rank-fusion constant
_CANDIDATES = 100    # passages pulled per retrieval arm before fusion
_QUALITY_WEIGHT = 0.02
_FTS_OVERFETCH_FACTOR = 8
_SEMANTIC_OVERFETCH_FACTOR = 8
_SEMANTIC_MIN_CANDIDATES = 128
_SQLITE_IN_BATCH = 900
_CASE_COLUMNS = (
    "case_uid, source, name, name_abbreviation, citation, parallel_citations, "
    "court, decision_date, year, docket_number, url, full_text, "
    "citation_count, latest_citing_year, cites_to, published_status, citable"
)
_CA_CITATION_RE = re.compile(
    r"\b\d+\s+Cal\.?\s*(?:App\.?\s*)?(?:Rptr\.?\s*)?"
    r"\d*(?:st|nd|rd|th|d)?\s+\d+\b",
    re.I,
)

_METADATA_STOP = {
    "the", "and", "for", "with", "from", "that", "this", "motion", "party",
    "parties", "burden", "standard", "issue", "issues", "judgment", "summary",
    "of", "in", "to", "on", "by", "as", "at", "is", "are",
}
_STRICT_FTS_MIN_CASES = 8


@dataclass(frozen=True)
class _PassageHit:
    case_uid: str
    text: str
    passage_type: str
    parenthetical_id: str
    page_label: str


def _fts_terms(q: str, *, drop_stopwords: bool) -> list[str]:
    terms = re.findall(r"[A-Za-z0-9]+", q or "")
    if drop_stopwords:
        return [
            term
            for term in terms
            if len(term) > 2 and term.lower() not in _METADATA_STOP
        ]
    return terms


def _quote_fts_terms(terms: list[str], operator: str) -> str:
    if not terms:
        return '""'
    joiner = f" {operator} " if operator else " "
    return joiner.join(f'"{term}"' for term in terms)


def _fts_query(q: str) -> str:
    # Broad fallback: OR the bare terms so partial overlaps still match; quote
    # to neutralize FTS syntax.
    return _quote_fts_terms(_fts_terms(q, drop_stopwords=False), "OR")


def _strict_fts_query(q: str) -> str:
    # First-pass issue research should require all meaningful terms. This keeps
    # generic legal words such as "control" from swamping stronger doctrinal hits.
    return _quote_fts_terms(_fts_terms(q, drop_stopwords=True), "")


def _snippet_terms(query: str) -> list[str]:
    seen: set[str] = set()
    terms: list[str] = []
    for term in re.findall(r"[A-Za-z][A-Za-z0-9]+", query or ""):
        lowered = term.lower()
        if len(lowered) <= 2 or lowered in _METADATA_STOP or lowered in seen:
            continue
        seen.add(lowered)
        terms.append(lowered)
    return terms


def _best_snippet_window(text: str, query: str, *, max_chars: int = 650) -> str:
    clean = re.sub(r"\s+", " ", text or "").strip()
    if len(clean) <= max_chars:
        return clean
    terms = _snippet_terms(query)
    if not terms:
        return clean[:max_chars].rstrip() + " ..."

    lower_text = clean.lower()
    hits: list[tuple[int, int]] = []
    for term in terms:
        start = 0
        while True:
            idx = lower_text.find(term, start)
            if idx < 0:
                break
            hits.append((idx, len(term)))
            start = idx + len(term)
    if not hits:
        return clean[:max_chars].rstrip() + " ..."

    hit_positions = sorted(idx for idx, _length in hits)
    best_start = 0
    best_score = -1.0
    for idx, length in hits:
        start = max(0, idx - 220)
        end = min(len(clean), start + max_chars)
        score = sum(1 for pos in hit_positions if start <= pos < end)
        if start <= idx and idx + length <= end:
            score += 0.25
        if score > best_score:
            best_score = score
            best_start = start

    end = min(len(clean), best_start + max_chars)
    snippet = clean[best_start:end].strip()
    if best_start > 0:
        snippet = "... " + snippet
    if end < len(clean):
        snippet = snippet.rstrip() + " ..."
    return snippet


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
        self._passage_type_counts: dict[str, int] = {}

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
    def _passage_type_count(self, passage_type: str) -> int:
        if passage_type not in self._passage_type_counts:
            try:
                count = self._conn().execute(
                    "SELECT COUNT(*) FROM passages WHERE passage_type=?",
                    (passage_type,),
                ).fetchone()[0]
            except sqlite3.OperationalError:
                count = 0
            self._passage_type_counts[passage_type] = int(count or 0)
        return self._passage_type_counts[passage_type]

    def _fts_case_ranking(
        self,
        query: str,
        limit: int,
        *,
        passage_type: str,
    ) -> tuple[list[str], dict[str, _PassageHit]]:
        variants = []
        strict_query = _strict_fts_query(query)
        broad_query = _fts_query(query)
        if strict_query != '""':
            variants.append(strict_query)
        if broad_query not in variants:
            variants.append(broad_query)
        if not variants:
            return [], {}

        order: list[str] = []
        snippets: dict[str, _PassageHit] = {}
        seen: set[str] = set()
        for idx, match_query in enumerate(variants):
            arm_order, arm_snippets = self._fts_case_ranking_for_match(
                match_query,
                limit,
                passage_type=passage_type,
            )
            for uid in arm_order:
                if uid in seen:
                    continue
                seen.add(uid)
                order.append(uid)
                if uid in arm_snippets:
                    snippets[uid] = arm_snippets[uid]
                if len(order) >= limit:
                    break
            if idx == 0 and len(order) >= min(_STRICT_FTS_MIN_CASES, limit):
                break
            if len(order) >= limit:
                break
        return order, snippets

    def _fts_case_ranking_for_match(
        self,
        match_query: str,
        limit: int,
        *,
        passage_type: str,
    ) -> tuple[list[str], dict[str, _PassageHit]]:
        con = self._conn()
        if self._passage_type_count(passage_type) <= 0:
            return [], {}
        overfetch = max(limit, limit * _FTS_OVERFETCH_FACTOR)
        try:
            fts_rows = con.execute(
                "SELECT rowid, bm25(passages_fts) AS score "
                "FROM passages_fts WHERE passages_fts MATCH ? "
                "ORDER BY score LIMIT ?",
                (match_query, overfetch),
            ).fetchall()
        except sqlite3.OperationalError:
            logger.debug("%s FTS ranking failed", passage_type, exc_info=True)
            return [], {}
        if not fts_rows:
            return [], {}

        vec_rows = [int(row["rowid"]) - 1 for row in fts_rows]
        placeholders = ",".join("?" for _ in vec_rows)
        try:
            passage_rows = con.execute(
                "SELECT vec_row, case_uid, text, passage_type, parenthetical_id, page_label "
                f"FROM passages INDEXED BY idx_passages_vec WHERE vec_row IN ({placeholders}) "
                "AND passage_type=?",
                [*vec_rows, passage_type],
            ).fetchall()
        except sqlite3.OperationalError:
            logger.debug("%s passage lookup failed", passage_type, exc_info=True)
            return [], {}
        passages_by_vec = {
            int(row["vec_row"]): row
            for row in passage_rows
            if row["vec_row"] is not None
        }

        seen: set[str] = set()
        order: list[str] = []
        snippets: dict[str, _PassageHit] = {}
        for fts_row in fts_rows:             # bm25() ascending = best first
            passage = passages_by_vec.get(int(fts_row["rowid"]) - 1)
            if passage is None:
                continue
            uid = passage["case_uid"]
            if uid in seen:
                continue
            seen.add(uid)
            order.append(uid)
            snippets[uid] = _PassageHit(
                case_uid=uid,
                text=passage["text"] or "",
                passage_type=passage["passage_type"] or "",
                parenthetical_id=passage["parenthetical_id"] or "",
                page_label=passage["page_label"] or "",
            )
            if len(order) >= limit:
                break
        return order, snippets

    def _bm25_case_ranking(self, query: str, limit: int) -> tuple[list[str], dict[str, _PassageHit]]:
        return self._fts_case_ranking(query, limit, passage_type="opinion")

    def _parenthetical_case_ranking(self, query: str, limit: int) -> tuple[list[str], dict[str, _PassageHit]]:
        return self._fts_case_ranking(query, limit, passage_type="parenthetical")

    @staticmethod
    def _norm_citation(citation: str) -> str:
        return re.sub(r"[^a-z0-9]+", "", (citation or "").lower())

    @staticmethod
    def _citation_variants(citation: str) -> list[str]:
        text = re.sub(r"\s+", " ", (citation or "").strip())
        if not text:
            return []
        variants = {text}
        spaced = re.sub(r"\bCal\.?\s*App\.?\s*", "Cal. App. ", text, flags=re.I)
        spaced = re.sub(r"\bCal\.?\s*Rptr\.?\s*", "Cal. Rptr. ", spaced, flags=re.I)
        spaced = re.sub(r"\bCal\.?\s*(\d)", r"Cal. \1", spaced, flags=re.I)
        spaced = re.sub(r"\s+", " ", spaced).strip()
        variants.add(spaced)
        return sorted(variants)

    @staticmethod
    def _extract_citation(text: str) -> str:
        match = _CA_CITATION_RE.search(text or "")
        return match.group(0).strip() if match else ""

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
            name_matches = sum(1 for term in terms if term in hay_name)
            cite_matches = sum(1 for term in terms if term in hay_cite)
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
            exact_citation_match = any(
                self._norm_citation(cite) and self._norm_citation(cite) in norm_query
                for cite in row_cites
            )
            if exact_citation_match:
                score += 30.0

            term_count = max(len(terms), 1)
            name_match_ratio = name_matches / term_count
            strong_name_lookup = (
                name_matches == term_count
                or (name_matches >= 2 and name_match_ratio >= 0.75)
            )
            strong_citation_lookup = exact_citation_match or cite_matches > 0
            if score > 0 and (strong_name_lookup or strong_citation_lookup):
                score += self._quality_score(row)
                scored.append((score, row["case_uid"]))
        scored.sort(key=lambda item: (-item[0], item[1]))
        return [uid for _score, uid in scored[:limit]]

    def _opinion_cases_for_vec_rows(self, vec_rows: list[int]) -> dict[int, str]:
        if not vec_rows:
            return {}
        out: dict[int, str] = {}
        con = self._conn()
        for start in range(0, len(vec_rows), _SQLITE_IN_BATCH):
            batch = vec_rows[start : start + _SQLITE_IN_BATCH]
            placeholders = ",".join("?" for _ in batch)
            rows = con.execute(
                "SELECT vec_row, case_uid FROM passages INDEXED BY idx_passages_vec "
                f"WHERE vec_row IN ({placeholders}) AND passage_type='opinion'",
                batch,
            ).fetchall()
            for row in rows:
                if row["vec_row"] is not None:
                    out[int(row["vec_row"])] = row["case_uid"]
        return out

    def _semantic_case_ranking(self, query: str, limit: int) -> list[str]:
        vecs = self._vecs()
        if vecs.shape[0] == 0:
            return []
        qv = self.embedder.encode([query])[0].astype(np.float32)
        qn = np.linalg.norm(qv)
        if qn > 0:
            qv = qv / qn
        sims = vecs @ qv
        n_vectors = int(sims.shape[0])
        if n_vectors <= 0:
            return []

        top_k = min(
            n_vectors,
            max(limit * _SEMANTIC_OVERFETCH_FACTOR, _SEMANTIC_MIN_CANDIDATES, limit),
        )
        order: list[str] = []
        seen: set[str] = set()
        while True:
            idx = np.argpartition(-sims, top_k - 1)[:top_k]
            idx = idx[np.argsort(-sims[idx])]
            vec_rows = [int(value) for value in idx.tolist()]
            case_by_vec_row = self._opinion_cases_for_vec_rows(vec_rows)
            order.clear()
            seen.clear()
            for vec_row in vec_rows:
                case_uid = case_by_vec_row.get(vec_row)
                if not case_uid or case_uid in seen:
                    continue
                seen.add(case_uid)
                order.append(case_uid)
                if len(order) >= limit:
                    return order
            if len(order) >= limit or top_k >= n_vectors:
                return order
            top_k = min(n_vectors, max(top_k * 2, top_k + limit))
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
        citation = self._extract_citation(query)
        if citation:
            hit = self.lookup_by_citation(citation)
            if hit and (not published_only or self._is_citable_row(hit)):
                return [
                    self._case_result(
                        str(hit["case_uid"]),
                        query,
                        relevance_score=1.0,
                    )
                ][:max_results]

        metadata = self._metadata_case_ranking(query, _CANDIDATES, published_only=published_only)
        bm25, bm25_snippets = self._bm25_case_ranking(query, _CANDIDATES)
        parenthetical, parenthetical_snippets = self._parenthetical_case_ranking(query, _CANDIDATES)
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
            self._case_result(
                uid,
                query,
                relevance_score=relevance_scores.get(uid, 0.0),
                passage_hint=parenthetical_snippets.get(uid) or bm25_snippets.get(uid),
            )
            for uid in ordered
        ]

    def _fallback_passage_for_case(self, case_uid: str):
        return self._conn().execute(
            "SELECT text, passage_type, parenthetical_id, page_label FROM passages "
            "WHERE case_uid=? AND passage_type='opinion' ORDER BY ordinal LIMIT 1",
            (case_uid,),
        ).fetchone()

    def _case_result(
        self,
        case_uid: str,
        query: str,
        *,
        relevance_score: float = 0.0,
        passage_hint: _PassageHit | None = None,
    ) -> CaseResult:
        con = self._conn()
        c = con.execute("SELECT * FROM cases WHERE case_uid=?", (case_uid,)).fetchone()
        snippet = ""
        snippet_source = ""
        snippet_parenthetical_id = ""
        snippet_page_label = ""
        display_name = ""
        if c:
            if passage_hint is not None:
                snippet = _best_snippet_window(passage_hint.text, query)
                snippet_source = passage_hint.passage_type
                snippet_parenthetical_id = passage_hint.parenthetical_id
                snippet_page_label = passage_hint.page_label
            else:
                full_text = c["full_text"] or ""
                if full_text:
                    snippet = _best_snippet_window(full_text, query)
                    snippet_source = "opinion"
                else:
                    p = self._fallback_passage_for_case(case_uid)
                    snippet = _best_snippet_window(p["text"], query) if p else ""
                    if p:
                        snippet_source = p["passage_type"] or ""
                        snippet_parenthetical_id = p["parenthetical_id"] or ""
                        snippet_page_label = p["page_label"] or ""
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
            snippet_page_label=snippet_page_label,
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

        variants = self._citation_variants(citation)
        if variants:
            placeholders = ",".join("?" for _ in variants)
            row = self._conn().execute(
                f"SELECT {_CASE_COLUMNS} FROM cases "
                f"WHERE citation COLLATE NOCASE IN ({placeholders}) "
                "LIMIT 1",
                variants,
            ).fetchone()
            if row:
                return dict(row)

        row = self._conn().execute(
            f"SELECT {_CASE_COLUMNS} FROM cases "
            "WHERE lower(replace(replace(replace(citation, ' ', ''), '.', ''), char(160), '')) = ? "
            "LIMIT 1",
            (norm,),
        ).fetchone()
        if row:
            return dict(row)

        for row in self._conn().execute(f"SELECT {_CASE_COLUMNS} FROM cases WHERE parallel_citations IS NOT NULL AND parallel_citations<>''"):
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
