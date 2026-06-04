"""Turn harvested firm cites into preferred research candidates.

Resolution: local corpus opinion text -> live CourtListener -> unverified flag.
Returned dicts match argument_research._run's candidate shape plus provenance.
"""
from __future__ import annotations

import logging
from typing import Any, List, Optional

logger = logging.getLogger(__name__)


class FirmAuthorityProvider:
    def __init__(self, index, corpus, cl_client: Optional[Any] = None,
                 embedder=None) -> None:
        self.index = index
        self.corpus = corpus
        self.cl_client = cl_client
        self._embedder = embedder

    def _get_embedder(self):
        if self._embedder is not None:
            return self._embedder
        from .embedding import get_embedder
        return get_embedder()

    def candidates_for(self, proposition: str, *, motion_type: str, side: str,
                       limit: int = 6) -> List[dict]:
        # Embed proposition for semantic rerank; fall back to FTS-only on failure.
        query_vec = None
        try:
            emb = self._get_embedder()
            query_vec = emb.encode([proposition])[0]
        except Exception:
            logger.warning("prop embed failed, falling back to FTS-only", exc_info=True)
            query_vec = None
        try:
            rows = self.index.authority_candidates(
                proposition, motion_type=motion_type, limit=limit, query_vec=query_vec)
        except TypeError:
            # Index may not support query_vec (e.g., legacy FakeIndex in tests); fall back.
            try:
                rows = self.index.authority_candidates(
                    proposition, motion_type=motion_type, limit=limit)
            except Exception:
                logger.warning("firm authority lookup failed", exc_info=True)
                return []
        except Exception:
            logger.warning("firm authority lookup failed", exc_info=True)
            return []
        out: List[dict] = []
        for r in rows:
            out.append(self._resolve(r))
        return out

    def _resolve(self, r: dict) -> dict:
        cite = r.get("reporter_cite", "")
        base = {
            "case_name": r.get("case_name", ""),
            "citation": cite,
            "year": r.get("year", ""),
            "opinion_url": "",
            "source": "firm",
            "source_brief": r.get("source_brief", ""),
            "passage": r.get("quoted_passage", ""),
            "proposition": r.get("proposition", ""),
        }
        # 1) local corpus
        try:
            hit = self.corpus.lookup_by_citation(cite) if self.corpus else None
        except Exception:
            hit = None
        if hit:
            uid = str(hit.get("case_uid") or hit.get("cluster_id") or "")
            text = ""
            try:
                text = self.corpus.get_opinion_text(uid) or ""
            except Exception:
                text = ""
            if text:
                base.update({"cluster_id": uid, "text": text, "verification": "local"})
                return base
        # 2) live CourtListener fallback
        if self.cl_client is not None:
            try:
                hit2 = self.cl_client.lookup_by_citation(cite)
                uid2 = str((hit2 or {}).get("case_uid") or "")
                text2 = self.cl_client.get_opinion_text(cite) or ""
            except Exception:
                uid2, text2 = "", ""
            if text2:
                base.update({"cluster_id": uid2 or ("cl:" + cite),
                             "text": text2, "verification": "courtlistener"})
                return base
        # 3) unverified — keep, flagged; no opinion text
        base.update({"cluster_id": "firm:" + (r.get("norm_cite", "") or cite),
                     "text": "", "verification": "unverified_firm"})
        return base
