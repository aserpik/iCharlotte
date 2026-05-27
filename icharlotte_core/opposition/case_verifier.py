"""Verify case citations against CourtListener.

Wraps CourtListenerClient cite-lookup + opinion-fetch with on-disk caching of
opinion text. NOT_FOUND short-circuits when the citation isn't in CourtListener's
California reporter index. Otherwise runs the same verifier prompt used by the
statute path.
"""

from __future__ import annotations

import json
import logging
import os
from typing import Callable

from icharlotte_core.legal_research.sources.courtlistener import (
    CourtListenerClient,
    opinion_url_for_cluster,
)
from icharlotte_core.opposition.citation_parser import Citation
from icharlotte_core.opposition.models import CitationVerification
from icharlotte_core.opposition.statute_verifier import _parse_verdict_response
from icharlotte_core.prompt_manager import get_prompt

logger = logging.getLogger(__name__)

LLMCallback = Callable[[str, str], str]

_VALID_VERDICTS = {"SUPPORTED", "PARTIAL", "NOT_SUPPORTED"}


class CaseVerifier:
    def __init__(
        self,
        *,
        courtlistener_client: CourtListenerClient,
        llm_callback: LLMCallback,
        cache_dir: str,
    ) -> None:
        self.cl = courtlistener_client
        self.llm = llm_callback
        self.cache_dir = cache_dir
        os.makedirs(self.cache_dir, exist_ok=True)

    def _cache_path(self, cluster_id: str | int) -> str:
        return os.path.join(self.cache_dir, f"{cluster_id}.json")

    def _load_cached_text(self, cluster_id: str | int) -> str | None:
        path = self._cache_path(cluster_id)
        if not os.path.exists(path):
            return None
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data.get("text") or None
        except (OSError, ValueError, json.JSONDecodeError):
            logger.warning("Could not read opinion cache: %s", path, exc_info=True)
            return None

    def _save_cached_text(self, cluster_id: str | int, text: str) -> None:
        path = self._cache_path(cluster_id)
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump({"cluster_id": str(cluster_id), "text": text}, f)
        except OSError:
            logger.warning("Could not write opinion cache: %s", path, exc_info=True)

    def verify(self, citation: Citation) -> CitationVerification:
        cv = CitationVerification(
            citation_text=citation.raw_text,
            normalized_citation=citation.normalized,
            kind="case",
            case_name=citation.case_name,
            date=citation.year,
            proposition=citation.proposition,
            body_offset=citation.body_offset,
        )

        # 1. CourtListener cite-lookup
        try:
            records = self.cl.lookup_citations(citation.raw_text) or []
        except Exception:
            logger.warning("CourtListener cite-lookup failed", exc_info=True)
            records = []

        cluster = _first_valid_cluster(records)
        if not cluster:
            cv.verdict = "NOT_FOUND"
            cv.note = (
                "This citation does not appear in CourtListener's California "
                "reporter index; it may be invented, mis-cited, or unpublished."
            )
            return cv

        cluster_id = str(
            cluster.get("id")
            or cluster.get("cluster_id")
            or cluster.get("clusterId")
            or ""
        ).strip()
        cv.cluster_id = cluster_id
        cv.opinion_url = opinion_url_for_cluster(cluster)
        if cluster.get("case_name") and not cv.case_name:
            cv.case_name = cluster["case_name"]

        # 2. Opinion text — cache check first
        text = self._load_cached_text(cluster_id) if cluster_id else None
        if text is None and cluster_id:
            try:
                text = self.cl.get_opinion_text(int(cluster_id))
            except (TypeError, ValueError):
                text = None
            except Exception:
                logger.warning("CourtListener opinion fetch failed", exc_info=True)
                text = None
            if text:
                self._save_cached_text(cluster_id, text)

        if not text:
            cv.verdict = "UNVERIFIED"
            cv.note = (
                "CourtListener returned a cluster but no opinion text was "
                "available; verify manually."
            )
            return cv

        # 3. LLM verdict
        prompt_template = get_prompt("oppose_motion", "verify_citation") or ""
        if not prompt_template:
            cv.verdict = "UNVERIFIED"
            cv.note = "Verifier prompt not configured."
            return cv

        user_prompt = prompt_template.format(
            proposition=citation.proposition or "(no proposition extracted)",
            citation_text=citation.raw_text,
            authority_text=text,
        )
        try:
            response = self.llm("", user_prompt) or ""
        except Exception:
            logger.warning("Case verifier LLM call failed", exc_info=True)
            cv.verdict = "UNVERIFIED"
            cv.note = "Verifier LLM call failed; verify manually."
            return cv

        verdict, evidence, note = _parse_verdict_response(response)
        if verdict not in _VALID_VERDICTS:
            cv.verdict = "UNVERIFIED"
            cv.note = "Could not parse verifier response; verify manually."
            return cv

        cv.verdict = verdict
        cv.evidence = evidence
        cv.note = note
        return cv


def _first_valid_cluster(records: list[dict]) -> dict:
    for record in records or []:
        if not isinstance(record, dict):
            continue
        status = record.get("status", record.get("status_code"))
        try:
            status_int = int(status) if status is not None else None
        except (TypeError, ValueError):
            status_int = None
        if status_int is not None and status_int != 200:
            continue
        clusters = record.get("clusters") or record.get("cluster") or []
        if isinstance(clusters, dict):
            return clusters
        if isinstance(clusters, list) and clusters and isinstance(clusters[0], dict):
            return clusters[0]
    return {}
