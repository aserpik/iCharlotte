"""CourtListener REST API v4 client for California case law search."""

import logging
import re
from typing import Dict, List, Optional
from urllib.parse import urljoin

import requests

from icharlotte_core.legal_research.models import CaseResult

logger = logging.getLogger(__name__)

BASE_URL = "https://www.courtlistener.com/api/rest/v4"

# California courts on CourtListener
CA_COURTS = (
    "cal,calctapp,calag,calapp1st,calapp2nd,"
    "calapp3rd,calapp4th,calapp5th,calapp6th"
)

REQUEST_TIMEOUT = 30  # seconds


def opinion_url_for_cluster(cluster: dict) -> str:
    """Return a public CourtListener opinion URL for a cluster record."""
    absolute_url = (cluster or {}).get("absolute_url") or ""
    if not absolute_url:
        return ""
    if absolute_url.startswith("http://") or absolute_url.startswith("https://"):
        return absolute_url
    return urljoin("https://www.courtlistener.com/", absolute_url)


class CourtListenerClient:
    """Stateless client for the CourtListener REST API v4.

    Usage::

        client = CourtListenerClient(token="your-cl-api-token")
        results = client.search_opinions("premises liability duty")
    """

    def __init__(self, token: str) -> None:
        self.token = token

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _headers(self) -> Dict[str, str]:
        """Return headers required for authenticated API requests."""
        return {
            "Authorization": f"Token {self.token}",
            "Content-Type": "application/json",
        }

    def _auth_headers(self) -> Dict[str, str]:
        """Return auth headers without forcing a JSON request content type."""
        return {"Authorization": f"Token {self.token}"}

    def _parse_result(self, r: dict) -> CaseResult:
        """Convert a single CourtListener search result dict to a CaseResult."""
        # Citation: use the first entry if available
        citations = r.get("citation") or []
        citation_str = citations[0] if citations else ""

        # Strip HTML tags from snippet
        snippet_raw = r.get("snippet") or ""
        snippet_clean = re.sub(r"<[^>]+>", "", snippet_raw)

        full_url = opinion_url_for_cluster(r)

        return CaseResult(
            name=r.get("caseName") or "",
            citation=citation_str,
            date=r.get("dateFiled") or "",
            court=r.get("court") or "",
            snippet=snippet_clean,
            url=full_url,
            cluster_id=r.get("cluster_id"),
        )

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def lookup_citations(self, text: str) -> list[dict]:
        """Look up citations found in arbitrary text.

        Returns raw CourtListener citation-lookup records, or an empty list on
        blank input or API failure.
        """
        if not text or not text.strip():
            return []

        try:
            resp = requests.post(
                f"{BASE_URL}/citation-lookup/",
                headers=self._auth_headers(),
                data={"text": text},
                timeout=REQUEST_TIMEOUT,
            )
            resp.raise_for_status()
            data = resp.json()
            if isinstance(data, list):
                return data
            if isinstance(data, dict):
                results = data.get("results")
                return results if isinstance(results, list) else []
            return []
        except Exception:
            logger.warning("CourtListener citation lookup failed", exc_info=True)
            return []

    def get_cluster(self, cluster_id: int | str) -> dict | None:
        """Fetch a CourtListener opinion cluster by ID."""
        try:
            resp = requests.get(
                f"{BASE_URL}/clusters/{cluster_id}/",
                headers=self._headers(),
                timeout=REQUEST_TIMEOUT,
            )
            resp.raise_for_status()
            data = resp.json()
            return data if isinstance(data, dict) else None
        except Exception:
            logger.warning(
                "CourtListener cluster fetch failed for cluster %s",
                cluster_id,
                exc_info=True,
            )
            return None

    def search_opinions(
        self,
        query: str,
        jurisdiction: str = "cal",
        max_results: int = 15,
    ) -> List[CaseResult]:
        """Search CourtListener for California case opinions.

        Args:
            query: Free-text search query.
            jurisdiction: Unused (reserved); California courts are always used.
            max_results: Maximum number of results to return.

        Returns:
            List of CaseResult objects, or empty list on error.
        """
        params = {
            "q": query,
            "type": "o",
            "court": CA_COURTS,
            "order_by": "score desc",
            "page_size": max_results,
        }
        try:
            resp = requests.get(
                f"{BASE_URL}/search/",
                headers=self._headers(),
                params=params,
                timeout=REQUEST_TIMEOUT,
            )
            resp.raise_for_status()
            data = resp.json()
            return [self._parse_result(r) for r in data.get("results", [])]
        except Exception:
            logger.warning("CourtListener search failed for query: %s", query, exc_info=True)
            return []

    def get_citing_cases(
        self,
        cluster_id: int,
        max_results: int = 10,
    ) -> List[CaseResult]:
        """Find cases that cite a given opinion cluster.

        Args:
            cluster_id: CourtListener cluster ID of the cited opinion.
            max_results: Maximum number of results to return.

        Returns:
            List of CaseResult objects, or empty list on error.
        """
        params = {
            "q": f"cites:({cluster_id})",
            "type": "o",
            "court": CA_COURTS,
            "order_by": "dateFiled desc",
            "page_size": max_results,
        }
        try:
            resp = requests.get(
                f"{BASE_URL}/search/",
                headers=self._headers(),
                params=params,
                timeout=REQUEST_TIMEOUT,
            )
            resp.raise_for_status()
            data = resp.json()
            return [self._parse_result(r) for r in data.get("results", [])]
        except Exception:
            logger.warning(
                "CourtListener citing-cases lookup failed for cluster %s",
                cluster_id,
                exc_info=True,
            )
            return []

    def get_opinion_text(self, cluster_id: int) -> Optional[str]:
        """Fetch the full text of an opinion by cluster ID.

        Retrieves the cluster metadata first, then fetches the first
        sub-opinion's plain text.

        Args:
            cluster_id: CourtListener cluster ID.

        Returns:
            Plain text of the opinion, or None on error.
        """
        try:
            # Step 1: Get cluster to find sub_opinions
            cluster_resp = requests.get(
                f"{BASE_URL}/clusters/{cluster_id}/",
                headers=self._headers(),
                timeout=REQUEST_TIMEOUT,
            )
            cluster_resp.raise_for_status()
            cluster_data = cluster_resp.json()

            sub_opinions = cluster_data.get("sub_opinions") or []
            if not sub_opinions:
                logger.warning("No sub_opinions for cluster %s", cluster_id)
                return None

            # Step 2: Fetch the first sub-opinion
            # sub_opinions can be URLs (strings) or dicts with "id"
            first = sub_opinions[0]
            if isinstance(first, str):
                opinion_url = first
            elif isinstance(first, dict):
                opinion_id = first.get("id")
                opinion_url = f"{BASE_URL}/opinions/{opinion_id}/"
            else:
                return None

            opinion_resp = requests.get(
                opinion_url,
                headers=self._headers(),
                timeout=REQUEST_TIMEOUT,
            )
            opinion_resp.raise_for_status()
            opinion_data = opinion_resp.json()

            # Prefer plain_text, fall back to stripping HTML
            text = opinion_data.get("plain_text") or ""
            if not text:
                html = opinion_data.get("html_with_citations") or ""
                text = re.sub(r"<[^>]+>", "", html)

            return text if text else None

        except Exception:
            logger.warning(
                "CourtListener opinion fetch failed for cluster %s",
                cluster_id,
                exc_info=True,
            )
            return None
