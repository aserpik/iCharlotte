"""CourtListener REST API v4 client for California case law search."""

import logging
import re
import time
from typing import Callable, Dict, List, Optional
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

# Retry policy for transient failures (HTTP 429 throttling, 5xx, timeouts).
# During a bulk verification run (many citations, parallel workers) the API
# will rate-limit some requests; without retry a transient 429 silently
# downgrades a citation to UNVERIFIED — which can mask a would-be
# NOT_SUPPORTED flag. Exponential backoff with a respected Retry-After.
_MAX_RETRIES = 3
_RETRY_BASE_DELAY = 1.0   # seconds; doubled each attempt
_MAX_RETRY_WAIT = 20.0    # cap any single sleep


def _send_with_retry(send_fn: Callable[[], "requests.Response"]) -> "requests.Response":
    """Call ``send_fn`` (a zero-arg request) with retry on 429/5xx/timeout.

    Returns the final Response. Retries are exhausted on persistent failure;
    the last response (or raised exception) propagates to the caller, which
    already handles non-2xx via ``raise_for_status`` + try/except.
    """
    delay = _RETRY_BASE_DELAY
    last_resp: "requests.Response | None" = None
    for attempt in range(_MAX_RETRIES + 1):
        try:
            resp = send_fn()
        except requests.RequestException:
            if attempt >= _MAX_RETRIES:
                raise
            time.sleep(min(delay, _MAX_RETRY_WAIT))
            delay *= 2
            continue

        last_resp = resp
        status = getattr(resp, "status_code", None)
        transient = status == 429 or (isinstance(status, int) and 500 <= status < 600)
        if transient and attempt < _MAX_RETRIES:
            wait = delay
            retry_after = None
            try:
                retry_after = resp.headers.get("Retry-After")
            except Exception:
                retry_after = None
            if retry_after:
                try:
                    wait = float(retry_after)
                except (TypeError, ValueError):
                    wait = delay
            logger.warning(
                "CourtListener returned %s; retrying in %.1fs (attempt %d/%d)",
                status, min(wait, _MAX_RETRY_WAIT), attempt + 1, _MAX_RETRIES,
            )
            time.sleep(min(wait, _MAX_RETRY_WAIT))
            delay *= 2
            continue
        return resp
    return last_resp  # pragma: no cover - loop always returns earlier


def opinion_url_for_cluster(cluster: dict) -> str:
    """Return a public CourtListener opinion URL for a cluster record."""
    absolute_url = (cluster or {}).get("absolute_url") or ""
    if not absolute_url:
        return ""
    if absolute_url.startswith("http://") or absolute_url.startswith("https://"):
        return absolute_url
    return urljoin("https://www.courtlistener.com/", absolute_url)


def preferred_california_citation(citations: list[str]) -> str:
    """Prefer a California official/appellate reporter citation when present."""
    if not citations:
        return ""
    for citation in citations:
        if re.search(r"\bCal\.\s*(?:App\.\s*)?(?:2d|3d|4th|5th|6th)?\s+\d+", citation):
            return citation
    return citations[0]


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
        # Citation: prefer a California reporter entry when available.
        citations = r.get("citation") or []
        citation_str = preferred_california_citation(citations)

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
            resp = _send_with_retry(
                lambda: requests.post(
                    f"{BASE_URL}/citation-lookup/",
                    headers=self._auth_headers(),
                    data={"text": text},
                    timeout=REQUEST_TIMEOUT,
                )
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
            resp = _send_with_retry(
                lambda: requests.get(
                    f"{BASE_URL}/clusters/{cluster_id}/",
                    headers=self._headers(),
                    timeout=REQUEST_TIMEOUT,
                )
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

        Iterates ALL sub-opinions (concurrence/dissent in addition to the
        lead opinion) and tries multiple text fields per sub-opinion.
        Falls back to cluster-level summary/syllabus when no full-text
        body is available — that's still enough for the verifier LLM to
        compare against the brief's proposition.

        Args:
            cluster_id: CourtListener cluster ID.

        Returns:
            Plain text of the opinion (or cluster summary), or None.
        """
        try:
            cluster_resp = _send_with_retry(
                lambda: requests.get(
                    f"{BASE_URL}/clusters/{cluster_id}/",
                    headers=self._headers(),
                    timeout=REQUEST_TIMEOUT,
                )
            )
            cluster_resp.raise_for_status()
            cluster_data = cluster_resp.json()

            sub_opinions = cluster_data.get("sub_opinions") or []

            # Try every sub-opinion in order; return the first non-empty body.
            for sub in sub_opinions:
                opinion_url: Optional[str] = None
                if isinstance(sub, str):
                    opinion_url = sub
                elif isinstance(sub, dict):
                    opinion_id = sub.get("id")
                    if opinion_id is not None:
                        opinion_url = f"{BASE_URL}/opinions/{opinion_id}/"
                if not opinion_url:
                    continue

                try:
                    opinion_resp = _send_with_retry(
                        lambda url=opinion_url: requests.get(
                            url,
                            headers=self._headers(),
                            timeout=REQUEST_TIMEOUT,
                        )
                    )
                    opinion_resp.raise_for_status()
                except Exception:
                    logger.warning(
                        "CourtListener sub-opinion fetch failed for %s",
                        opinion_url,
                        exc_info=True,
                    )
                    continue

                opinion_data = opinion_resp.json()
                text = _first_nonempty_text(opinion_data)
                if text:
                    return text

            # Fallback: cluster-level summary fields. These are short but
            # contain enough doctrinal substance for the verifier to
            # compare against the brief's proposition.
            cluster_text = _first_nonempty_text(cluster_data)
            if cluster_text:
                return cluster_text

            logger.warning(
                "No opinion text or cluster summary for cluster %s", cluster_id,
            )
            return None

        except Exception:
            logger.warning(
                "CourtListener opinion fetch failed for cluster %s",
                cluster_id,
                exc_info=True,
            )
            return None


# ---------------------------------------------------------------------------
# Text-extraction helpers
# ---------------------------------------------------------------------------

# Fields to try in order. CourtListener stores opinion bodies in multiple
# columns depending on the ingestion source (Resource.org, Columbia,
# LawBox, Harvard CAP, recent OCR, etc.). Cluster-level "syllabus" /
# "summary" / "headmatter" / "headnotes" are short but still useful for
# the verifier LLM when full text isn't available.
_TEXT_FIELDS = (
    "plain_text",
    "html_with_citations",
    "html",
    "html_columbia",
    "html_lawbox",
    "xml_harvard",
    "syllabus",
    "summary",
    "headmatter",
    "headnotes",
)


def _first_nonempty_text(record: dict) -> str:
    """Return the first non-empty body field from a CL opinion/cluster record."""
    for key in _TEXT_FIELDS:
        raw = record.get(key) or ""
        if not raw:
            continue
        stripped = re.sub(r"<[^>]+>", "", raw).strip()
        if stripped:
            return stripped
    return ""
