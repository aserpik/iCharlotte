"""Client for scraping recent opinions from courts.ca.gov."""
import logging
import re
from typing import List
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

from icharlotte_core.legal_research.models import CaseResult

logger = logging.getLogger(__name__)

# Regex for dates in MM/DD/YYYY format
_DATE_RE = re.compile(r"\d{2}/\d{2}/\d{4}")


class CACourtsClient:
    """Scraper for California Courts recent opinions page."""

    OPINIONS_URL = "https://www.courts.ca.gov/opinions.htm"
    _TIMEOUT = 15

    def search_recent(self, query: str, max_results: int = 10) -> List[CaseResult]:
        """Search recent opinions by keyword matching on case names.

        Args:
            query: Search terms; any term matching the link text counts.
            max_results: Maximum number of results to return.

        Returns:
            List of CaseResult, empty on any error.
        """
        try:
            resp = requests.get(self.OPINIONS_URL, timeout=self._TIMEOUT)
            resp.raise_for_status()
        except Exception:
            logger.warning("Failed to fetch CA Courts opinions page", exc_info=True)
            return []

        try:
            return self._parse(resp.text, query, max_results)
        except Exception:
            logger.warning("Failed to parse CA Courts opinions page", exc_info=True)
            return []

    def _parse(
        self, html: str, query: str, max_results: int
    ) -> List[CaseResult]:
        """Parse the opinions HTML and return matching CaseResult entries."""
        soup = BeautifulSoup(html, "html.parser")
        results: List[CaseResult] = []

        query_terms = query.lower().split()

        for link in soup.find_all("a", href=True):
            href: str = link["href"]
            if not href.lower().endswith(".pdf"):
                continue

            link_text = link.get_text(strip=True)
            if not link_text:
                continue

            text_lower = link_text.lower()
            if not any(term in text_lower for term in query_terms):
                continue

            # Build full URL
            url = urljoin(self.OPINIONS_URL, href)

            # Try to extract date from sibling/parent table cells
            date = self._extract_date(link)

            results.append(
                CaseResult(
                    name=link_text,
                    citation="",
                    date=date,
                    court="",
                    snippet="",
                    url=url,
                )
            )

            if len(results) >= max_results:
                break

        return results

    @staticmethod
    def _extract_date(link_tag) -> str:
        """Try to find a MM/DD/YYYY date in the same table row as the link."""
        row = link_tag.find_parent("tr")
        if row:
            for cell in row.find_all("td"):
                cell_text = cell.get_text(strip=True)
                m = _DATE_RE.search(cell_text)
                if m:
                    return m.group(0)
        return ""
