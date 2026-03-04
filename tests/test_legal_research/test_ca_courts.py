"""Tests for CACourtsClient — courts.ca.gov recent opinions scraper."""
import unittest
from unittest.mock import patch, MagicMock

import requests

from icharlotte_core.legal_research.sources.ca_courts import CACourtsClient


MOCK_OPINIONS_HTML = """
<html>
<body>
<table>
  <tr>
    <td>02/28/2026</td>
    <td><a href="/opinions/documents/S123456.PDF">Smith v. Jones Corp.</a></td>
    <td>Supreme Court</td>
  </tr>
  <tr>
    <td>02/25/2026</td>
    <td><a href="/opinions/documents/B234567.pdf">Doe v. City of Los Angeles</a></td>
    <td>Court of Appeal, Second District</td>
  </tr>
  <tr>
    <td>02/20/2026</td>
    <td><a href="/opinions/documents/C345678.PDF">People v. Garcia</a></td>
    <td>Court of Appeal, Third District</td>
  </tr>
  <tr>
    <td>02/18/2026</td>
    <td><a href="/other/doc.html">Not a PDF link</a></td>
    <td>Some Court</td>
  </tr>
</table>
</body>
</html>
"""


class TestCACourtsClient(unittest.TestCase):
    """Tests for CACourtsClient."""

    def setUp(self):
        self.client = CACourtsClient()

    @patch("icharlotte_core.legal_research.sources.ca_courts.requests.get")
    def test_search_recent_opinions(self, mock_get):
        """Should parse opinion links and filter by keyword."""
        resp = MagicMock()
        resp.status_code = 200
        resp.text = MOCK_OPINIONS_HTML
        resp.raise_for_status = MagicMock()
        mock_get.return_value = resp

        results = self.client.search_recent("Smith")

        self.assertEqual(len(results), 1)
        r = results[0]
        self.assertEqual(r.name, "Smith v. Jones Corp.")
        self.assertIn("S123456.PDF", r.url)
        self.assertEqual(r.date, "02/28/2026")
        self.assertEqual(r.citation, "")

    @patch("icharlotte_core.legal_research.sources.ca_courts.requests.get")
    def test_search_recent_multiple_matches(self, mock_get):
        """Should return multiple matches when query matches several cases."""
        resp = MagicMock()
        resp.status_code = 200
        resp.text = MOCK_OPINIONS_HTML
        resp.raise_for_status = MagicMock()
        mock_get.return_value = resp

        # "v." appears in all case names
        results = self.client.search_recent("v.")

        self.assertEqual(len(results), 3)

    @patch("icharlotte_core.legal_research.sources.ca_courts.requests.get")
    def test_search_case_insensitive(self, mock_get):
        """Keyword matching should be case-insensitive."""
        resp = MagicMock()
        resp.status_code = 200
        resp.text = MOCK_OPINIONS_HTML
        resp.raise_for_status = MagicMock()
        mock_get.return_value = resp

        results = self.client.search_recent("garcia")

        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].name, "People v. Garcia")

    @patch("icharlotte_core.legal_research.sources.ca_courts.requests.get")
    def test_max_results_limit(self, mock_get):
        """Should respect max_results parameter."""
        resp = MagicMock()
        resp.status_code = 200
        resp.text = MOCK_OPINIONS_HTML
        resp.raise_for_status = MagicMock()
        mock_get.return_value = resp

        results = self.client.search_recent("v.", max_results=2)

        self.assertEqual(len(results), 2)

    @patch("icharlotte_core.legal_research.sources.ca_courts.requests.get")
    def test_handles_connection_error(self, mock_get):
        """Should return empty list on connection error."""
        mock_get.side_effect = requests.ConnectionError("Network unreachable")

        results = self.client.search_recent("anything")

        self.assertEqual(results, [])

    @patch("icharlotte_core.legal_research.sources.ca_courts.requests.get")
    def test_handles_http_error(self, mock_get):
        """Should return empty list on HTTP error status."""
        resp = MagicMock()
        resp.raise_for_status.side_effect = requests.HTTPError("503 Server Error")
        mock_get.return_value = resp

        results = self.client.search_recent("anything")

        self.assertEqual(results, [])

    @patch("icharlotte_core.legal_research.sources.ca_courts.requests.get")
    def test_no_matches_returns_empty(self, mock_get):
        """Should return empty list when no case names match query."""
        resp = MagicMock()
        resp.status_code = 200
        resp.text = MOCK_OPINIONS_HTML
        resp.raise_for_status = MagicMock()
        mock_get.return_value = resp

        results = self.client.search_recent("zzzznonexistent")

        self.assertEqual(results, [])


if __name__ == "__main__":
    unittest.main()
