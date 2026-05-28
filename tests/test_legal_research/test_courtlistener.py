"""Tests for CourtListenerClient — all HTTP calls are mocked."""

import unittest
from unittest.mock import patch, MagicMock

from icharlotte_core.legal_research.sources.courtlistener import (
    CourtListenerClient,
    CA_COURTS,
)
from icharlotte_core.legal_research.models import CaseResult


# ---------------------------------------------------------------------------
# Sample API response payloads
# ---------------------------------------------------------------------------

SAMPLE_SEARCH_RESPONSE = {
    "count": 2,
    "next": None,
    "previous": None,
    "results": [
        {
            "caseName": "Smith v. Jones",
            "citation": ["123 Cal.App.4th 456", "789 Cal.Rptr.2d 101"],
            "dateFiled": "2020-05-15",
            "court": "calctapp",
            "snippet": "The court held that <em>negligence</em> was established.",
            "absolute_url": "/opinion/12345/smith-v-jones/",
            "cluster_id": 12345,
        },
        {
            "caseName": "Doe v. Roe",
            "citation": ["45 Cal.5th 678"],
            "dateFiled": "2021-03-22",
            "court": "cal",
            "snippet": "Plaintiff failed to demonstrate <em>causation</em>.",
            "absolute_url": "/opinion/67890/doe-v-roe/",
            "cluster_id": 67890,
        },
    ],
}

SAMPLE_CLUSTER_RESPONSE = {
    "id": 12345,
    "case_name": "Smith v. Jones",
    "sub_opinions": [
        {"id": 111, "type": "010combined"},
    ],
}

SAMPLE_OPINION_RESPONSE = {
    "id": 111,
    "plain_text": "This is the full opinion text of Smith v. Jones.",
    "html_with_citations": "",
}


class TestCourtListenerInit(unittest.TestCase):
    """Construction and header tests."""

    def test_init_stores_token(self):
        client = CourtListenerClient(token="test-token-abc")
        self.assertEqual(client.token, "test-token-abc")

    def test_headers_include_auth(self):
        client = CourtListenerClient(token="my-secret")
        headers = client._headers()
        self.assertEqual(headers["Authorization"], "Token my-secret")
        # Should also request JSON
        self.assertIn("Content-Type", headers)


class TestSearchOpinions(unittest.TestCase):
    """Tests for search_opinions()."""

    def _make_client(self):
        return CourtListenerClient(token="tok")

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_search_opinions_returns_case_results(self, mock_get):
        resp = MagicMock()
        resp.status_code = 200
        resp.json.return_value = SAMPLE_SEARCH_RESPONSE
        resp.raise_for_status = MagicMock()
        mock_get.return_value = resp

        client = self._make_client()
        results = client.search_opinions("negligence")

        self.assertEqual(len(results), 2)
        self.assertIsInstance(results[0], CaseResult)
        self.assertEqual(results[0].name, "Smith v. Jones")
        self.assertEqual(results[0].citation, "123 Cal.App.4th 456")
        self.assertEqual(results[0].date, "2020-05-15")
        self.assertEqual(results[0].cluster_id, 12345)
        self.assertEqual(results[1].name, "Doe v. Roe")
        self.assertEqual(results[1].cluster_id, 67890)

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_search_opinions_handles_empty_results(self, mock_get):
        resp = MagicMock()
        resp.status_code = 200
        resp.json.return_value = {"count": 0, "results": []}
        resp.raise_for_status = MagicMock()
        mock_get.return_value = resp

        client = self._make_client()
        results = client.search_opinions("nonexistent query xyz")

        self.assertEqual(results, [])

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_search_opinions_handles_api_error(self, mock_get):
        resp = MagicMock()
        resp.status_code = 429
        resp.raise_for_status.side_effect = Exception("429 Too Many Requests")
        mock_get.return_value = resp

        client = self._make_client()
        results = client.search_opinions("negligence")

        # Should return empty list, not raise
        self.assertEqual(results, [])

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_search_filters_by_california(self, mock_get):
        resp = MagicMock()
        resp.status_code = 200
        resp.json.return_value = {"count": 0, "results": []}
        resp.raise_for_status = MagicMock()
        mock_get.return_value = resp

        client = self._make_client()
        client.search_opinions("test query")

        # Verify the call was made with correct params
        mock_get.assert_called_once()
        call_kwargs = mock_get.call_args
        params = call_kwargs.kwargs.get("params") or call_kwargs[1].get("params")
        self.assertEqual(params["type"], "o")
        self.assertEqual(params["court"], CA_COURTS)
        self.assertIn("test query", params["q"])
        self.assertEqual(params["order_by"], "score desc")

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_search_respects_max_results(self, mock_get):
        resp = MagicMock()
        resp.status_code = 200
        resp.json.return_value = {"count": 0, "results": []}
        resp.raise_for_status = MagicMock()
        mock_get.return_value = resp

        client = self._make_client()
        client.search_opinions("test", max_results=5)

        params = mock_get.call_args.kwargs.get("params") or mock_get.call_args[1].get("params")
        self.assertEqual(params["page_size"], 5)

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_search_prefers_california_reporter_citation(self, mock_get):
        resp = MagicMock()
        resp.status_code = 200
        resp.json.return_value = {
            "results": [
                {
                    "caseName": "Example v. Party",
                    "citation": ["471 P.3d 509", "10 Cal.5th 1"],
                    "dateFiled": "2020-05-15",
                    "court": "cal",
                    "snippet": "Relevant civil issue.",
                    "absolute_url": "/opinion/1/example/",
                    "cluster_id": 1,
                }
            ]
        }
        resp.raise_for_status = MagicMock()
        mock_get.return_value = resp

        client = self._make_client()
        results = client.search_opinions("civil issue")

        self.assertEqual(results[0].citation, "10 Cal.5th 1")


class TestGetCitingCases(unittest.TestCase):
    """Tests for get_citing_cases()."""

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_get_citing_cases(self, mock_get):
        citing_response = {
            "count": 1,
            "results": [
                {
                    "caseName": "Later Case v. Party",
                    "citation": ["99 Cal.App.5th 100"],
                    "dateFiled": "2023-01-10",
                    "court": "calctapp",
                    "snippet": "Citing Smith v. Jones for the negligence standard.",
                    "absolute_url": "/opinion/99999/later-case-v-party/",
                    "cluster_id": 99999,
                },
            ],
        }
        resp = MagicMock()
        resp.status_code = 200
        resp.json.return_value = citing_response
        resp.raise_for_status = MagicMock()
        mock_get.return_value = resp

        client = CourtListenerClient(token="tok")
        results = client.get_citing_cases(12345)

        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].name, "Later Case v. Party")

        # Verify query includes cites:(cluster_id)
        params = mock_get.call_args.kwargs.get("params") or mock_get.call_args[1].get("params")
        self.assertIn("cites:(12345)", params["q"])


class TestGetOpinionText(unittest.TestCase):
    """Tests for get_opinion_text()."""

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_get_opinion_text_returns_text(self, mock_get):
        # First call: cluster, second call: opinion
        cluster_resp = MagicMock()
        cluster_resp.status_code = 200
        cluster_resp.json.return_value = SAMPLE_CLUSTER_RESPONSE
        cluster_resp.raise_for_status = MagicMock()

        opinion_resp = MagicMock()
        opinion_resp.status_code = 200
        opinion_resp.json.return_value = SAMPLE_OPINION_RESPONSE
        opinion_resp.raise_for_status = MagicMock()

        mock_get.side_effect = [cluster_resp, opinion_resp]

        client = CourtListenerClient(token="tok")
        text = client.get_opinion_text(12345)

        self.assertEqual(text, "This is the full opinion text of Smith v. Jones.")

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_get_opinion_text_handles_error(self, mock_get):
        resp = MagicMock()
        resp.raise_for_status.side_effect = Exception("404")
        mock_get.return_value = resp

        client = CourtListenerClient(token="tok")
        text = client.get_opinion_text(99999)

        self.assertIsNone(text)

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_get_opinion_text_iterates_to_second_sub_opinion(self, mock_get):
        # Regression: first sub-opinion has no body fields populated; the
        # second has plain_text. We must keep going, not give up.
        cluster_resp = MagicMock()
        cluster_resp.raise_for_status = MagicMock()
        cluster_resp.json.return_value = {
            "sub_opinions": [
                "https://www.courtlistener.com/api/rest/v4/opinions/1/",
                "https://www.courtlistener.com/api/rest/v4/opinions/2/",
            ],
        }

        empty_opinion = MagicMock()
        empty_opinion.raise_for_status = MagicMock()
        empty_opinion.json.return_value = {"plain_text": "", "html_with_citations": ""}

        good_opinion = MagicMock()
        good_opinion.raise_for_status = MagicMock()
        good_opinion.json.return_value = {"plain_text": "Holding text here."}

        mock_get.side_effect = [cluster_resp, empty_opinion, good_opinion]

        client = CourtListenerClient(token="tok")
        text = client.get_opinion_text(12345)

        self.assertEqual(text, "Holding text here.")

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_get_opinion_text_tries_alternate_html_fields(self, mock_get):
        # plain_text and html_with_citations are empty; html_columbia has
        # the body. Must extract it.
        cluster_resp = MagicMock()
        cluster_resp.raise_for_status = MagicMock()
        cluster_resp.json.return_value = {
            "sub_opinions": [
                "https://www.courtlistener.com/api/rest/v4/opinions/1/"
            ],
        }
        opinion_resp = MagicMock()
        opinion_resp.raise_for_status = MagicMock()
        opinion_resp.json.return_value = {
            "plain_text": "",
            "html_with_citations": "",
            "html_columbia": "<p>The court held X.</p>",
        }
        mock_get.side_effect = [cluster_resp, opinion_resp]

        client = CourtListenerClient(token="tok")
        text = client.get_opinion_text(12345)

        self.assertEqual(text, "The court held X.")

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_get_opinion_text_falls_back_to_cluster_syllabus(self, mock_get):
        # All sub-opinions are empty; cluster has a syllabus that should
        # be returned as the authority text.
        cluster_resp = MagicMock()
        cluster_resp.raise_for_status = MagicMock()
        cluster_resp.json.return_value = {
            "sub_opinions": [
                "https://www.courtlistener.com/api/rest/v4/opinions/1/"
            ],
            "syllabus": "<p>The Court held that discovery is limited.</p>",
        }
        opinion_resp = MagicMock()
        opinion_resp.raise_for_status = MagicMock()
        opinion_resp.json.return_value = {"plain_text": "", "html": ""}
        mock_get.side_effect = [cluster_resp, opinion_resp]

        client = CourtListenerClient(token="tok")
        text = client.get_opinion_text(12345)

        self.assertEqual(text, "The Court held that discovery is limited.")

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_get_opinion_text_returns_none_when_nothing_available(self, mock_get):
        cluster_resp = MagicMock()
        cluster_resp.raise_for_status = MagicMock()
        cluster_resp.json.return_value = {"sub_opinions": []}
        mock_get.return_value = cluster_resp

        client = CourtListenerClient(token="tok")
        text = client.get_opinion_text(12345)

        self.assertIsNone(text)


if __name__ == "__main__":
    unittest.main()
