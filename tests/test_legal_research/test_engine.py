import json
import unittest
from unittest.mock import MagicMock, patch
from icharlotte_core.legal_research.engine import LegalResearchEngine
from icharlotte_core.legal_research.models import CaseResult, StatuteResult, ResearchResult


class TestLegalResearchEngine(unittest.TestCase):
    def setUp(self):
        self.engine = LegalResearchEngine(courtlistener_token="test-token")

    @patch.object(LegalResearchEngine, '_enrich_top_cases')
    @patch.object(LegalResearchEngine, '_search_sources')
    @patch.object(LegalResearchEngine, '_plan_queries')
    def test_research_returns_result(self, mock_plan, mock_search, mock_enrich):
        cases = [CaseResult(
            name="Rowland v. Christian", citation="69 Cal.2d 108",
            date="1968-08-08", court="Supreme Court of California",
            snippet="Duty of care", url="http://test", cluster_id=1
        )]
        statutes = [StatuteResult(
            code="CIV", section="1714", title="Civil Code",
            text="Everyone is responsible...", url="http://test"
        )]
        mock_plan.return_value = {
            "case_queries": ["premises liability duty of care"],
            "statute_queries": ["Civil Code section 1714"],
            "legal_doctrines": ["premises liability"],
        }
        mock_search.return_value = (cases, statutes, [], [])
        mock_enrich.return_value = cases  # pass through

        def mock_llm(system, user):
            return "Analysis with citations."

        result = self.engine.research("premises liability", llm_callback=mock_llm)
        self.assertIsInstance(result, ResearchResult)
        self.assertEqual(len(result.cases), 1)
        self.assertEqual(len(result.statutes), 1)
        self.assertEqual(result.query, "premises liability")

    def test_research_with_no_results(self):
        # Mock all source clients to return empty
        self.engine.cl_client = MagicMock()
        self.engine.cl_client.search_opinions.return_value = []
        self.engine.leginfo_client = MagicMock()
        self.engine.leginfo_client.get_section.return_value = None
        self.engine.leginfo_client.parse_code_reference.return_value = (None, None)
        self.engine.courts_client = MagicMock()
        self.engine.courts_client.search_recent.return_value = []

        def mock_llm(system, user):
            if "search terms" in system.lower() or "extract" in system.lower():
                return json.dumps({
                    "case_queries": ["nonexistent topic"],
                    "statute_queries": [],
                    "legal_topics": ["nothing"],
                })
            return "No relevant authority found."

        result = self.engine.research("nonexistent topic", llm_callback=mock_llm)
        self.assertIsInstance(result, ResearchResult)
        self.assertEqual(len(result.cases), 0)

    def test_cache_key_generation(self):
        key1 = self.engine._cache_key("premises liability dog bite")
        key2 = self.engine._cache_key("premises liability dog bite")
        key3 = self.engine._cache_key("different query")
        self.assertEqual(key1, key2)
        self.assertNotEqual(key1, key3)

    def test_from_cache_preserves_snippet_provenance(self):
        result = self.engine._from_cache({
            "query": "summary judgment",
            "cases": [
                {
                    "name": "Aguilar v. Atlantic Richfield Co.",
                    "citation": "25 Cal. 4th 826",
                    "date": "2001-06-14",
                    "court": "Cal.",
                    "snippet": "describing Aguilar as allocating the summary judgment burden",
                    "snippet_source": "parenthetical",
                    "snippet_parenthetical_id": "900",
                    "url": "u",
                    "cluster_id": "cap:aguilar",
                }
            ],
        })

        self.assertEqual(result.cases[0].snippet_source, "parenthetical")
        self.assertEqual(result.cases[0].snippet_parenthetical_id, "900")


if __name__ == "__main__":
    unittest.main()
