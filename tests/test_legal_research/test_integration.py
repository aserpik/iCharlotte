"""End-to-end integration test for the legal research pipeline.

Mocks all external calls (HTTP requests + LLM) and verifies the full
pipeline from query to verified result.
"""
import json
import unittest
from unittest.mock import MagicMock, patch

from icharlotte_core.legal_research.engine import LegalResearchEngine
from icharlotte_core.legal_research.models import (
    CaseResult,
    ResearchResult,
    StatuteResult,
    VerificationStatus,
)


# ---------------------------------------------------------------------------
# Mock HTTP responses
# ---------------------------------------------------------------------------

def _courtlistener_response():
    """Fake CourtListener search API response with Rowland v. Christian."""
    return {
        "count": 1,
        "results": [
            {
                "caseName": "Rowland v. Christian",
                "citation": ["69 Cal.2d 108"],
                "dateFiled": "1968-08-08",
                "court": "cal",
                "snippet": (
                    "A person who maintains premises in a dangerous condition "
                    "owes a <em>duty of care</em> to all who enter."
                ),
                "absolute_url": "/opinion/1234/rowland-v-christian/",
                "cluster_id": 1234,
            }
        ],
    }


def _leginfo_html():
    """Fake leginfo.legislature.ca.gov HTML for Civil Code section 1714."""
    return (
        '<html><body>'
        '<div id="codeLawSectionNoHead">'
        '(a) Everyone is responsible, not only for the result of his or her '
        'willful acts, but also for an injury occasioned to another by his or '
        'her want of ordinary care or skill in the management of his or her '
        'property or person, except so far as the latter has, willfully or by '
        'want of ordinary care, brought the injury upon himself or herself.'
        '</div>'
        '</body></html>'
    )


def _ca_courts_html():
    """Fake CA Courts opinions page with no matching links."""
    return '<html><body><p>No recent opinions.</p></body></html>'


def _mock_requests_get(url, **kwargs):
    """Route mock requests.get calls based on URL to the correct fake response."""
    resp = MagicMock()
    resp.raise_for_status = MagicMock()

    if "courtlistener.com" in url:
        resp.status_code = 200
        resp.json.return_value = _courtlistener_response()
    elif "leginfo.legislature.ca.gov" in url:
        resp.status_code = 200
        resp.text = _leginfo_html()
    elif "courts.ca.gov" in url:
        resp.status_code = 200
        resp.text = _ca_courts_html()
    else:
        resp.status_code = 404
        resp.raise_for_status.side_effect = Exception(f"Unexpected URL: {url}")

    return resp


# ---------------------------------------------------------------------------
# Mock LLM callback
# ---------------------------------------------------------------------------

class _LLMCallTracker:
    """Tracks calls to the mock LLM and returns canned responses."""

    def __init__(self):
        self.calls = []

    def __call__(self, system_prompt: str, user_prompt: str) -> str:
        call_number = len(self.calls) + 1
        self.calls.append((system_prompt, user_prompt))

        # Call 1: Query planning -- return structured search plan
        if call_number == 1:
            return json.dumps({
                "case_queries": ["premises liability duty of care California"],
                "statute_queries": ["Civil Code section 1714"],
                "legal_topics": ["premises liability"],
            })

        # Call 2: Synthesis -- return a memo citing Rowland and Civ. Code 1714
        if call_number == 2:
            return (
                "LEGAL RESEARCH MEMORANDUM\n\n"
                "I. Issue\n"
                "Whether a property owner owes a duty of care to persons "
                "injured on the premises.\n\n"
                "II. Governing Law\n"
                "Under California law, the general duty of care in premises "
                "liability cases is established by Rowland v. Christian (1968) "
                "69 Cal.2d 108 and codified in Civ. Code, \u00a7 1714.\n\n"
                "III. Analysis\n"
                "In Rowland v. Christian, the California Supreme Court held "
                "that a property owner owes a duty of ordinary care to all "
                "persons on the premises. The court enumerated factors for "
                "evaluating the scope of this duty.\n\n"
                "Civil Code section 1714 provides that everyone is responsible "
                "for injuries caused by want of ordinary care.\n\n"
                "IV. Conclusion\n"
                "The property owner likely owes a duty of care under Rowland "
                "and Civ. Code, \u00a7 1714."
            )

        # Call 3: Verification -- return PASS for both citations
        if call_number == 3:
            return json.dumps({
                "verifications": [
                    {
                        "citation": "Rowland v. Christian (1968) 69 Cal.2d 108",
                        "status": "PASS",
                        "detail": "Citation verified against CourtListener source data.",
                        "original": None,
                        "corrected": None,
                    },
                    {
                        "citation": "Civ. Code, \u00a7 1714",
                        "status": "PASS",
                        "detail": "Statute text confirmed from CA Legislative Info.",
                        "original": None,
                        "corrected": None,
                    },
                ],
            })

        # Fallback for unexpected calls
        return "Unexpected LLM call"


# ---------------------------------------------------------------------------
# Test class
# ---------------------------------------------------------------------------

class TestLegalResearchIntegration(unittest.TestCase):
    """Full pipeline test with mocked API + LLM calls."""

    @patch("requests.get", side_effect=_mock_requests_get)
    def test_full_pipeline(self, mock_get):
        # ---- Set up engine and LLM callback ----
        engine = LegalResearchEngine(courtlistener_token="test-token-123")
        llm = _LLMCallTracker()
        statuses = []

        # ---- Run full pipeline ----
        result = engine.research(
            query="What duty of care does a property owner owe under California premises liability law?",
            llm_callback=llm,
            status_callback=lambda msg: statuses.append(msg),
        )

        # ---- Assert result structure ----
        self.assertIsInstance(result, ResearchResult)

        # Cases: should contain Rowland v. Christian
        self.assertGreaterEqual(len(result.cases), 1)
        rowland = next(
            (c for c in result.cases if "Rowland" in c.name), None
        )
        self.assertIsNotNone(rowland, "Rowland v. Christian not found in cases")
        self.assertEqual(rowland.citation, "69 Cal.2d 108")
        self.assertEqual(rowland.date, "1968-08-08")

        # Statutes: should contain Civ. Code section 1714
        self.assertGreaterEqual(len(result.statutes), 1)
        civ1714 = next(
            (s for s in result.statutes if s.code == "CIV" and s.section == "1714"),
            None,
        )
        self.assertIsNotNone(civ1714, "Civil Code section 1714 not found in statutes")
        self.assertIn("ordinary care", civ1714.text)

        # Memo: should be non-empty and cite both authorities
        self.assertTrue(len(result.memo) > 100, "Memo too short")
        self.assertIn("Rowland", result.memo)
        self.assertIn("1714", result.memo)

        # Verification: should have entries with PASS status
        self.assertGreaterEqual(len(result.verification), 1)
        pass_count = sum(1 for v in result.verification if v.status == "PASS")
        self.assertGreaterEqual(pass_count, 1, "Expected at least one PASS verification")

        # ---- Assert authority block content ----
        authority = result.format_authority_block()
        self.assertIn("[LEGAL AUTHORITY]", authority)
        self.assertIn("Rowland v. Christian", authority)
        self.assertIn("Case Law:", authority)
        self.assertIn("Statutes:", authority)
        self.assertIn("\u00a7 1714", authority)

        # ---- Assert LLM was called 3 times (plan, synthesis, verification) ----
        self.assertEqual(len(llm.calls), 3, f"Expected 3 LLM calls, got {len(llm.calls)}")

        # ---- Assert status callbacks were emitted ----
        self.assertGreaterEqual(len(statuses), 3, "Expected at least 3 status updates")
        self.assertEqual(statuses[-1], "Research complete")

        # ---- Assert HTTP was called (CourtListener + Leginfo + CA Courts) ----
        self.assertTrue(mock_get.called, "requests.get was never called")
        call_urls = [call.args[0] for call in mock_get.call_args_list]
        self.assertTrue(
            any("courtlistener.com" in u for u in call_urls),
            f"CourtListener was not called. URLs: {call_urls}",
        )
        self.assertTrue(
            any("leginfo.legislature.ca.gov" in u for u in call_urls),
            f"CA Leginfo was not called. URLs: {call_urls}",
        )
        self.assertTrue(
            any("courts.ca.gov" in u for u in call_urls),
            f"CA Courts was not called. URLs: {call_urls}",
        )


if __name__ == "__main__":
    unittest.main()
