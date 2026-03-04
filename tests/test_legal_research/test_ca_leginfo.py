"""Tests for the California Legislative Info scraper client."""
import unittest
from unittest.mock import patch, MagicMock

from icharlotte_core.legal_research.sources.ca_leginfo import (
    CALegInfoClient,
    CODE_ALIASES,
    CODE_TITLES,
)
from icharlotte_core.legal_research.models import StatuteResult


class TestParseCodeReference(unittest.TestCase):
    """Tests for CALegInfoClient.parse_code_reference."""

    def setUp(self):
        self.client = CALegInfoClient()

    def test_parse_code_reference_civil_code(self):
        """Civil Code section 1714 -> ('CIV', '1714')."""
        code, section = self.client.parse_code_reference("Civil Code section 1714")
        self.assertEqual(code, "CIV")
        self.assertEqual(section, "1714")

    def test_parse_code_reference_ccp(self):
        """Code of Civil Procedure section 437c -> ('CCP', '437c')."""
        code, section = self.client.parse_code_reference(
            "Code of Civil Procedure section 437c"
        )
        self.assertEqual(code, "CCP")
        self.assertEqual(section, "437c")

    def test_parse_code_reference_abbreviated(self):
        """Civ. Code section-symbol 1714 -> ('CIV', '1714')."""
        code, section = self.client.parse_code_reference("Civ. Code \u00a7 1714")
        self.assertEqual(code, "CIV")
        self.assertEqual(section, "1714")

    def test_parse_code_reference_penal_code(self):
        """Penal Code section 187 -> ('PEN', '187')."""
        code, section = self.client.parse_code_reference("Penal Code section 187")
        self.assertEqual(code, "PEN")
        self.assertEqual(section, "187")

    def test_parse_code_reference_no_match(self):
        """Unrecognised text returns (None, None)."""
        code, section = self.client.parse_code_reference("some random text")
        self.assertIsNone(code)
        self.assertIsNone(section)

    def test_parse_code_reference_evidence_code(self):
        """Evid. Code section 352 -> ('EVID', '352')."""
        code, section = self.client.parse_code_reference("Evid. Code \u00a7 352")
        self.assertEqual(code, "EVID")
        self.assertEqual(section, "352")


class TestGetSection(unittest.TestCase):
    """Tests for CALegInfoClient.get_section with mocked HTTP."""

    def setUp(self):
        self.client = CALegInfoClient()

    @patch("icharlotte_core.legal_research.sources.ca_leginfo.requests.get")
    def test_get_section_returns_statute(self, mock_get):
        """Successful fetch returns a StatuteResult with parsed text."""
        html = """
        <html><body>
        <div id="codeLawSectionNoContent">
            <b>1714.</b> (a) Everyone is responsible, not only for the result
            of his or her willful acts, but also for an injury occasioned to
            another by his or her want of ordinary care or skill in the
            management of his or her property or person, except so far as the
            latter has, willfully or by want of ordinary care, brought the
            injury upon himself or herself.
        </div>
        </body></html>
        """
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.text = html
        mock_response.raise_for_status = MagicMock()
        mock_get.return_value = mock_response

        result = self.client.get_section("CIV", "1714")

        self.assertIsNotNone(result)
        self.assertIsInstance(result, StatuteResult)
        self.assertEqual(result.code, "CIV")
        self.assertEqual(result.section, "1714")
        self.assertIn("responsible", result.text)
        self.assertIn("leginfo.legislature.ca.gov", result.url)

    @patch("icharlotte_core.legal_research.sources.ca_leginfo.requests.get")
    def test_get_section_handles_not_found(self, mock_get):
        """Missing section div returns None."""
        html = "<html><body><p>No matching section found.</p></body></html>"
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.text = html
        mock_response.raise_for_status = MagicMock()
        mock_get.return_value = mock_response

        result = self.client.get_section("CIV", "99999999")
        self.assertIsNone(result)

    @patch("icharlotte_core.legal_research.sources.ca_leginfo.requests.get")
    def test_get_section_handles_http_error(self, mock_get):
        """HTTP error returns None."""
        mock_get.side_effect = Exception("Connection error")
        result = self.client.get_section("CIV", "1714")
        self.assertIsNone(result)

    @patch("icharlotte_core.legal_research.sources.ca_leginfo.requests.get")
    def test_get_section_too_short_text(self, mock_get):
        """If scraped text is too short, return None."""
        html = '<html><body><div id="codeLawSectionNoContent">Hi</div></body></html>'
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.text = html
        mock_response.raise_for_status = MagicMock()
        mock_get.return_value = mock_response

        result = self.client.get_section("CIV", "1714")
        self.assertIsNone(result)


class TestSearchCode(unittest.TestCase):
    """Tests for CALegInfoClient.search_code."""

    def setUp(self):
        self.client = CALegInfoClient()

    @patch("icharlotte_core.legal_research.sources.ca_leginfo.requests.get")
    def test_search_code_returns_results(self, mock_get):
        """search_code returns a list with one StatuteResult on success."""
        html = """
        <html><body>
        <div id="codeLawSectionNoContent">
            <b>437c.</b> (a) A party may move for summary judgment in an action
            if it is contended that the action has no merit or that there is
            no defense to the action or proceeding.
        </div>
        </body></html>
        """
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.text = html
        mock_response.raise_for_status = MagicMock()
        mock_get.return_value = mock_response

        results = self.client.search_code("CCP", "437c")

        self.assertEqual(len(results), 1)
        self.assertIsInstance(results[0], StatuteResult)
        self.assertEqual(results[0].code, "CCP")
        self.assertEqual(results[0].section, "437c")

    @patch("icharlotte_core.legal_research.sources.ca_leginfo.requests.get")
    def test_search_code_returns_empty_on_not_found(self, mock_get):
        """search_code returns empty list when section not found."""
        html = "<html><body><p>Not found</p></body></html>"
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.text = html
        mock_response.raise_for_status = MagicMock()
        mock_get.return_value = mock_response

        results = self.client.search_code("CIV", "99999999")
        self.assertEqual(results, [])


class TestCodeAliases(unittest.TestCase):
    """Verify CODE_ALIASES and CODE_TITLES dicts are populated."""

    def test_aliases_contains_civil_code(self):
        self.assertEqual(CODE_ALIASES["civil code"], "CIV")

    def test_aliases_contains_ccp_full(self):
        self.assertEqual(CODE_ALIASES["code of civil procedure"], "CCP")

    def test_aliases_contains_abbreviation(self):
        self.assertEqual(CODE_ALIASES["civ. code"], "CIV")
        self.assertEqual(CODE_ALIASES["ccp"], "CCP")

    def test_code_titles_has_civ(self):
        self.assertEqual(CODE_TITLES["CIV"], "Civil Code")

    def test_code_titles_has_ccp(self):
        self.assertEqual(CODE_TITLES["CCP"], "Code of Civil Procedure")


if __name__ == "__main__":
    unittest.main()
