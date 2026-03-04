"""Tests for legal research data models."""
import unittest
from icharlotte_core.legal_research.models import (
    CaseResult,
    StatuteResult,
    VerificationStatus,
    ResearchResult,
    CODE_DISPLAY_NAMES,
)


class TestCaseResult(unittest.TestCase):
    """Tests for CaseResult dataclass."""

    def test_creation_required_fields(self):
        case = CaseResult(
            name="Smith v. Jones",
            citation="123 Cal.App.4th 456",
            date="2020-01-15",
            court="Court of Appeal",
            snippet="The court held that...",
            url="https://example.com/case/123",
        )
        self.assertEqual(case.name, "Smith v. Jones")
        self.assertEqual(case.citation, "123 Cal.App.4th 456")
        self.assertEqual(case.date, "2020-01-15")
        self.assertEqual(case.court, "Court of Appeal")
        self.assertEqual(case.snippet, "The court held that...")
        self.assertEqual(case.url, "https://example.com/case/123")

    def test_optional_fields_default(self):
        case = CaseResult(
            name="Smith v. Jones",
            citation="123 Cal.App.4th 456",
            date="2020-01-15",
            court="Court of Appeal",
            snippet="snippet",
            url="https://example.com",
        )
        self.assertIsNone(case.cluster_id)
        self.assertIsNone(case.negative_treatment)
        self.assertAlmostEqual(case.relevance_score, 0.0)

    def test_optional_fields_set(self):
        case = CaseResult(
            name="Smith v. Jones",
            citation="123 Cal.App.4th 456",
            date="2020-01-15",
            court="Court of Appeal",
            snippet="snippet",
            url="https://example.com",
            cluster_id=42,
            negative_treatment="Overruled by Doe v. Roe",
            relevance_score=0.95,
        )
        self.assertEqual(case.cluster_id, 42)
        self.assertEqual(case.negative_treatment, "Overruled by Doe v. Roe")
        self.assertAlmostEqual(case.relevance_score, 0.95)

    def test_formatted_citation(self):
        case = CaseResult(
            name="Smith v. Jones",
            citation="123 Cal.App.4th 456",
            date="2020-01-15",
            court="Court of Appeal",
            snippet="snippet",
            url="https://example.com",
        )
        self.assertEqual(
            case.formatted_citation,
            "Smith v. Jones (2020) 123 Cal.App.4th 456",
        )

    def test_formatted_citation_different_year(self):
        case = CaseResult(
            name="Doe v. Roe",
            citation="45 Cal.5th 789",
            date="2018-06-30",
            court="Supreme Court",
            snippet="snippet",
            url="https://example.com",
        )
        self.assertEqual(
            case.formatted_citation,
            "Doe v. Roe (2018) 45 Cal.5th 789",
        )


class TestStatuteResult(unittest.TestCase):
    """Tests for StatuteResult dataclass."""

    def test_creation(self):
        statute = StatuteResult(
            code="CIV",
            section="1714",
            title="General duty of care",
            text="Everyone is responsible for injuries...",
            url="https://example.com/statute",
        )
        self.assertEqual(statute.code, "CIV")
        self.assertEqual(statute.section, "1714")
        self.assertEqual(statute.title, "General duty of care")

    def test_formatted_citation_civ(self):
        statute = StatuteResult(
            code="CIV",
            section="1714",
            title="General duty of care",
            text="Everyone is responsible...",
            url="https://example.com",
        )
        self.assertEqual(statute.formatted_citation, "Civ. Code, \u00a7 1714")

    def test_formatted_citation_ccp(self):
        statute = StatuteResult(
            code="CCP",
            section="335.1",
            title="Statute of limitations",
            text="...",
            url="https://example.com",
        )
        self.assertEqual(
            statute.formatted_citation, "Code Civ. Proc., \u00a7 335.1"
        )

    def test_formatted_citation_unknown_code(self):
        """Unknown codes should fall back to the raw abbreviation."""
        statute = StatuteResult(
            code="UNKNOWN",
            section="100",
            title="Test",
            text="...",
            url="https://example.com",
        )
        self.assertEqual(statute.formatted_citation, "UNKNOWN, \u00a7 100")


class TestVerificationStatus(unittest.TestCase):
    """Tests for VerificationStatus dataclass."""

    def test_creation_pass(self):
        vs = VerificationStatus(
            citation="Smith v. Jones (2020) 123 Cal.App.4th 456",
            status="PASS",
            detail="Citation verified via CourtListener",
        )
        self.assertEqual(vs.status, "PASS")
        self.assertIsNone(vs.original)
        self.assertIsNone(vs.corrected)

    def test_creation_fixed(self):
        vs = VerificationStatus(
            citation="Smith v. Jones (2020) 123 Cal.App.4th 456",
            status="FIXED",
            detail="Year corrected from 2019 to 2020",
            original="Smith v. Jones (2019) 123 Cal.App.4th 456",
            corrected="Smith v. Jones (2020) 123 Cal.App.4th 456",
        )
        self.assertEqual(vs.status, "FIXED")
        self.assertIsNotNone(vs.original)
        self.assertIsNotNone(vs.corrected)

    def test_creation_flagged(self):
        vs = VerificationStatus(
            citation="Fake v. Case (2020) 999 Cal.5th 1",
            status="FLAGGED",
            detail="Citation not found in any database",
        )
        self.assertEqual(vs.status, "FLAGGED")


class TestResearchResult(unittest.TestCase):
    """Tests for ResearchResult dataclass."""

    def test_empty_result(self):
        result = ResearchResult(query="negligence standard of care")
        self.assertEqual(result.query, "negligence standard of care")
        self.assertEqual(result.cases, [])
        self.assertEqual(result.statutes, [])
        self.assertEqual(result.memo, "")
        self.assertEqual(result.verification, [])

    def test_format_authority_block_empty(self):
        result = ResearchResult(query="test query")
        block = result.format_authority_block()
        self.assertIn("[LEGAL AUTHORITY]", block)

    def test_format_authority_block_with_cases_and_statutes(self):
        case = CaseResult(
            name="Smith v. Jones",
            citation="123 Cal.App.4th 456",
            date="2020-01-15",
            court="Court of Appeal",
            snippet="The court held that...",
            url="https://example.com",
        )
        statute = StatuteResult(
            code="CIV",
            section="1714",
            title="General duty of care",
            text="Everyone is responsible...",
            url="https://example.com",
        )
        result = ResearchResult(
            query="negligence",
            cases=[case],
            statutes=[statute],
        )
        block = result.format_authority_block()
        self.assertIn("[LEGAL AUTHORITY]", block)
        self.assertIn("Smith v. Jones (2020) 123 Cal.App.4th 456", block)
        self.assertIn("Civ. Code, \u00a7 1714", block)

    def test_to_dict(self):
        case = CaseResult(
            name="Smith v. Jones",
            citation="123 Cal.App.4th 456",
            date="2020-01-15",
            court="Court of Appeal",
            snippet="snippet",
            url="https://example.com",
            relevance_score=0.8,
        )
        vs = VerificationStatus(
            citation="Smith v. Jones (2020) 123 Cal.App.4th 456",
            status="PASS",
            detail="Verified",
        )
        result = ResearchResult(
            query="test",
            cases=[case],
            memo="Analysis memo",
            verification=[vs],
        )
        d = result.to_dict()
        self.assertEqual(d["query"], "test")
        self.assertIsInstance(d["cases"], list)
        self.assertEqual(len(d["cases"]), 1)
        self.assertEqual(d["cases"][0]["name"], "Smith v. Jones")
        self.assertEqual(d["cases"][0]["relevance_score"], 0.8)
        self.assertEqual(d["memo"], "Analysis memo")
        self.assertEqual(len(d["verification"]), 1)
        self.assertEqual(d["verification"][0]["status"], "PASS")

    def test_to_dict_roundtrip_types(self):
        """Ensure to_dict returns only JSON-serializable types."""
        import json

        result = ResearchResult(query="test")
        d = result.to_dict()
        # Should not raise
        json.dumps(d)


class TestCodeDisplayNames(unittest.TestCase):
    """Tests for CODE_DISPLAY_NAMES mapping."""

    def test_major_codes_present(self):
        expected_codes = [
            "CIV", "CCP", "PEN", "VEH", "GOV", "EVID", "FAM",
            "PROB", "BPC", "INS", "LAB", "HSC", "WIC", "EDC",
        ]
        for code in expected_codes:
            self.assertIn(code, CODE_DISPLAY_NAMES, f"Missing code: {code}")

    def test_display_names_not_empty(self):
        for code, name in CODE_DISPLAY_NAMES.items():
            self.assertTrue(len(name) > 0, f"Empty display name for {code}")


if __name__ == "__main__":
    unittest.main()
