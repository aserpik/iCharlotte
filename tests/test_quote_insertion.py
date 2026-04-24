"""Tests for quote insertion feature."""
import os
import tempfile
import unittest
from unittest.mock import patch, MagicMock


# NOTE: the old LLM quote-output parser (_parse_quote_results) was replaced
# by the grounded ID-based pipeline in tests/test_mediation_brief_quote_search.py.
# The old parser tested free-form LLM output; the new parser resolves IDs
# against a pre-computed candidate list and therefore needs different fixtures.


class TestQuoteInsertion(unittest.TestCase):

    def test_insert_quote_at_end_of_section(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        gen.sections = {
            "liability": "Existing liability argument text.\n\nSUBSECTION: No Duty\nDuty argument here."
        }
        quote = {
            "deponent": "Haydel",
            "source": "depo.pdf",
            "page_line": "35:4-8",
            "qa_text": "Q. Did you see it?\nA. Yes, I saw it.",
            "relevance": "Admits seeing hazard",
        }
        gen.insert_quotes_quick([quote], "liability", None)
        updated = gen.sections["liability"]
        self.assertIn("DEPO_QUOTE_START", updated)
        self.assertIn("Q. Did you see it?", updated)
        self.assertIn("(Haydel Depo Trns., at p. 35:4-8.)", updated)

    def test_insert_quote_into_specific_subsection(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        gen.sections = {
            "liability": (
                "Intro paragraph.\n\n"
                "SUBSECTION: No Duty\n"
                "Duty argument here.\n\n"
                "SUBSECTION: No Breach\n"
                "Breach argument here."
            )
        }
        quote = {
            "deponent": "Haydel",
            "source": "depo.pdf",
            "page_line": "35:4-8",
            "qa_text": "Q. Did you see it?\nA. Yes.",
            "relevance": "Relevant to duty",
        }
        gen.insert_quotes_quick([quote], "liability", "No Duty")
        updated = gen.sections["liability"]
        duty_pos = updated.index("Duty argument")
        quote_pos = updated.index("DEPO_QUOTE_START")
        breach_pos = updated.index("SUBSECTION: No Breach")
        self.assertGreater(quote_pos, duty_pos)
        self.assertLess(quote_pos, breach_pos)


class TestSavedPath(unittest.TestCase):

    def test_saved_path_initialized_to_none(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        self.assertIsNone(gen.saved_path)

    def test_saved_path_cleared_on_reset(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        gen.saved_path = "/some/path.docx"
        gen.reset()
        self.assertIsNone(gen.saved_path)


if __name__ == '__main__':
    unittest.main()
