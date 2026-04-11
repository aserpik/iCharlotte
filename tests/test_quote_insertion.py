"""Tests for quote insertion feature."""
import os
import tempfile
import unittest
from unittest.mock import patch, MagicMock


class TestQuoteResultParsing(unittest.TestCase):

    def test_parse_single_quote_result(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        llm_output = (
            "QUOTE_RESULT_START\n"
            "DEPONENT: Haydel\n"
            "SOURCE: 25.05.29 Depo of Benjamin Haydel.pdf\n"
            "PAGE_LINE: 35:4-8\n"
            "RELEVANCE: Plaintiff admits seeing the plastic sheeting\n"
            "Q. Did you see the plastic on the ground before you fell?\n"
            "A. Yeah, I saw it -- I saw it coming up the stairs.\n"
            "QUOTE_RESULT_END"
        )
        results = gen._parse_quote_results(llm_output)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["deponent"], "Haydel")
        self.assertEqual(results[0]["source"], "25.05.29 Depo of Benjamin Haydel.pdf")
        self.assertEqual(results[0]["page_line"], "35:4-8")
        self.assertIn("Q. Did you see", results[0]["qa_text"])
        self.assertIn("A. Yeah, I saw it", results[0]["qa_text"])

    def test_parse_multiple_quote_results(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        llm_output = (
            "QUOTE_RESULT_START\n"
            "DEPONENT: Haydel\n"
            "SOURCE: depo1.pdf\n"
            "PAGE_LINE: 35:4-8\n"
            "RELEVANCE: First relevant passage\n"
            "Q. First question?\n"
            "A. First answer.\n"
            "QUOTE_RESULT_END\n\n"
            "QUOTE_RESULT_START\n"
            "DEPONENT: Smith\n"
            "SOURCE: depo2.pdf\n"
            "PAGE_LINE: 12:1-5\n"
            "RELEVANCE: Second relevant passage\n"
            "Q. Second question?\n"
            "A. Second answer.\n"
            "QUOTE_RESULT_END"
        )
        results = gen._parse_quote_results(llm_output)
        self.assertEqual(len(results), 2)
        self.assertEqual(results[0]["deponent"], "Haydel")
        self.assertEqual(results[1]["deponent"], "Smith")

    def test_parse_no_matches(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        results = gen._parse_quote_results("NO_MATCHES_FOUND")
        self.assertEqual(results, [])

    def test_parse_empty_response(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        results = gen._parse_quote_results("")
        self.assertEqual(results, [])


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
