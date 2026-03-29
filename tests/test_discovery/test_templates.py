"""
Tests for icharlotte_core.discovery.templates module.

Covers:
- substitute_variables: placeholder and blank replacement
- extract_requests_from_text: parsing SI and RPD request text
- TemplateLoader: loading from actual template .docx files (skipped if absent)
"""
import os
import unittest

from icharlotte_core.discovery.models import DiscoveryRequest, DiscoveryType
from icharlotte_core.discovery.templates import (
    TemplateLoader,
    extract_requests_from_text,
    substitute_variables,
)

# Path to the discovery template directory
DISCOVERY_DIR = os.path.join(os.path.dirname(__file__), "..", "..", "discovery")
DISCOVERY_DIR = os.path.normpath(DISCOVERY_DIR)


# ---------------------------------------------------------------------------
# TestSubstituteVariables
# ---------------------------------------------------------------------------

class TestSubstituteVariables(unittest.TestCase):
    """Tests for the substitute_variables function."""

    def test_replaces_named_placeholders(self):
        """Named {key} placeholders are replaced with values from the dict."""
        text = "Defendant, {defendant_name} hereby propounds."
        result = substitute_variables(text, {"defendant_name": "ACME CORP"})
        self.assertEqual(result, "Defendant, ACME CORP hereby propounds.")

    def test_replaces_responding_party_blank(self):
        """Contextual blank after 'Responding Party, ____' is filled."""
        text = "Responding Party, ____ shall answer"
        result = substitute_variables(
            text, {"responding_party_name": "RUXANDRA RASCHKOVSKY"}
        )
        self.assertIn("RUXANDRA RASCHKOVSKY", result)
        self.assertNotIn("____", result)

    def test_replaces_propounding_party_blank(self):
        """Contextual blank after 'Propounding Party, ____' is filled."""
        text = "Propounding Party, ____ demands production"
        result = substitute_variables(
            text, {"propounding_party_name": "SERVITEK ELECTRIC, INC."}
        )
        self.assertIn("SERVITEK ELECTRIC, INC.", result)
        self.assertNotIn("____", result)

    def test_replaces_plaintiff_blank(self):
        """Contextual blank after 'Plaintiff, ____' or 'Plaintiff ____' is filled."""
        text = 'refers to the Responding Party, Plaintiff, ____ and includes'
        result = substitute_variables(
            text, {"responding_party_name": "JOHN DOE"}
        )
        self.assertIn("JOHN DOE", result)

    def test_preserves_non_placeholder_text(self):
        """Text without any placeholders or blanks passes through unchanged."""
        text = "This is normal text with no placeholders."
        result = substitute_variables(text, {"foo": "bar"})
        self.assertEqual(result, text)

    def test_unresolved_blanks_remain(self):
        """Blanks that don't match any contextual pattern stay visible."""
        text = "Signature: ____"
        result = substitute_variables(text, {})
        self.assertIn("____", result)

    def test_multiple_placeholders(self):
        """Multiple {key} placeholders in one string are all replaced."""
        text = "{propounding} propounds to {responding}."
        result = substitute_variables(
            text, {"propounding": "DEF", "responding": "PLT"}
        )
        self.assertEqual(result, "DEF propounds to PLT.")

    def test_missing_named_placeholder_left_as_is(self):
        """A {key} with no matching dict entry is left unchanged."""
        text = "Hello {unknown_key} world"
        result = substitute_variables(text, {})
        self.assertEqual(result, "Hello {unknown_key} world")


# ---------------------------------------------------------------------------
# TestExtractRequestsFromText
# ---------------------------------------------------------------------------

class TestExtractRequestsFromText(unittest.TestCase):
    """Tests for the extract_requests_from_text function."""

    def test_extract_si_requests(self):
        """Parses SI-formatted text into numbered DiscoveryRequest objects."""
        text = (
            "SPECIAL INTERROGATORIES, SET ONE\n"
            "SPECIAL INTERROGATORY NO. 1:\n"
            "State your full legal name.\n"
            "\n"
            "SPECIAL INTERROGATORY NO. 2:\n"
            "Identify all witnesses.\n"
            "(The term \"IDENTIFY\" means to provide full name and address.)\n"
            "\n"
            "SPECIAL INTERROGATORY NO. 3:\n"
            "Describe the incident.\n"
        )
        requests = extract_requests_from_text(text, DiscoveryType.SI)
        self.assertEqual(len(requests), 3)
        self.assertEqual(requests[0].number, 1)
        self.assertEqual(requests[1].number, 2)
        self.assertEqual(requests[2].number, 3)

    def test_si_request_numbers(self):
        """Request numbers are correctly parsed from headers."""
        text = (
            "SPECIAL INTERROGATORY NO. 10:\n"
            "First question.\n"
            "\n"
            "SPECIAL INTERROGATORY NO. 11:\n"
            "Second question.\n"
        )
        requests = extract_requests_from_text(text, DiscoveryType.SI)
        self.assertEqual(requests[0].number, 10)
        self.assertEqual(requests[1].number, 11)

    def test_si_request_text_content(self):
        """Request body text is captured correctly."""
        text = (
            "SPECIAL INTERROGATORY NO. 1:\n"
            "If YOU contend that DEFENDANT was negligent, state all facts.\n"
        )
        requests = extract_requests_from_text(text, DiscoveryType.SI)
        self.assertEqual(len(requests), 1)
        self.assertIn("contend that DEFENDANT", requests[0].text)

    def test_preserves_inline_definitions(self):
        """Lines starting with ( after the request body are inline definitions."""
        text = (
            "SPECIAL INTERROGATORY NO. 1:\n"
            "State all facts about the INCIDENT.\n"
            '(The term "INCIDENT" refers to the accident of October 11.)\n'
            '(The term "DEFENDANT" refers to ACME CORP.)\n'
        )
        requests = extract_requests_from_text(text, DiscoveryType.SI)
        self.assertEqual(len(requests), 1)
        self.assertEqual(len(requests[0].definitions), 2)
        self.assertIn("INCIDENT", requests[0].definitions[0])
        self.assertIn("DEFENDANT", requests[0].definitions[1])

    def test_skips_empty_and_slash_lines(self):
        """Empty lines and /// lines are skipped during parsing."""
        text = (
            "SPECIAL INTERROGATORY NO. 1:\n"
            "State your name.\n"
            "\n"
            "///\n"
            "///\n"
            "\n"
            "SPECIAL INTERROGATORY NO. 2:\n"
            "Identify all persons.\n"
        )
        requests = extract_requests_from_text(text, DiscoveryType.SI)
        self.assertEqual(len(requests), 2)
        # The body of request 1 should not contain ///
        self.assertNotIn("///", requests[0].text)

    def test_extract_rpd_requests_with_headers(self):
        """RPD requests with explicit numbered headers are parsed correctly."""
        text = (
            "REQUESTS FOR PRODUCTION OF DOCUMENTS, SET ONE\n"
            "\n"
            "REQUEST FOR PRODUCTION NO. 1:\n"
            "Produce all documents relating to the incident.\n"
            '("DOCUMENT(S)" means any writing as defined by Evidence Code.)\n'
            "\n"
            "REQUEST FOR PRODUCTION NO. 2:\n"
            "Produce all photographs of the scene.\n"
        )
        requests = extract_requests_from_text(text, DiscoveryType.RPD)
        self.assertEqual(len(requests), 2)
        self.assertEqual(requests[0].number, 1)
        self.assertEqual(requests[1].number, 2)
        self.assertEqual(len(requests[0].definitions), 1)

    def test_extract_rpd_requests_without_headers(self):
        """RPD requests without numbered headers are auto-numbered sequentially."""
        text = (
            "REQUEST FOR PRODUCTION OF DOCUMENTS, SET ONE\n"
            "\n"
            "Produce all documents relating to the incident.\n"
            '("DOCUMENT(S)" means any writing as defined by Evidence Code.)\n'
            "\n"
            "Produce all photographs of the scene.\n"
            "\n"
            "Produce medical records from all providers.\n"
            '("MEDICAL RECORDS" means records of care, treatment.)\n'
        )
        requests = extract_requests_from_text(text, DiscoveryType.RPD)
        self.assertEqual(len(requests), 3)
        self.assertEqual(requests[0].number, 1)
        self.assertEqual(requests[1].number, 2)
        self.assertEqual(requests[2].number, 3)
        self.assertEqual(len(requests[0].definitions), 1)
        self.assertEqual(len(requests[2].definitions), 1)

    def test_empty_text_returns_empty_list(self):
        """Empty or whitespace-only text returns no requests."""
        self.assertEqual(extract_requests_from_text("", DiscoveryType.SI), [])
        self.assertEqual(extract_requests_from_text("   \n\n  ", DiscoveryType.SI), [])

    def test_extract_rfa_requests(self):
        """RFA requests parse with the correct pattern."""
        text = (
            "REQUEST FOR ADMISSION NO. 1:\n"
            "Admit that you were present at the scene.\n"
            "\n"
            "REQUEST FOR ADMISSION NO. 2:\n"
            "Admit that the document is authentic.\n"
        )
        requests = extract_requests_from_text(text, DiscoveryType.RFA)
        self.assertEqual(len(requests), 2)
        self.assertEqual(requests[0].number, 1)
        self.assertIn("present at the scene", requests[0].text)


# ---------------------------------------------------------------------------
# TestTemplateLoader (requires actual template files)
# ---------------------------------------------------------------------------

NEGLIGENCE_SI_PATH = os.path.join(
    DISCOVERY_DIR, "Standard Negligence Discovery (5800.070)", "SI(1) tPltf.docx"
)
NEGLIGENCE_RPD_PATH = os.path.join(
    DISCOVERY_DIR, "Standard Negligence Discovery (5800.070)", "RPD(1) tPltf.docx"
)
DEFINED_TERMS_PATH = os.path.join(DISCOVERY_DIR, "DISCOVERY DEFINED TERMS.docx")

TEMPLATES_AVAILABLE = (
    os.path.isfile(NEGLIGENCE_SI_PATH)
    and os.path.isfile(NEGLIGENCE_RPD_PATH)
    and os.path.isfile(DEFINED_TERMS_PATH)
)


@unittest.skipUnless(TEMPLATES_AVAILABLE, "Template files not found in discovery/")
class TestTemplateLoader(unittest.TestCase):
    """Integration tests that load actual .docx template files."""

    @classmethod
    def setUpClass(cls):
        cls.loader = TemplateLoader(DISCOVERY_DIR)

    def test_load_standard_negligence_si(self):
        """Standard negligence SI template produces > 50 requests."""
        requests = self.loader.load_standard_requests("negligence", DiscoveryType.SI)
        self.assertGreater(len(requests), 50)
        # Verify they are numbered
        self.assertEqual(requests[0].number, 1)
        # Numbers should be sequential (or at least non-zero)
        for req in requests:
            self.assertGreater(req.number, 0)
            self.assertTrue(len(req.text) > 0, f"Request {req.number} has empty text")

    def test_load_standard_negligence_rpd(self):
        """Standard negligence RPD template produces > 20 requests."""
        requests = self.loader.load_standard_requests("negligence", DiscoveryType.RPD)
        self.assertGreater(len(requests), 20)
        # Verify they are numbered
        self.assertEqual(requests[0].number, 1)

    def test_load_defined_terms(self):
        """Defined terms document contains expected legal term definitions."""
        terms = self.loader.load_defined_terms()
        self.assertIsInstance(terms, str)
        self.assertIn("DOCUMENT", terms)
        self.assertIn("PERSON", terms)
        # The defined terms doc uses blank placeholders for party names
        self.assertIn("____", terms)

    def test_load_si_instructions(self):
        """Instructions section is extracted from SI template."""
        instructions = self.loader.load_instructions(
            DiscoveryType.SI,
            os.path.join(
                DISCOVERY_DIR,
                "Standard Negligence Discovery (5800.070)",
                "SI(1) tPltf.docx",
            ),
        )
        self.assertIsInstance(instructions, str)
        self.assertIn("INSTRUCTIONS TO ANSWERING PARTY", instructions)
        # Should NOT contain actual interrogatory text
        self.assertNotIn("SPECIAL INTERROGATORY NO.", instructions)

    def test_load_rpd_instructions(self):
        """Instructions section is extracted from RPD template."""
        instructions = self.loader.load_instructions(
            DiscoveryType.RPD,
            os.path.join(
                DISCOVERY_DIR,
                "Standard Negligence Discovery (5800.070)",
                "RPD(1) tPltf.docx",
            ),
        )
        self.assertIsInstance(instructions, str)
        self.assertIn("INSTRUCTIONS TO ANSWERING PARTY", instructions)

    def test_load_preamble(self):
        """Preamble 'TO ... ATTORNEYS OF RECORD' block is extracted."""
        preamble = self.loader.load_preamble(
            DiscoveryType.SI,
            os.path.join(
                DISCOVERY_DIR,
                "Standard Negligence Discovery (5800.070)",
                "SI(1) tPltf.docx",
            ),
        )
        self.assertIsInstance(preamble, str)
        self.assertIn("ATTORNEYS OF RECORD", preamble)

    def test_si_requests_have_definitions(self):
        """Some SI requests should have inline definitions."""
        requests = self.loader.load_standard_requests("negligence", DiscoveryType.SI)
        has_defs = [r for r in requests if r.definitions]
        self.assertGreater(len(has_defs), 0, "Expected some requests with definitions")


if __name__ == "__main__":
    unittest.main()
