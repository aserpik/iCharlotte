"""Tests for the discovery response parser (Phase 1)."""
import unittest

from icharlotte_core.discovery.response_parser import (
    ParsedDiscovery,
    ParsedRequest,
    detect_discovery_type,
    detect_compound,
    extract_defined_terms,
    build_parse_prompt,
    parse_llm_response,
)


class TestDetectDiscoveryType(unittest.TestCase):
    def test_detect_fi(self):
        text = "FORM INTERROGATORY NO. 1.1:\nState the name..."
        self.assertEqual(detect_discovery_type(text), "FI")

    def test_detect_si(self):
        text = "SPECIAL INTERROGATORY NO. 1:\nDescribe in detail..."
        self.assertEqual(detect_discovery_type(text), "SI")

    def test_detect_rfa(self):
        text = "REQUEST FOR ADMISSION NO. 1:\nAdmit that..."
        self.assertEqual(detect_discovery_type(text), "RFA")

    def test_detect_rpd(self):
        text = "REQUEST FOR PRODUCTION NO. 1:\nAll documents..."
        self.assertEqual(detect_discovery_type(text), "RPD")

    def test_detect_unknown_returns_none(self):
        text = "This is some random legal document text."
        self.assertIsNone(detect_discovery_type(text))


class TestDetectCompound(unittest.TestCase):
    def test_simple_question_not_compound(self):
        text = "State the name of each witness."
        self.assertFalse(detect_compound(text))

    def test_and_conjunction_compound(self):
        text = "State all facts AND identify all documents."
        self.assertTrue(detect_compound(text))

    def test_multiple_subparts_not_compound(self):
        text = "State: (a) the name; (b) the address; (c) the phone number."
        self.assertFalse(detect_compound(text))

    def test_multiple_action_verbs_compound(self):
        text = "Identify each person and describe the basis for your contention and state all facts supporting your claim."
        self.assertTrue(detect_compound(text))


class TestExtractDefinedTerms(unittest.TestCase):
    def test_all_caps_terms(self):
        text = "Describe the INCIDENT involving the VEHICLE."
        terms = extract_defined_terms(text)
        self.assertIn("INCIDENT", terms)
        self.assertIn("VEHICLE", terms)

    def test_ignores_short_caps(self):
        text = "State if YOU are A corporation."
        terms = extract_defined_terms(text)
        self.assertIn("YOU", terms)
        self.assertNotIn("A", terms)

    def test_no_defined_terms(self):
        text = "State the name of the witness."
        terms = extract_defined_terms(text)
        self.assertEqual(terms, [])


class TestBuildParsePrompt(unittest.TestCase):
    def test_prompt_contains_instructions(self):
        prompt = build_parse_prompt("Some discovery text here")
        self.assertIn("discovery type", prompt.lower())
        self.assertIn("propounding party", prompt.lower())
        self.assertIn("JSON", prompt)

    def test_prompt_includes_document_text(self):
        prompt = build_parse_prompt("SPECIAL INTERROGATORY NO. 1: Describe...")
        self.assertIn("SPECIAL INTERROGATORY NO. 1", prompt)

    def test_prompt_maps_judicial_council_caption_labels(self):
        # Judicial Council form interrogatories (DISC-001/005) label the
        # caption "Asking Party" (= propounding) and "Answering Party"
        # (= responding). PDF extraction often separates these labels from
        # their values, so the prompt must define the mapping explicitly.
        prompt = build_parse_prompt("Some discovery text here").lower()
        self.assertIn("asking party", prompt)
        self.assertIn("answering party", prompt)

    def test_prompt_excludes_attorney_for_line_from_responding_party(self):
        # The "Attorney For" line names the propounding side's client; it must
        # not be mistaken for the responding party.
        prompt = build_parse_prompt("Some discovery text here").lower()
        self.assertIn("attorney for", prompt)


class TestParseLlmResponse(unittest.TestCase):
    def test_valid_json_response(self):
        llm_json = '''{
            "discovery_type": "SI",
            "propounding_party": "Plaintiff JOHN DOE",
            "set_number": 1,
            "case_number": "23STCV12345",
            "requests": [
                {"number": "1", "text": "Describe in detail how the INCIDENT occurred."},
                {"number": "2", "text": "State all facts supporting your contention AND identify all witnesses."}
            ]
        }'''
        parsed = parse_llm_response(llm_json, our_client_name="Defendant ACME Corp")
        self.assertEqual(parsed.discovery_type, "SI")
        self.assertEqual(parsed.propounding_party, "Plaintiff JOHN DOE")
        self.assertEqual(parsed.responding_party, "Defendant ACME Corp")
        self.assertEqual(parsed.set_number, 1)
        self.assertEqual(parsed.case_number, "23STCV12345")
        self.assertEqual(len(parsed.requests), 2)
        self.assertFalse(parsed.requests[0].is_compound)
        self.assertTrue(parsed.requests[1].is_compound)
        self.assertIn("INCIDENT", parsed.requests[0].defined_terms_used)

    def test_malformed_json_raises(self):
        with self.assertRaises(ValueError):
            parse_llm_response("not valid json", our_client_name="Test")

    def test_markdown_fenced_json(self):
        llm_json = '```json\n{"discovery_type": "RFA", "propounding_party": "Plaintiff X", "set_number": 1, "case_number": "123", "requests": []}\n```'
        parsed = parse_llm_response(llm_json, our_client_name="Defendant Y")
        self.assertEqual(parsed.discovery_type, "RFA")

    def test_normalizes_verbose_discovery_type_label(self):
        llm_json = '''{
            "discovery_type": "Form Interrogatories",
            "propounding_party": "Plaintiff X",
            "responding_party": "Defendant Y",
            "set_number": 1,
            "case_number": "123",
            "requests": []
        }'''
        parsed = parse_llm_response(llm_json)
        self.assertEqual(parsed.discovery_type, "FI")


if __name__ == "__main__":
    unittest.main()
