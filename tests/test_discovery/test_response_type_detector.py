import unittest

from icharlotte_core.discovery.response_type_detector import (
    DiscoveryTypeGuess,
    detect_type_from_filename,
    detect_type_from_text,
    normalize_discovery_type,
    resolve_detected_type,
)


class ResponseTypeDetectorTests(unittest.TestCase):
    def test_filename_detects_form_interrogatories(self):
        guess = detect_type_from_filename("Defendant FROGG Set One.pdf")
        self.assertEqual(guess.discovery_type, "FI")
        self.assertEqual(guess.source, "filename")

    def test_filename_detects_plural_frogs(self):
        guess = detect_type_from_filename("Plaintiff FROGS Set One.pdf")
        self.assertEqual(guess.discovery_type, "FI")

    def test_filename_detects_special_interrogatories(self):
        guess = detect_type_from_filename("Plaintiff_srogg_2.pdf")
        self.assertEqual(guess.discovery_type, "SI")

    def test_filename_detects_rfa(self):
        guess = detect_type_from_filename("RFA to Defendant.pdf")
        self.assertEqual(guess.discovery_type, "RFA")

    def test_filename_detects_rpd_from_rfp(self):
        guess = detect_type_from_filename("RFP Set 1.pdf")
        self.assertEqual(guess.discovery_type, "RPD")

    def test_text_detects_request_for_production(self):
        guess = detect_type_from_text(
            "REQUESTS FOR PRODUCTION OF DOCUMENTS, SET ONE"
        )
        self.assertEqual(guess.discovery_type, "RPD")
        self.assertEqual(guess.source, "text")

    def test_resolve_prefers_filename_when_text_absent(self):
        result = resolve_detected_type(
            filename_guess=DiscoveryTypeGuess("SI", "filename", "srogg"),
            text_guess=DiscoveryTypeGuess(None, "text", ""),
        )
        self.assertEqual(result.discovery_type, "SI")
        self.assertFalse(result.needs_user_choice)

    def test_resolve_flags_conflict(self):
        result = resolve_detected_type(
            filename_guess=DiscoveryTypeGuess("SI", "filename", "srogg"),
            text_guess=DiscoveryTypeGuess("RFA", "text", "request for admission"),
        )
        self.assertIsNone(result.discovery_type)
        self.assertTrue(result.needs_user_choice)
        self.assertIn("conflict", result.reason.lower())

    def test_resolve_needs_choice_when_no_detection(self):
        result = resolve_detected_type(
            filename_guess=DiscoveryTypeGuess(None, "filename", ""),
            text_guess=DiscoveryTypeGuess(None, "text", ""),
        )
        self.assertIsNone(result.discovery_type)
        self.assertTrue(result.needs_user_choice)

    def test_normalize_llm_discovery_type_labels(self):
        self.assertEqual(normalize_discovery_type("Form Interrogatories"), "FI")
        self.assertEqual(normalize_discovery_type("FROGS"), "FI")
        self.assertEqual(normalize_discovery_type("Special Interrogatories"), "SI")
        self.assertEqual(normalize_discovery_type("SROGS"), "SI")
        self.assertEqual(normalize_discovery_type("Requests for Admission"), "RFA")
        self.assertEqual(normalize_discovery_type("Request for Production"), "RPD")


if __name__ == "__main__":
    unittest.main()
