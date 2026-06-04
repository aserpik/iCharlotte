"""Tests for the response drafter (Phase 3)."""
import unittest

from icharlotte_core.discovery.response_drafter import (
    format_single_response,
    is_fi_series,
    get_fi_fixed_response,
    build_si_prompt,
    build_rfa_prompt,
    build_rpd_prompt,
    build_fi_17_1_prompt,
    detect_inapplicable_fi,
    assemble_plain_text,
)
from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
from icharlotte_core.discovery.response_rules import ResponseRules


class TestFormatSingleResponse(unittest.TestCase):
    def test_fi_format(self):
        result = format_single_response(
            disc_type="FI", request_number="1.1",
            request_text="State the name...",
            objections="Objection text here.",
            substantive="Attorney info here.",
            waiver="Subject to and without waiving...",
            reservation="Discovery and investigation...",
        )
        self.assertIn("FORM INTERROGATORY NO. 1.1:", result)
        self.assertIn("RESPONSE TO FORM INTERROGATORY NO. 1.1:", result)
        self.assertIn("Objection text here.", result)
        self.assertIn("Subject to and without waiving...", result)
        self.assertIn("Attorney info here.", result)
        self.assertIn("Discovery and investigation...", result)

    def test_si_format(self):
        result = format_single_response(
            disc_type="SI", request_number="5",
            request_text="Describe the incident.",
            objections="Objection text.",
            substantive="Response here.",
            waiver="Subject to...",
            reservation="Discovery...",
        )
        self.assertIn("SPECIAL INTERROGATORY NO. 5:", result)
        self.assertIn("RESPONSE TO SPECIAL INTERROGATORY NO. 5:", result)

    def test_rfa_format(self):
        result = format_single_response(
            disc_type="RFA", request_number="3",
            request_text="Admit that...", objections="Obj.",
            substantive="Deny.", waiver="Subject to...", reservation="Disc...",
        )
        self.assertIn("REQUEST FOR ADMISSION NO. 3:", result)
        self.assertIn("RESPONSE TO REQUEST FOR ADMISSION NO. 3:", result)

    def test_rpd_format(self):
        result = format_single_response(
            disc_type="RPD", request_number="7",
            request_text="All documents...", objections="Obj.",
            substantive="Will comply.", waiver="Subject to...", reservation="Disc...",
        )
        self.assertIn("REQUEST FOR PRODUCTION NO. 7:", result)
        self.assertIn("RESPONSE TO REQUEST NO. 7:", result)

    def test_multiple_objections_render_as_one_paragraph(self):
        """All objections share one paragraph; only the substantive starts a new one.

        Objections arrive joined by a blank line (the internal block separator).
        They must collapse onto a single line so the rendered document shows every
        objection in one paragraph, with the substantive response beginning the
        next paragraph.
        """
        result = format_single_response(
            disc_type="RFA", request_number="3",
            request_text="Admit that the light was red.",
            objections="Responding Party objects on grounds A.\n\n"
                       "Responding Party further objects on grounds B.",
            substantive="Deny.",
            waiver="Subject to and without waiving the foregoing objections, "
                   "Responding Party responds as follows:",
            reservation="Discovery is ongoing.",
        )
        # No blank line splitting the two objections apart.
        self.assertNotIn("grounds A.\n\nResponding Party further objects", result)
        # Both objections sit on the same line, followed by the waiver.
        self.assertIn(
            "Responding Party objects on grounds A. "
            "Responding Party further objects on grounds B. "
            "Subject to and without waiving the foregoing objections, "
            "Responding Party responds as follows:",
            result,
        )
        # The substantive response begins on its own line.
        self.assertIn("\nDeny. Discovery is ongoing.", result)


class TestIsFiSeries(unittest.TestCase):
    def test_series_3(self):
        self.assertTrue(is_fi_series("3.1", [3]))
        self.assertTrue(is_fi_series("3.2", [3]))
        self.assertFalse(is_fi_series("3.1", [4]))

    def test_series_16(self):
        self.assertTrue(is_fi_series("16.1", [16]))
        self.assertTrue(is_fi_series("16.9", [16]))

    def test_non_series(self):
        self.assertFalse(is_fi_series("6.1", [3, 4, 7, 12, 13, 14, 20]))

    def test_multi_series(self):
        self.assertTrue(is_fi_series("7.3", [3, 4, 7]))
        self.assertTrue(is_fi_series("4.1", [3, 4, 7]))


class TestGetFiFixedResponse(unittest.TestCase):
    def test_1_1_is_fixed(self):
        rules = ResponseRules()
        resp = get_fi_fixed_response("1.1", rules)
        self.assertIsNotNone(resp)
        self.assertIn("attorneys of record", resp.lower())

    def test_15_1_is_fixed(self):
        rules = ResponseRules()
        resp = get_fi_fixed_response("15.1", rules)
        self.assertIsNotNone(resp)
        self.assertIn("general denial", resp.lower())

    def test_16_x_is_fixed(self):
        rules = ResponseRules()
        for n in ["16.1", "16.2", "16.3", "16.4", "16.5", "16.6", "16.7", "16.8", "16.9", "16.10"]:
            resp = get_fi_fixed_response(n, rules)
            self.assertIsNotNone(resp, f"Expected fixed response for {n}")
            self.assertEqual(resp, "")

    def test_17_1_is_placeholder(self):
        rules = ResponseRules()
        resp = get_fi_fixed_response("17.1", rules)
        self.assertIsNotNone(resp)
        self.assertIn("PENDING", resp)

    def test_non_fixed_returns_none(self):
        rules = ResponseRules()
        for n in ["2.1", "5.1", "6.1", "8.1", "9.1", "10.1", "11.1", "18.1"]:
            resp = get_fi_fixed_response(n, rules)
            self.assertIsNone(resp, f"Expected None for {n}")


class TestDetectInapplicableFi(unittest.TestCase):
    def test_wrongful_death_in_negligence(self):
        self.assertTrue(detect_inapplicable_fi("10.1", case_type="negligence"))

    def test_applicable_fi(self):
        self.assertFalse(detect_inapplicable_fi("1.1", case_type="negligence"))
        self.assertFalse(detect_inapplicable_fi("3.1", case_type="negligence"))


class TestBuildPrompts(unittest.TestCase):
    def test_si_prompt_minimal(self):
        rules = ResponseRules(si_response_style="minimal")
        prompt = build_si_prompt("Describe the incident.", "Context: accident at intersection.", rules)
        self.assertIn("narrowly", prompt.lower())
        self.assertIn("few words", prompt.lower())

    def test_si_prompt_detailed(self):
        rules = ResponseRules(si_response_style="detailed")
        prompt = build_si_prompt("Describe the incident.", "Context: accident.", rules)
        self.assertNotIn("narrowly", prompt.lower())

    def test_rfa_prompt_cautious(self):
        rules = ResponseRules(rfa_default_posture="cautious")
        prompt = build_rfa_prompt("Admit defendant was negligent.", "Context: disputed.", rules)
        self.assertIn("Admit", prompt)
        self.assertIn("Deny", prompt)
        self.assertIn("insufficient", prompt.lower())
        self.assertIn("lean toward", prompt.lower())

    def test_rpd_prompt_context_dependent(self):
        rules = ResponseRules(rpd_default_posture="context_dependent")
        prompt = build_rpd_prompt("All training documents.", "Context: records available.", rules)
        self.assertIn("unable to comply", prompt.lower())
        self.assertIn("will comply", prompt.lower())

    def test_fi_17_1_prompt(self):
        prompt = build_fi_17_1_prompt("RFA 1: Admit. RFA 2: Deny.", ResponseRules())
        self.assertIn("17.1", prompt)
        self.assertIn("RFA 1: Admit", prompt)


class TestAssemblePlainText(unittest.TestCase):
    def test_si_document(self):
        parsed = ParsedDiscovery(
            discovery_type="SI",
            propounding_party="Plaintiff JOHN DOE",
            responding_party="Defendant ACME CORP",
            set_number=1, set_word="ONE", case_number="23STCV12345",
            requests=[ParsedRequest(number="1", text="Describe the incident.")],
        )
        rules = ResponseRules()
        objections_map = {"1": "Objection text."}
        responses_map = {"1": "Substantive response."}
        text = assemble_plain_text(parsed, rules, objections_map, responses_map)
        self.assertIn("PROPOUNDING PARTY:", text)
        self.assertIn("Plaintiff JOHN DOE", text)
        self.assertIn("RESPONDING PARTY:", text)
        self.assertIn("Defendant ACME CORP", text)
        self.assertIn("SET NO.:", text)
        self.assertIn("ONE (1)", text)
        self.assertIn("PRELIMINARY STATEMENT", text)
        self.assertIn("SPECIAL INTERROGATORY NO. 1:", text)
        self.assertIn("RESPONSE TO SPECIAL INTERROGATORY NO. 1:", text)
        self.assertIn("Objection text.", text)
        self.assertIn("Substantive response.", text)

    def test_rfa_includes_general_objections(self):
        parsed = ParsedDiscovery(
            discovery_type="RFA",
            propounding_party="Plaintiff X",
            responding_party="Defendant Y",
            set_number=1, set_word="ONE", case_number="123",
            requests=[ParsedRequest(number="1", text="Admit that...")],
        )
        rules = ResponseRules()
        text = assemble_plain_text(parsed, rules, {"1": "Obj."}, {"1": "Deny."})
        self.assertIn("GENERAL OBJECTIONS", text)

    def test_rpd_includes_general_objections(self):
        parsed = ParsedDiscovery(
            discovery_type="RPD",
            propounding_party="Plaintiff X",
            responding_party="Defendant Y",
            set_number=1, set_word="ONE", case_number="123",
            requests=[ParsedRequest(number="1", text="Produce all...")],
        )
        rules = ResponseRules()
        text = assemble_plain_text(parsed, rules, {"1": "Obj."}, {"1": "Will comply."})
        self.assertIn("GENERAL OBJECTIONS", text)

    def test_si_no_general_objections(self):
        parsed = ParsedDiscovery(
            discovery_type="SI",
            propounding_party="P", responding_party="D",
            set_number=1, set_word="ONE", case_number="123",
            requests=[ParsedRequest(number="1", text="Q")],
        )
        rules = ResponseRules()
        text = assemble_plain_text(parsed, rules, {"1": "O"}, {"1": "R"})
        self.assertNotIn("GENERAL OBJECTIONS", text)


if __name__ == "__main__":
    unittest.main()
