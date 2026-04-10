"""Tests for ResponseRules dataclass, serialization, and default loading."""
import json
import os
import unittest
import tempfile

from icharlotte_core.discovery.response_rules import ResponseRules


class TestResponseRules(unittest.TestCase):

    def test_create_with_defaults(self):
        rules = ResponseRules()
        self.assertEqual(rules.objection_aggressiveness, "aggressive")
        self.assertTrue(rules.always_include_privacy_objection)
        self.assertTrue(rules.always_include_privilege_objection)
        self.assertTrue(rules.always_include_burden_objection)
        self.assertTrue(rules.auto_flag_compound)
        self.assertTrue(rules.auto_flag_broad_definitions)
        self.assertEqual(rules.si_response_style, "minimal")
        self.assertEqual(rules.rfa_default_posture, "cautious")
        self.assertEqual(rules.rpd_default_posture, "context_dependent")
        self.assertFalse(rules.fi_17_1_auto_refresh)
        self.assertIn("Subject to and without waiving", rules.waiver_language)
        self.assertIn("Discovery and investigation are ongoing", rules.reservation_clause)
        self.assertEqual(rules.custom_instructions, "")

    def test_to_dict_roundtrip(self):
        rules = ResponseRules()
        rules.custom_instructions = "Test custom instruction"
        d = rules.to_dict()
        restored = ResponseRules.from_dict(d)
        self.assertEqual(rules.to_dict(), restored.to_dict())

    def test_from_dict_partial(self):
        partial = {"objection_aggressiveness": "conservative", "custom_instructions": "be brief"}
        rules = ResponseRules.from_dict(partial)
        self.assertEqual(rules.objection_aggressiveness, "conservative")
        self.assertEqual(rules.custom_instructions, "be brief")
        self.assertTrue(rules.always_include_privacy_objection)
        self.assertEqual(rules.si_response_style, "minimal")

    def test_save_and_load_json(self):
        rules = ResponseRules()
        rules.objection_aggressiveness = "moderate"
        with tempfile.NamedTemporaryFile(suffix=".json", delete=False, mode="w") as f:
            path = f.name
        try:
            rules.save_to_json(path)
            loaded = ResponseRules.load_from_json(path)
            self.assertEqual(loaded.objection_aggressiveness, "moderate")
            self.assertEqual(loaded.waiver_language, rules.waiver_language)
        finally:
            os.unlink(path)

    def test_preliminary_statements_exist_for_all_types(self):
        rules = ResponseRules()
        self.assertTrue(len(rules.preliminary_statement_fi) > 100)
        self.assertTrue(len(rules.preliminary_statement_si) > 100)
        self.assertTrue(len(rules.preliminary_statement_rfa) > 100)
        self.assertTrue(len(rules.preliminary_statement_rpd) > 100)

    def test_intro_templates_have_placeholders(self):
        rules = ResponseRules()
        self.assertIn("{responding_party}", rules.intro_template_fi)
        self.assertIn("{propounding_party}", rules.intro_template_fi)
        self.assertIn("{set_word_title}", rules.intro_template_fi)

    def test_general_objections_only_for_rfa_rpd(self):
        rules = ResponseRules()
        self.assertTrue(len(rules.general_objections_rfa) > 100)
        self.assertTrue(len(rules.general_objections_rpd) > 100)

    def test_verification_template_exists(self):
        """Verification template contains {verifier_name} and {document_title} placeholders."""
        rules = ResponseRules()
        self.assertIn("{document_title}", rules.verification_template)
        self.assertIn("{verifier_name}", rules.verification_template)


if __name__ == "__main__":
    unittest.main()
