"""Tests for the objection selector (Phase 2)."""
import unittest

from icharlotte_core.discovery.objection_selector import (
    ObjectionMenu,
    select_fi_objections,
    rule_based_preselect,
    build_objection_prompt,
    parse_objection_response,
    merge_objections,
    format_objections,
)
from icharlotte_core.discovery.response_parser import ParsedRequest
from icharlotte_core.discovery.response_rules import ResponseRules


class TestObjectionMenu(unittest.TestCase):
    def test_default_menu_has_objections(self):
        menu = ObjectionMenu.load_defaults()
        self.assertGreaterEqual(len(menu.objections), 10)

    def test_get_by_id(self):
        menu = ObjectionMenu.load_defaults()
        obj = menu.get(1)
        self.assertIn("speculation", obj.lower())
        self.assertIn("vague", obj.lower())

    def test_all_text_returns_formatted(self):
        menu = ObjectionMenu.load_defaults()
        text = menu.all_text()
        self.assertIn("1.", text)
        self.assertIn("speculation", text.lower())


class TestSelectFiObjections(unittest.TestCase):
    def test_returns_fixed_objections(self):
        rules = ResponseRules()
        objections = select_fi_objections(rules)
        self.assertIn("compound", objections.lower())
        self.assertIn("vague", objections.lower())

    def test_number_specific_fi_objections(self):
        rules = ResponseRules(fi_objections="Legacy blanket objection.")

        self.assertEqual(select_fi_objections(rules, "1.1"), "")
        self.assertIn("vague and ambiguous", select_fi_objections(rules, "15.1"))
        self.assertIn("calls for an expert opinion", select_fi_objections(rules, "16.9"))


class TestRuleBasedPreselect(unittest.TestCase):
    def test_compound_flag(self):
        req = ParsedRequest(number="1", text="State AND identify.", is_compound=True, defined_terms_used=[])
        rules = ResponseRules()
        menu = ObjectionMenu.load_defaults()
        ids = rule_based_preselect(req, rules, menu)
        self.assertIn(9, ids)  # compound objection

    def test_privacy_always_included(self):
        req = ParsedRequest(number="1", text="State your income.", is_compound=False, defined_terms_used=[])
        rules = ResponseRules(always_include_privacy_objection=True)
        menu = ObjectionMenu.load_defaults()
        ids = rule_based_preselect(req, rules, menu)
        self.assertIn(2, ids)

    def test_privilege_always_included(self):
        req = ParsedRequest(number="1", text="Any text.", is_compound=False, defined_terms_used=[])
        rules = ResponseRules(always_include_privilege_objection=True)
        menu = ObjectionMenu.load_defaults()
        ids = rule_based_preselect(req, rules, menu)
        self.assertIn(4, ids)

    def test_burden_always_included(self):
        req = ParsedRequest(number="1", text="Any text.", is_compound=False, defined_terms_used=[])
        rules = ResponseRules(always_include_burden_objection=True)
        menu = ObjectionMenu.load_defaults()
        ids = rule_based_preselect(req, rules, menu)
        self.assertIn(6, ids)

    def test_broad_definition_flag(self):
        req = ParsedRequest(number="1", text="Produce all DOCUMENTS.", is_compound=False, defined_terms_used=["DOCUMENTS"])
        rules = ResponseRules(auto_flag_broad_definitions=True)
        menu = ObjectionMenu.load_defaults()
        ids = rule_based_preselect(req, rules, menu)
        self.assertTrue(ids & {10, 11})

    def test_expert_keyword(self):
        req = ParsedRequest(number="1", text="State your expert opinion.", is_compound=False, defined_terms_used=[])
        rules = ResponseRules()
        menu = ObjectionMenu.load_defaults()
        ids = rule_based_preselect(req, rules, menu)
        self.assertIn(3, ids)

    def test_disabled_flags_not_included(self):
        req = ParsedRequest(number="1", text="Simple question.", is_compound=False, defined_terms_used=[])
        rules = ResponseRules(
            always_include_privacy_objection=False,
            always_include_privilege_objection=False,
            always_include_burden_objection=False,
        )
        menu = ObjectionMenu.load_defaults()
        ids = rule_based_preselect(req, rules, menu)
        self.assertNotIn(2, ids)
        self.assertNotIn(4, ids)
        self.assertNotIn(6, ids)


class TestMergeObjections(unittest.TestCase):
    def test_union(self):
        self.assertEqual(merge_objections({1, 2, 3}, {2, 4, 5}), {1, 2, 3, 4, 5})

    def test_empty(self):
        self.assertEqual(merge_objections(set(), {1, 2}), {1, 2})


class TestBuildObjectionPrompt(unittest.TestCase):
    def test_prompt_contains_request_and_menu(self):
        menu = ObjectionMenu.load_defaults()
        prompt = build_objection_prompt("State all facts about the INCIDENT.", menu, "aggressive")
        self.assertIn("INCIDENT", prompt)
        self.assertIn("speculation", prompt.lower())

    def test_aggressive_prompt_emphasizes_inclusion(self):
        menu = ObjectionMenu.load_defaults()
        prompt = build_objection_prompt("Test.", menu, "aggressive")
        self.assertIn("over-inclu", prompt.lower())


class TestParseObjectionResponse(unittest.TestCase):
    def test_parse_comma_list(self):
        self.assertEqual(parse_objection_response("1, 2, 4, 6"), {1, 2, 4, 6})

    def test_parse_json_array(self):
        self.assertEqual(parse_objection_response("[1, 3, 5]"), {1, 3, 5})

    def test_parse_with_text(self):
        self.assertEqual(parse_objection_response("Objections 1, 3, and 7 apply."), {1, 3, 7})


class TestFormatObjections(unittest.TestCase):
    def test_formats_multiple(self):
        menu = ObjectionMenu.load_defaults()
        text = format_objections({1, 4}, menu)
        self.assertIn("speculation", text.lower())
        self.assertIn("privilege", text.lower())

    def test_term_substitution(self):
        menu = ObjectionMenu.load_defaults()
        text = format_objections({10}, menu, term="DOCUMENTS")
        self.assertIn("DOCUMENTS", text)


if __name__ == "__main__":
    unittest.main()
