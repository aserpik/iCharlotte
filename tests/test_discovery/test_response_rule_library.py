import tempfile
import unittest
from pathlib import Path

from icharlotte_core.discovery.response_rule_library import (
    ALWAYS_VAGUE_TEXT,
    RuleCategory,
    RuleMode,
    ResponseRule,
    built_in_rules_for,
    get_quick_objection_rules,
    load_global_rules,
    save_global_rules,
)


class ResponseRuleLibraryTests(unittest.TestCase):
    def test_si_rules_include_mandatory_vague_rule(self):
        rules = built_in_rules_for("SI")
        vague = next(r for r in rules if r.id == "always_vague_ambiguous_overbroad")
        self.assertEqual(vague.mode, RuleMode.MANDATORY)
        self.assertEqual(vague.category, RuleCategory.OBJECTION)
        self.assertTrue(vague.enabled_by_default)
        self.assertIn("ALWAYS include", vague.name)
        self.assertEqual(vague.output_text, ALWAYS_VAGUE_TEXT)

    def test_si_rules_include_minimal_substantive_rule(self):
        rules = built_in_rules_for("SI")
        substantive = [r for r in rules if r.category == RuleCategory.SUBSTANTIVE]
        self.assertEqual([r.id for r in substantive], ["minimal_direct_answer"])

    def test_rfa_rules_only_objection_rules(self):
        rules = built_in_rules_for("RFA")
        self.assertTrue(rules)
        self.assertTrue(all(r.category == RuleCategory.OBJECTION for r in rules))

    def test_rpd_rules_only_objection_rules(self):
        rules = built_in_rules_for("RPD")
        self.assertTrue(rules)
        self.assertTrue(all(r.category == RuleCategory.OBJECTION for r in rules))

    def test_fi_fixed_has_no_default_rule_cards(self):
        self.assertEqual(built_in_rules_for("FI", fi_mode="fixed"), [])

    def test_fi_custom_uses_si_style_rules(self):
        rules = built_in_rules_for("FI", fi_mode="custom")
        self.assertTrue(any(r.id == "minimal_direct_answer" for r in rules))
        self.assertTrue(any(r.id == "always_vague_ambiguous_overbroad" for r in rules))

    def test_quick_objections_include_undefined_term_template(self):
        quick = get_quick_objection_rules()
        undefined = next(r for r in quick if r.id == "quick_undefined_term")
        self.assertIn("{term}", undefined.output_text)

    def test_global_rules_round_trip(self):
        rule = ResponseRule(
            id="custom_privacy",
            name="Custom privacy rule",
            category=RuleCategory.OBJECTION,
            mode=RuleMode.CONDITIONAL,
            applies_to=("SI",),
            description="Use for private material.",
            output_text="Custom objection.",
            enabled_by_default=False,
            is_global=True,
        )
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "rules.json"
            save_global_rules(str(path), [rule])
            loaded = load_global_rules(str(path))
        self.assertEqual(loaded, [rule])


if __name__ == "__main__":
    unittest.main()
