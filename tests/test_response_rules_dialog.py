"""Tests for the ResponseRulesDialog UI."""
import os
import sys
import unittest

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from PySide6.QtWidgets import QApplication
app = QApplication.instance() or QApplication(sys.argv)

from icharlotte_core.ui.response_rules_dialog import ResponseRulesDialog
from icharlotte_core.discovery.response_rules import ResponseRules


class TestResponseRulesDialog(unittest.TestCase):

    def test_dialog_creates(self):
        rules = ResponseRules()
        dlg = ResponseRulesDialog(rules)
        self.assertIsNotNone(dlg)

    def test_dialog_loads_rules(self):
        rules = ResponseRules(objection_aggressiveness="conservative")
        dlg = ResponseRulesDialog(rules)
        self.assertTrue(dlg.rb_conservative.isChecked())

    def test_dialog_loads_moderate(self):
        rules = ResponseRules(objection_aggressiveness="moderate")
        dlg = ResponseRulesDialog(rules)
        self.assertTrue(dlg.rb_moderate_obj.isChecked())

    def test_get_rules_returns_modified(self):
        rules = ResponseRules()
        dlg = ResponseRulesDialog(rules)
        dlg.rb_moderate_obj.setChecked(True)
        dlg.custom_instructions_edit.setPlainText("Be extra cautious")
        result = dlg.get_rules()
        self.assertEqual(result.objection_aggressiveness, "moderate")
        self.assertEqual(result.custom_instructions, "Be extra cautious")

    def test_reset_defaults(self):
        rules = ResponseRules(objection_aggressiveness="conservative", custom_instructions="test")
        dlg = ResponseRulesDialog(rules)
        dlg._on_reset_defaults()
        result = dlg.get_rules()
        self.assertEqual(result.objection_aggressiveness, "aggressive")
        self.assertEqual(result.custom_instructions, "")

    def test_three_tabs_exist(self):
        rules = ResponseRules()
        dlg = ResponseRulesDialog(rules)
        self.assertEqual(dlg.tab_widget.count(), 3)
        self.assertEqual(dlg.tab_widget.tabText(0), "Strategy")
        self.assertEqual(dlg.tab_widget.tabText(1), "Boilerplate")
        self.assertEqual(dlg.tab_widget.tabText(2), "Custom Instructions")

    def test_strategy_checkboxes(self):
        rules = ResponseRules(always_include_privacy_objection=False)
        dlg = ResponseRulesDialog(rules)
        self.assertFalse(dlg.cb_privacy.isChecked())
        dlg.cb_privacy.setChecked(True)
        result = dlg.get_rules()
        self.assertTrue(result.always_include_privacy_objection)

    def test_si_style_radio(self):
        rules = ResponseRules(si_response_style="detailed")
        dlg = ResponseRulesDialog(rules)
        self.assertTrue(dlg.rb_si_detailed.isChecked())
        dlg.rb_si_minimal.setChecked(True)
        result = dlg.get_rules()
        self.assertEqual(result.si_response_style, "minimal")

    def test_rfa_posture_radio(self):
        rules = ResponseRules(rfa_default_posture="cooperative")
        dlg = ResponseRulesDialog(rules)
        self.assertTrue(dlg.rb_rfa_cooperative.isChecked())

    def test_rpd_posture_radio(self):
        rules = ResponseRules(rpd_default_posture="unable_to_comply")
        dlg = ResponseRulesDialog(rules)
        self.assertTrue(dlg.rb_rpd_unable.isChecked())

    def test_boilerplate_fields_populated(self):
        rules = ResponseRules()
        dlg = ResponseRulesDialog(rules)
        # Check that boilerplate fields are not empty
        self.assertTrue(len(dlg.waiver_edit.toPlainText()) > 0)
        self.assertTrue(len(dlg.reservation_edit.toPlainText()) > 0)
        self.assertTrue(len(dlg.prelim_fi_edit.toPlainText()) > 0)


if __name__ == "__main__":
    unittest.main()
