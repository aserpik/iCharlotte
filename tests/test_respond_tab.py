"""Tests for the RespondTab UI widget."""
import os
import sys
import unittest
from unittest.mock import patch, MagicMock

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from PySide6.QtWidgets import QApplication
app = QApplication.instance() or QApplication(sys.argv)

from icharlotte_core.ui.respond_tab import RespondTab


class TestRespondTabCreation(unittest.TestCase):

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    def test_creates_without_error(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        tab = RespondTab()
        self.assertIsNotNone(tab)

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    def test_has_two_document_lists(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        tab = RespondTab()
        self.assertIsNotNone(tab.discovery_list)
        self.assertIsNotNone(tab.context_list)

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    def test_has_generate_button(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        tab = RespondTab()
        self.assertEqual(tab.generate_btn.text(), "Generate Responses")

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    def test_has_rules_button(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        tab = RespondTab()
        self.assertIsNotNone(tab.rules_btn)

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    def test_has_refresh_17_1_button(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        tab = RespondTab()
        self.assertIsNotNone(tab.refresh_17_1_btn)
        self.assertFalse(tab.refresh_17_1_btn.isEnabled())

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    def test_empty_state_shown_initially(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        tab = RespondTab()
        self.assertFalse(tab.empty_label.isHidden())


class TestRespondTabLoadCase(unittest.TestCase):

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    @patch("icharlotte_core.ui.respond_tab.CaseDataManager")
    def test_load_case_sets_file_number(self, mock_cdm_cls, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        mock_cdm = MagicMock()
        mock_cdm.get_value.return_value = None
        mock_cdm_cls.return_value = mock_cdm

        tab = RespondTab()
        tab.load_case("1234.001")
        self.assertEqual(tab.file_number, "1234.001")

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    @patch("Scripts.case_data_manager.CaseDataManager")
    def test_load_case_normalizes_stale_fixed_fi_rules(self, mock_cdm_cls, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        mock_cdm = MagicMock()
        mock_cdm.get_value.return_value = {
            "fi_15_1_response": "Old saved 15.1 response.",
            "fi_objections_by_number": {
                "12.1": "Old saved 12.1 objection.",
            },
        }
        mock_cdm_cls.return_value = mock_cdm

        tab = RespondTab()
        tab.load_case("1234.001")

        self.assertTrue(mock_cdm.get_value.called)
        self.assertIn("equally available", tab.rules.fi_objections_by_number["12.1"])
        self.assertEqual(
            tab.rules.fi_responses_by_number["15.1"],
            (
                "A general denial is interposed as a matter of right based in part on "
                "California Code of Civil Procedure § 431.30. As to affirmative defenses, "
                "this interrogatory is premature at this time."
            ),
        )


if __name__ == "__main__":
    unittest.main()
