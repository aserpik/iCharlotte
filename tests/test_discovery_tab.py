"""
Tests for the Discovery Tab UI module.

Verifies widget construction, conditional visibility, party management,
and document box behavior without requiring a live QApplication for most
tests.
"""
import os
import sys
import unittest
from unittest.mock import patch, MagicMock

# Ensure offscreen rendering for CI
os.environ["QT_QPA_PLATFORM"] = "offscreen"

from PySide6.QtWidgets import QApplication

# Ensure a QApplication exists before importing widgets
app = QApplication.instance() or QApplication(sys.argv)

from icharlotte_core.ui.discovery_tab import (
    DiscoveryTab, PropoundTab, PartyEditDialog,
)
from icharlotte_core.ui.respond_tab import RespondTab
from icharlotte_core.discovery.models import (
    Party, PartyRole, DiscoveryMode, DiscoveryType, CustomStyle,
    generate_abbreviation,
)


class TestDiscoveryTabCreation(unittest.TestCase):
    """Verify DiscoveryTab and sub-tabs can be constructed."""

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    @patch("icharlotte_core.ui.discovery_tab.ModelFetcher")
    def test_discovery_tab_creates(self, mock_propound_fetcher, mock_respond_fetcher):
        """DiscoveryTab and its PropoundTab sub-tab should instantiate."""
        for m in [mock_propound_fetcher, mock_respond_fetcher]:
            inst = MagicMock()
            inst.isRunning.return_value = False
            m.return_value = inst

        tab = DiscoveryTab()
        self.assertIsInstance(tab.propound_tab, PropoundTab)
        self.assertIsInstance(tab.respond_tab, RespondTab)

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    @patch("icharlotte_core.ui.discovery_tab.ModelFetcher")
    def test_load_case_delegates(self, mock_propound_fetcher, mock_respond_fetcher):
        for m in [mock_propound_fetcher, mock_respond_fetcher]:
            inst = MagicMock()
            inst.isRunning.return_value = False
            m.return_value = inst

        tab = DiscoveryTab()
        tab.load_case("9999.001")
        self.assertEqual(tab.propound_tab.file_number, "9999.001")
        self.assertEqual(tab.respond_tab.file_number, "9999.001")


class TestPropoundTabVisibility(unittest.TestCase):
    """Verify conditional visibility rules for mode changes."""

    @patch("icharlotte_core.ui.discovery_tab.ModelFetcher")
    def setUp(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        self.pt = PropoundTab()

    def test_initial_mode_is_standard(self):
        self.assertTrue(self.pt.rb_standard.isChecked())
        self.assertEqual(self.pt._current_mode(), DiscoveryMode.INITIAL_STANDARD)

    def test_standard_mode_visibility(self):
        """In standard mode: standard type shown, custom/llm/prompt hidden."""
        self.pt.rb_standard.setChecked(True)
        self.pt._on_mode_changed()
        # Use isHidden() — isVisible() requires the parent window to be shown
        self.assertFalse(self.pt.standard_type_group.isHidden())
        self.assertTrue(self.pt.custom_style_group.isHidden())
        self.assertTrue(self.pt.llm_group.isHidden())
        self.assertTrue(self.pt.prompt_group.isHidden())
        self.assertEqual(self.pt.generate_btn.text(), "Generate Discovery")

    def test_custom_mode_visibility(self):
        """In custom mode: custom style/llm/prompt shown, standard type hidden."""
        self.pt.rb_custom.setChecked(True)
        self.pt._on_mode_changed()
        self.assertTrue(self.pt.standard_type_group.isHidden())
        self.assertFalse(self.pt.custom_style_group.isHidden())
        self.assertFalse(self.pt.llm_group.isHidden())
        self.assertFalse(self.pt.prompt_group.isHidden())
        self.assertEqual(self.pt.generate_btn.text(), "Generate Discovery")

    def test_additional_mode_visibility(self):
        """In additional mode: llm/prompt shown, standard type/custom style hidden."""
        self.pt.rb_additional.setChecked(True)
        self.pt._on_mode_changed()
        self.assertTrue(self.pt.standard_type_group.isHidden())
        self.assertTrue(self.pt.custom_style_group.isHidden())
        self.assertFalse(self.pt.llm_group.isHidden())
        self.assertFalse(self.pt.prompt_group.isHidden())
        self.assertEqual(self.pt.generate_btn.text(), "Generate Additional Discovery")


class TestDiscoveryTypes(unittest.TestCase):

    @patch("icharlotte_core.ui.discovery_tab.ModelFetcher")
    def test_default_types(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher

        pt = PropoundTab()
        types = pt.selected_discovery_types()
        self.assertIn(DiscoveryType.SI, types)
        self.assertIn(DiscoveryType.RPD, types)
        self.assertNotIn(DiscoveryType.RFA, types)


class TestCustomStylePlaceholder(unittest.TestCase):

    @patch("icharlotte_core.ui.discovery_tab.ModelFetcher")
    def setUp(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        self.pt = PropoundTab()

    def test_custom_only_placeholder(self):
        self.pt.rb_custom_only.setChecked(True)
        self.pt._on_custom_style_changed()
        self.assertIn("what discovery", self.pt.prompt_edit.placeholderText().lower())

    def test_standard_plus_placeholder(self):
        self.pt.rb_standard_plus.setChecked(True)
        self.pt._on_custom_style_changed()
        self.assertIn("additional", self.pt.prompt_edit.placeholderText().lower())

    def test_modified_placeholder(self):
        self.pt.rb_modified.setChecked(True)
        self.pt._on_custom_style_changed()
        self.assertIn("modify", self.pt.prompt_edit.placeholderText().lower())


class TestPartyManagement(unittest.TestCase):

    @patch("icharlotte_core.ui.discovery_tab.ModelFetcher")
    def setUp(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        self.pt = PropoundTab()

    def test_add_and_refresh_parties(self):
        p1 = Party("Alice", PartyRole.PLAINTIFF, True)
        p2 = Party("Bob Corp", PartyRole.DEFENDANT, False)
        self.pt.parties = [p1, p2]
        self.pt._regenerate_abbreviations()
        self.pt._refresh_party_combo()
        self.assertEqual(self.pt.party_combo.count(), 2)
        # Check label format
        self.assertIn("Alice", self.pt.party_combo.itemText(0))
        self.assertIn("Plaintiff", self.pt.party_combo.itemText(0))

    def test_remove_party(self):
        p1 = Party("Alice", PartyRole.PLAINTIFF, True)
        p2 = Party("Bob Corp", PartyRole.DEFENDANT, False)
        self.pt.parties = [p1, p2]
        self.pt._regenerate_abbreviations()
        self.pt._refresh_party_combo()
        self.pt._remove_party(0)
        self.assertEqual(len(self.pt.parties), 1)
        self.assertEqual(self.pt.party_combo.count(), 1)


class TestPartyEditDialog(unittest.TestCase):

    def test_dialog_creates(self):
        dlg = PartyEditDialog()
        self.assertIsNotNone(dlg.name_edit)
        self.assertIsNotNone(dlg.role_combo)

    def test_dialog_prepopulated(self):
        party = Party("Test Corp", PartyRole.DEFENDANT, True)
        dlg = PartyEditDialog(party=party)
        self.assertEqual(dlg.name_edit.text(), "Test Corp")
        self.assertTrue(dlg.our_client_cb.isChecked())

    def test_get_party(self):
        dlg = PartyEditDialog()
        dlg.name_edit.setText("New Party")
        dlg.role_combo.setCurrentIndex(0)  # Plaintiff
        dlg.our_client_cb.setChecked(True)
        party = dlg.get_party()
        self.assertEqual(party.name, "New Party")
        self.assertEqual(party.role, PartyRole.PLAINTIFF)
        self.assertTrue(party.is_our_client)


class TestDocumentBox(unittest.TestCase):

    @patch("icharlotte_core.ui.discovery_tab.ModelFetcher")
    def setUp(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        self.pt = PropoundTab()

    def test_add_document(self):
        # Use a file that exists
        path = os.path.abspath(__file__)
        self.pt._add_document(path)
        self.assertEqual(self.pt.doc_list.count(), 1)

    def test_duplicate_skipped(self):
        path = os.path.abspath(__file__)
        self.pt._add_document(path)
        self.pt._add_document(path)
        self.assertEqual(self.pt.doc_list.count(), 1)


if __name__ == "__main__":
    unittest.main()
