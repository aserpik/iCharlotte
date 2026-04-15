"""Tests for the Import Reports feature in ChatTab."""
import unittest


class TestCarrierReportRegex(unittest.TestCase):
    """Verify CARRIER_REPORT_RE matches the spec's must/must-not lists."""

    def setUp(self):
        from icharlotte_core.ui.tabs import CARRIER_REPORT_RE
        self.pattern = CARRIER_REPORT_RE

    def _matches(self, name: str) -> bool:
        return self.pattern.match(name) is not None

    # --- Must match -----------------------------------------------------
    def test_basic_carrier001_docx(self):
        self.assertTrue(self._matches("carrier001.docx"))

    def test_carrier015_doc_lowercase(self):
        self.assertTrue(self._matches("carrier015.doc"))

    def test_trailing_space_and_parens(self):
        self.assertTrue(self._matches("carrier002 (FSR).docx"))

    def test_trailing_parens_no_space(self):
        self.assertTrue(self._matches("carrier003(lit plan).docx"))

    def test_uppercase_carrier_and_extension(self):
        self.assertTrue(self._matches("Carrier007.DOCX"))

    def test_trailing_dash_suffix(self):
        self.assertTrue(self._matches("carrier010 - Final.docx"))

    def test_all_caps_carrier(self):
        self.assertTrue(self._matches("CARRIER005.docx"))

    # --- Must NOT match -------------------------------------------------
    def test_bracket_prefix_rejected(self):
        self.assertFalse(self._matches("[draft]carrier001.docx"))

    def test_word_prefix_rejected(self):
        self.assertFalse(self._matches("draft_carrier001.docx"))

    def test_carrier000_below_range(self):
        self.assertFalse(self._matches("carrier000.docx"))

    def test_carrier016_above_range(self):
        self.assertFalse(self._matches("carrier016.docx"))

    def test_four_digit_run_rejected(self):
        self.assertFalse(self._matches("carrier0011.docx"))

    def test_wrong_extension_pdf(self):
        self.assertFalse(self._matches("carrier001.pdf"))

    def test_no_number(self):
        self.assertFalse(self._matches("carrier.docx"))

    def test_two_digit_number(self):
        self.assertFalse(self._matches("carrier01.docx"))

    def test_carrier100_out_of_range(self):
        self.assertFalse(self._matches("carrier100.docx"))


if __name__ == "__main__":
    unittest.main()
