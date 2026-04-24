"""Tests for icharlotte_core.med_record_extractor._parse_page_no."""

import unittest

from icharlotte_core.med_record_extractor import _parse_page_no


class TestParsePageNo(unittest.TestCase):
    def test_lowercase_pg_no_single_page(self):
        filename, start, end = _parse_page_no("Exhibits\n\nPg no: 112/170")
        self.assertEqual(filename, "Exhibits")
        self.assertEqual(start, 112)
        self.assertEqual(end, 112)

    def test_lowercase_pg_no_page_range(self):
        filename, start, end = _parse_page_no("Exhibits\n\nPg no: 27-28/170")
        self.assertEqual(filename, "Exhibits")
        self.assertEqual(start, 27)
        self.assertEqual(end, 28)

    def test_lowercase_pg_no_multi_page_range(self):
        filename, start, end = _parse_page_no("Exhibits\n\nPg no: 161-163/170")
        self.assertEqual(filename, "Exhibits")
        self.assertEqual(start, 161)
        self.assertEqual(end, 163)

    def test_capital_pg_no_still_works(self):
        filename, start, end = _parse_page_no(
            "SALTARELLI000001-SALTARELLI000772\n\nPg. No: 501-505/772"
        )
        self.assertEqual(filename, "SALTARELLI000001-SALTARELLI000772")
        self.assertEqual(start, 501)
        self.assertEqual(end, 505)

    def test_pg_colon_form_still_works(self):
        filename, start, end = _parse_page_no(
            "60337-0014_ ORTHOPEDICS & SPORTS MEDICINE.pdf\n\nPg: 23-24/69"
        )
        self.assertEqual(filename, "60337-0014_ ORTHOPEDICS & SPORTS MEDICINE.pdf")
        self.assertEqual(start, 23)
        self.assertEqual(end, 24)

    def test_bare_number_slash_total_still_works(self):
        filename, start, end = _parse_page_no("homedepot\n\n93/535")
        self.assertEqual(filename, "homedepot")
        self.assertEqual(start, 93)
        self.assertEqual(end, 93)

    def test_malformed_trailing_dash_fails_gracefully(self):
        # "Pg no: 17-" has no end page and no /total -- unparseable.
        # Must not crash; start==0 signals caller to surface an error.
        filename, start, end = _parse_page_no(
            "Comprehensive Spine and Pain medical and billing\n\nPg no: 17-"
        )
        self.assertEqual(filename, "Comprehensive Spine and Pain medical and billing")
        self.assertEqual(start, 0)
        self.assertEqual(end, 0)


if __name__ == "__main__":
    unittest.main()
