"""Tests for Scripts/med_record.py PDF extraction helpers.

Covers the OCR preprocessing, render-DPI calculation, page anchoring, and
failure-reporting logic added to Pass 1 of the med_record agent.
"""
import os
import sys
import unittest
from unittest.mock import MagicMock

# Make Scripts/ importable
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "Scripts"))

import med_record as mr
from PIL import Image


class TestComputeOcrZoom(unittest.TestCase):
    """_compute_ocr_zoom should target ~300 dpi but cap memory on oversized pages."""

    def _page_with_size(self, width_pts, height_pts):
        page = MagicMock()
        page.rect.width = width_pts
        page.rect.height = height_pts
        return page

    def test_us_letter_targets_300_dpi(self):
        # US Letter: 612 x 792 points (8.5" x 11" at 72 dpi)
        page = self._page_with_size(612, 792)
        zoom = mr._compute_ocr_zoom(page)
        expected = mr.OCR_TARGET_DPI / mr.PDF_POINTS_PER_INCH  # ~4.17
        self.assertAlmostEqual(zoom, expected, places=2)

    def test_oversized_page_capped_by_pixel_dim(self):
        # 17" x 22" engineering drawing = 1224 x 1584 points
        # At target zoom 4.17: 1584 * 4.17 = 6605 px — too big.
        # Cap = OCR_MAX_PIXEL_DIM / longest_dim_pts
        page = self._page_with_size(1224, 1584)
        zoom = mr._compute_ocr_zoom(page)
        cap = mr.OCR_MAX_PIXEL_DIM / 1584
        self.assertAlmostEqual(zoom, cap, places=2)
        self.assertLess(zoom, mr.OCR_TARGET_DPI / mr.PDF_POINTS_PER_INCH)

    def test_tiny_page_does_not_overzoom(self):
        # Small 4" x 6" form = 288 x 432 points
        # Pixel-dim cap would allow ~10x zoom, but we cap at target dpi.
        page = self._page_with_size(288, 432)
        zoom = mr._compute_ocr_zoom(page)
        self.assertAlmostEqual(zoom, mr.OCR_TARGET_DPI / mr.PDF_POINTS_PER_INCH, places=2)

    def test_zero_dimensions_falls_back_to_target(self):
        page = self._page_with_size(0, 0)
        zoom = mr._compute_ocr_zoom(page)
        # No dim info -> just use target, no cap
        self.assertGreaterEqual(zoom, mr.OCR_MIN_DPI / mr.PDF_POINTS_PER_INCH)

    def test_zoom_never_below_min_dpi(self):
        # Pathological huge page that would push cap below min
        page = self._page_with_size(100000, 100000)
        zoom = mr._compute_ocr_zoom(page)
        self.assertGreaterEqual(zoom, mr.OCR_MIN_DPI / mr.PDF_POINTS_PER_INCH)


class TestPreprocessForOcr(unittest.TestCase):
    def test_rgb_input_returned_as_grayscale(self):
        img = Image.new("RGB", (50, 50), color=(120, 80, 200))
        out = mr._preprocess_for_ocr(img)
        self.assertEqual(out.mode, "L")
        self.assertEqual(out.size, (50, 50))

    def test_grayscale_input_passes_through(self):
        img = Image.new("L", (40, 40), color=128)
        out = mr._preprocess_for_ocr(img)
        self.assertEqual(out.mode, "L")

    def test_low_contrast_image_gets_stretched(self):
        # Build a low-contrast image where pixel values span only 100..160
        # (a dense band, not just two outliers — autocontrast w/ cutoff=1 would
        # clip away sparse outliers). After preprocessing, the band should
        # stretch to roughly the full 0..255 range.
        img = Image.new("L", (30, 30))
        for y in range(30):
            for x in range(30):
                # Values cycle through 100..160 evenly across the image
                img.putpixel((x, y), 100 + ((x + y) * 60 // 58))
        before = img.getextrema()
        self.assertEqual(before, (100, 160))  # sanity: test image is what we expect

        out = mr._preprocess_for_ocr(img)
        after = out.getextrema()
        # Autocontrast should stretch the band noticeably wider
        self.assertLess(after[0], 100)
        self.assertGreater(after[1], 160)


class TestFormatPageBlock(unittest.TestCase):
    def test_normal_page_wraps_in_anchor(self):
        out = mr._format_page_block(5, "Patient seen on 1/1/24.", None)
        self.assertIn("=== PAGE 5 ===", out)
        self.assertIn("Patient seen on 1/1/24.", out)
        self.assertNotIn("FAILED", out)
        self.assertNotIn("EMPTY", out)

    def test_failed_page_reports_reason(self):
        out = mr._format_page_block(7, "", "PDF parser error: bad xref")
        self.assertIn("=== PAGE 7 [EXTRACTION FAILED: PDF parser error: bad xref] ===", out)

    def test_failed_page_with_text_still_marked_failed(self):
        # Partial text + error → still treated as a failure for transparency
        out = mr._format_page_block(8, "partial", "memory limit hit")
        self.assertIn("FAILED", out)

    def test_empty_text_marked_empty(self):
        out = mr._format_page_block(10, "", None)
        self.assertIn("=== PAGE 10 [EMPTY] ===", out)

    def test_whitespace_only_treated_as_empty(self):
        out = mr._format_page_block(11, "   \n  \t  ", None)
        self.assertIn("EMPTY", out)


class TestAssemblePages(unittest.TestCase):
    def test_pages_ordered_by_number(self):
        results = {
            3: ("Page three content", None),
            1: ("Page one content", None),
            2: ("Page two content", None),
        }
        out = mr._assemble_pages(results, 3)
        i1 = out.index("Page one content")
        i2 = out.index("Page two content")
        i3 = out.index("Page three content")
        self.assertLess(i1, i2)
        self.assertLess(i2, i3)

    def test_all_pages_have_anchor(self):
        results = {1: ("A", None), 2: ("B", None), 3: ("C", None)}
        out = mr._assemble_pages(results, 3)
        for i in (1, 2, 3):
            self.assertIn(f"=== PAGE {i} ===", out)

    def test_missing_page_reported_as_failed(self):
        # Page 2 never made it into the dict (e.g., memory abort)
        results = {1: ("Page one", None), 3: ("Page three", None)}
        out = mr._assemble_pages(results, 3)
        self.assertIn("=== PAGE 2 [EXTRACTION FAILED:", out)
        # Other pages still present
        self.assertIn("Page one", out)
        self.assertIn("Page three", out)

    def test_failed_page_reason_appears(self):
        results = {
            1: ("ok content", None),
            2: ("", "memory limit exceeded"),
            3: ("more content", None),
        }
        out = mr._assemble_pages(results, 3)
        self.assertIn("=== PAGE 2 [EXTRACTION FAILED: memory limit exceeded] ===", out)

    def test_empty_page_marked_empty_not_failed(self):
        results = {1: ("text", None), 2: ("", None), 3: ("more", None)}
        out = mr._assemble_pages(results, 3)
        self.assertIn("=== PAGE 2 [EMPTY] ===", out)
        self.assertNotIn("FAILED", out)


if __name__ == "__main__":
    unittest.main()
