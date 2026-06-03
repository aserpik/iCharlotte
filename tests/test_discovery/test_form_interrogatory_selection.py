import unittest

import fitz

from icharlotte_core.discovery.form_interrogatory_selection import (
    complete_selected_form_interrogatories,
    extract_selected_form_interrogatory_numbers,
    filter_parsed_form_interrogatories,
    scan_form_interrogatories,
)
from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest


class FormInterrogatorySelectionTests(unittest.TestCase):
    def test_extracts_selected_flattened_checkbox_numbers(self):
        doc = fitz.open()
        page = doc.new_page(width=612, height=792)

        _draw_fi_row(page, "1.1", 54, selected=True)
        _draw_fi_row(page, "3.1", 104, selected=True)
        _draw_fi_row(page, "6.1", 154, selected=False)

        path = "synthetic_frog_selection.pdf"
        try:
            doc.save(path)
            doc.close()

            selected = extract_selected_form_interrogatory_numbers(path)
        finally:
            if not doc.is_closed:
                doc.close()
            import os
            if os.path.exists(path):
                os.remove(path)

        self.assertEqual(selected, ["1.1", "3.1"])

    def test_filters_parsed_fi_requests_to_selected_numbers(self):
        parsed = ParsedDiscovery(
            discovery_type="Form Interrogatories",
            propounding_party="Plaintiff",
            responding_party="Defendant",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[
                ParsedRequest(number="1.1", text="State who answered."),
                ParsedRequest(number="3.1", text="Are you a corporation?"),
                ParsedRequest(number="6.1", text="Do you claim injuries?"),
            ],
        )

        filtered = filter_parsed_form_interrogatories(parsed, ["1.1", "3.1"])

        self.assertEqual(filtered.discovery_type, "FI")
        self.assertEqual([req.number for req in filtered.requests], ["1.1", "3.1"])

    def test_filter_leaves_requests_when_no_selected_numbers_available(self):
        parsed = ParsedDiscovery(
            discovery_type="FI",
            propounding_party="Plaintiff",
            responding_party="Defendant",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[ParsedRequest(number="6.1", text="Do you claim injuries?")],
        )

        filtered = filter_parsed_form_interrogatories(parsed, [])

        self.assertEqual([req.number for req in filtered.requests], ["6.1"])

    def test_completes_empty_fi_parse_from_selected_flattened_checkboxes(self):
        doc = fitz.open()
        page = doc.new_page(width=612, height=792)

        _draw_fi_row(page, "1.1", 54, selected=True, text="State who answered.")
        _draw_fi_row(page, "3.1", 104, selected=True, text="Are you a corporation?")
        _draw_fi_row(page, "6.1", 154, selected=False, text="Do you claim injuries?")

        path = "synthetic_frog_completion.pdf"
        try:
            doc.save(path)
            doc.close()
            parsed = ParsedDiscovery(
                discovery_type="Form Interrogatories",
                propounding_party="Plaintiff",
                responding_party="Defendant",
                set_number=1,
                set_word="ONE",
                case_number="123",
                requests=[],
            )

            completed = complete_selected_form_interrogatories(parsed, path)
        finally:
            if not doc.is_closed:
                doc.close()
            import os
            if os.path.exists(path):
                os.remove(path)

        self.assertEqual(completed.discovery_type, "FI")
        self.assertEqual([req.number for req in completed.requests], ["1.1", "3.1"])
        self.assertIn("State who answered", completed.requests[0].text)
        self.assertIn("Are you a corporation", completed.requests[1].text)

    def test_complete_prefers_form_text_over_llm_text(self):
        # The parse LLM is fed two-column text that interleaves adjacent
        # interrogatories; the canonical form text read from the PDF must win.
        doc = fitz.open()
        page = doc.new_page(width=612, height=792)
        _draw_fi_row(
            page, "3.1", 54, selected=True,
            text="Are you a corporation? If so, state the corporate name.",
        )
        path = "synthetic_frog_prefer.pdf"
        try:
            doc.save(path)
            doc.close()
            parsed = ParsedDiscovery(
                discovery_type="FI",
                propounding_party="P",
                responding_party="D",
                set_number=1,
                set_word="ONE",
                case_number="1",
                requests=[
                    ParsedRequest(
                        number="3.1",
                        text="(a) the kind of coverage; (b) the insurance company",
                    )
                ],
            )
            completed = complete_selected_form_interrogatories(parsed, path, ["3.1"])
        finally:
            if not doc.is_closed:
                doc.close()
            import os
            if os.path.exists(path):
                os.remove(path)

        self.assertEqual([r.number for r in completed.requests], ["3.1"])
        self.assertIn("Are you a corporation", completed.requests[0].text)
        self.assertNotIn("coverage", completed.requests[0].text)


def _draw_fi_row(
    page,
    number: str,
    y: float,
    selected: bool,
    text: str = "Sample interrogatory text",
) -> None:
    box = fitz.Rect(36, y - 7, 54, y + 2)
    page.draw_rect(box, color=(0, 0, 0), width=0.6)
    if selected:
        page.draw_rect(fitz.Rect(43, y - 4.5, 48, y + 0.5), fill=(0, 0, 0))
    page.insert_text((60, y), f"{number} {text}", fontsize=12)


def _draw_fi_row_disc001(
    page,
    number: str,
    y: float,
    selected: bool,
    text: str = "Sample interrogatory text",
) -> None:
    """Mimic the real DISC-001 rendering: a box drawn as four thin edge-slivers,
    and a selection drawn as a filled curve glyph (checkmark), not a filled rect."""
    x0, x1 = 36.0, 49.6
    top, bot = y - 7.0, y + 4.3  # ~11pt tall box
    # four edge-slivers (thin filled rectangles)
    page.draw_rect(fitz.Rect(x0, top, x0 + 0.5, bot), fill=(0, 0, 0))      # left
    page.draw_rect(fitz.Rect(x1 - 0.5, top, x1, bot), fill=(0, 0, 0))      # right
    page.draw_rect(fitz.Rect(x0, top, x1, top + 0.5), fill=(0, 0, 0))      # top
    page.draw_rect(fitz.Rect(x0, bot - 0.5, x1, bot), fill=(0, 0, 0))      # bottom
    if selected:
        # filled glyph (curve items) ~7pt — like the form software's checkmark
        page.draw_circle((x0 + 6.8, y - 1.0), 3.4, fill=(0, 0, 0))
    page.insert_text((60, y), f"{number} {text}", fontsize=12)


class FormInterrogatoryDisc001Tests(unittest.TestCase):
    """Detection on the real Judicial Council rendering (edge-sliver boxes,
    filled-curve checkmarks) — the style that produced 'No parsed requests'."""

    def _build(self, rows):
        doc = fitz.open()
        page = doc.new_page(width=612, height=792)
        for number, y, selected in rows:
            _draw_fi_row_disc001(page, number, y, selected)
        return doc

    def test_detects_curve_checkmark_in_edge_sliver_box(self):
        doc = self._build([("1.1", 60, True), ("2.1", 110, True), ("2.3", 160, False)])
        path = "synthetic_disc001_selection.pdf"
        try:
            doc.save(path)
            doc.close()
            selected = extract_selected_form_interrogatory_numbers(path)
        finally:
            if not doc.is_closed:
                doc.close()
            import os
            if os.path.exists(path):
                os.remove(path)
        self.assertEqual(selected, ["1.1", "2.1"])

    def test_scan_lists_present_interrogatories_with_checked_flags(self):
        doc = self._build([
            ("1.1", 60, True),
            ("2.1", 110, False),
            ("2.2", 160, True),
        ])
        path = "synthetic_disc001_scan.pdf"
        try:
            doc.save(path)
            doc.close()
            scanned = scan_form_interrogatories(path)
        finally:
            if not doc.is_closed:
                doc.close()
            import os
            if os.path.exists(path):
                os.remove(path)

        by_number = {s.number: s for s in scanned}
        self.assertEqual(set(by_number), {"1.1", "2.1", "2.2"})
        self.assertTrue(by_number["1.1"].checked)
        self.assertFalse(by_number["2.1"].checked)
        self.assertTrue(by_number["2.2"].checked)
        # text is captured for the manual list
        self.assertIn("Sample interrogatory text", by_number["1.1"].text)

    def test_scan_ignores_bare_numbers_without_a_checkbox(self):
        doc = fitz.open()
        page = doc.new_page(width=612, height=792)
        _draw_fi_row_disc001(page, "1.1", 60, selected=True)
        # A statutory cross-reference with no checkbox should not be listed.
        page.insert_text((60, 200), "pursuant to Code of Civil Procedure 2033.710 herein", fontsize=12)
        path = "synthetic_disc001_noise.pdf"
        try:
            doc.save(path)
            doc.close()
            scanned = scan_form_interrogatories(path)
        finally:
            if not doc.is_closed:
                doc.close()
            import os
            if os.path.exists(path):
                os.remove(path)
        self.assertEqual([s.number for s in scanned], ["1.1"])


if __name__ == "__main__":
    unittest.main()
