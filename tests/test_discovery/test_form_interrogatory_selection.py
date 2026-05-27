import unittest

import fitz

from icharlotte_core.discovery.form_interrogatory_selection import (
    complete_selected_form_interrogatories,
    extract_selected_form_interrogatory_numbers,
    filter_parsed_form_interrogatories,
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


if __name__ == "__main__":
    unittest.main()
