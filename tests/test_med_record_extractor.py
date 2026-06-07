"""Tests for icharlotte_core.med_record_extractor._parse_page_no."""

import tempfile
import unittest
from pathlib import Path

from docx import Document

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


def _build_chronology_docx(
    path: Path,
    *,
    include_synopsis: bool = True,
    flags_header: str = "Red Flags/Comments",
) -> Path:
    doc = Document()
    doc.add_paragraph("CHRONOLOGICAL MEDICAL SUMMARY")
    doc.add_paragraph("Client Name: Test Plaintiff")
    if include_synopsis:
        doc.add_paragraph("BRIEF SYNOPSIS OF POST-INJURY MEDICAL RECORD:")
        doc.add_paragraph(
            "On 09/21/2020, Test Plaintiff presented to Kaiser Permanente. "
            "She was evaluated by Henry Louis Marr, DO, for a right ankle injury."
        )
        doc.add_paragraph(
            "On 09/21/2020, she had an X-ray performed by James Michael Erskine, MD."
        )
    table = doc.add_table(rows=1, cols=5)
    headers = ["DATE", "PAGE NO", "PROVIDER", "DESCRIPTION", flags_header]
    for index, header in enumerate(headers):
        table.rows[0].cells[index].text = header
    section = table.add_row().cells
    for cell in section:
        cell.text = "POST-INJURY MEDICAL RECORDS"
    row = table.add_row().cells
    row[0].text = "09/21/2020"
    row[1].text = "Hall - Doc Produced HALL 000001 to 002530 7-21-2023\n\nPg No: 599-604/2530"
    row[2].text = "Kaiser Permanente\nHenry Louis Marr, DO"
    row[3].text = "EMERGENCY DEPARTMENT NOTE\nChief Complaint: Ankle injury."
    row[4].text = ""
    row = table.add_row().cells
    row[0].text = "09/21/2020"
    row[1].text = "Hall - Doc Produced HALL 000001 to 002530 7-21-2023\n\nPg No: 598/2530"
    row[2].text = "Kaiser Permanente\nJames Michael Erskine, MD"
    row[3].text = "RADIOLOGY REPORT\nX-ray right ankle."
    row[4].text = ""
    doc.save(path)
    return path


class TestChronologyDocumentParser(unittest.TestCase):
    def test_parse_chronology_document_extracts_synopsis_and_rows(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        with tempfile.TemporaryDirectory() as td:
            source = _build_chronology_docx(Path(td) / "chronology.docx")
            parsed = parse_chronology_document(str(source))

        self.assertEqual(len(parsed.synopsis_paragraphs), 2)
        self.assertIn("Henry Louis Marr", parsed.synopsis_paragraphs[0].text)
        self.assertEqual(len(parsed.rows), 2)
        self.assertEqual(parsed.rows[0].date, "09/21/2020")
        self.assertEqual(parsed.rows[0].page_start, 599)
        self.assertEqual(parsed.rows[0].page_end, 604)
        self.assertEqual(parsed.rows[0].record_filename, "Hall - Doc Produced HALL 000001 to 002530 7-21-2023")
        self.assertFalse(parsed.blocking_errors)

    def test_parse_chronology_document_warns_when_synopsis_missing(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        with tempfile.TemporaryDirectory() as td:
            source = _build_chronology_docx(Path(td) / "chronology.docx", include_synopsis=False)
            parsed = parse_chronology_document(str(source))

        self.assertEqual(parsed.synopsis_paragraphs, [])
        self.assertEqual(len(parsed.rows), 2)
        self.assertIn("No Brief Synopsis section found.", parsed.warnings)

    def test_parse_chronology_document_blocks_when_table_missing(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        with tempfile.TemporaryDirectory() as td:
            source = Path(td) / "no_table.docx"
            doc = Document()
            doc.add_paragraph("BRIEF SYNOPSIS OF POST-INJURY MEDICAL RECORD:")
            doc.add_paragraph("On 09/21/2020, Test Plaintiff saw Kaiser Permanente.")
            doc.save(source)
            parsed = parse_chronology_document(str(source))

        self.assertEqual(parsed.rows, [])
        self.assertIn("No usable 5-column chronology table found.", parsed.blocking_errors)

    def test_parse_chronology_document_blocks_when_flags_header_missing(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        with tempfile.TemporaryDirectory() as td:
            source = _build_chronology_docx(
                Path(td) / "chronology.docx",
                flags_header="Notes",
            )
            parsed = parse_chronology_document(str(source))

        self.assertEqual(parsed.rows, [])
        self.assertIn("No usable 5-column chronology table found.", parsed.blocking_errors)


if __name__ == "__main__":
    unittest.main()
