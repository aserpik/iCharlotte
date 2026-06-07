"""Tests for med record extractor page parsing and chronology document parsing."""

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


class TestNormalizeDate(unittest.TestCase):
    def test_two_digit_year_uses_1950s_1960s_as_past_dates(self):
        from icharlotte_core.med_record_extractor import _normalize_date

        self.assertEqual(_normalize_date("09/21/68"), "09/21/1968")
        self.assertEqual(_normalize_date("September 21, 68"), "09/21/1968")


def _build_chronology_docx(
    path: Path,
    *,
    include_synopsis: bool = True,
    flags_header: str = "Red Flags/Comments",
    synopsis_extra_paragraphs: tuple[str, ...] = (),
    first_page_no: str = (
        "Hall - Doc Produced HALL 000001 to 002530 7-21-2023\n\n"
        "Pg No: 599-604/2530"
    ),
    include_decoy_table: bool = False,
    after_table_paragraphs: tuple[str, ...] = (),
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
        for paragraph in synopsis_extra_paragraphs:
            doc.add_paragraph(paragraph)
        doc.add_paragraph(
            "On 09/21/2020, she had an X-ray performed by James Michael Erskine, MD."
        )
    if include_decoy_table:
        table = doc.add_table(rows=1, cols=5)
        headers = [
            "Updated Date",
            "Pageno Notes",
            "Provider",
            "Description",
            "Red Flags/Comments",
        ]
        for index, header in enumerate(headers):
            table.rows[0].cells[index].text = header
        row = table.add_row().cells
        row[0].text = "01/01/1999"
        row[1].text = "Decoy Record\n\nPg No: 1/1"
        row[2].text = "Wrong Provider"
        row[3].text = "Not a chronology table."
        row[4].text = ""
    table = doc.add_table(rows=1, cols=5)
    headers = ["DATE", "PAGE NO", "PROVIDER", "DESCRIPTION", flags_header]
    for index, header in enumerate(headers):
        table.rows[0].cells[index].text = header
    section = table.add_row().cells
    for cell in section:
        cell.text = "POST-INJURY MEDICAL RECORDS"
    row = table.add_row().cells
    row[0].text = "09/21/2020"
    row[1].text = first_page_no
    row[2].text = "Kaiser Permanente\nHenry Louis Marr, DO"
    row[3].text = "EMERGENCY DEPARTMENT NOTE\nChief Complaint: Ankle injury."
    row[4].text = ""
    row = table.add_row().cells
    row[0].text = "09/21/2020"
    row[1].text = "Hall - Doc Produced HALL 000001 to 002530 7-21-2023\n\nPg No: 598/2530"
    row[2].text = "Kaiser Permanente\nJames Michael Erskine, MD"
    row[3].text = "RADIOLOGY REPORT\nX-ray right ankle."
    row[4].text = ""
    for paragraph in after_table_paragraphs:
        doc.add_paragraph(paragraph)
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

    def test_parse_chronology_document_collects_synopsis_continuation_paragraph(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        with tempfile.TemporaryDirectory() as td:
            source = _build_chronology_docx(
                Path(td) / "chronology.docx",
                synopsis_extra_paragraphs=(
                    "She returned home with crutches and follow-up instructions.",
                ),
            )
            parsed = parse_chronology_document(str(source))

        self.assertEqual(len(parsed.synopsis_paragraphs), 3)
        self.assertEqual(
            parsed.synopsis_paragraphs[1].text,
            "She returned home with crutches and follow-up instructions.",
        )

    def test_parse_chronology_document_stops_synopsis_at_table_boundary(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        with tempfile.TemporaryDirectory() as td:
            source = _build_chronology_docx(
                Path(td) / "chronology.docx",
                after_table_paragraphs=(
                    "On 10/01/2020, this footer-style note should not be selectable.",
                ),
            )
            parsed = parse_chronology_document(str(source))

        self.assertEqual(len(parsed.synopsis_paragraphs), 2)
        self.assertNotIn(
            "10/01/2020",
            "\n".join(paragraph.text for paragraph in parsed.synopsis_paragraphs),
        )

    def test_parse_chronology_document_warns_when_page_no_is_malformed(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        with tempfile.TemporaryDirectory() as td:
            source = _build_chronology_docx(
                Path(td) / "chronology.docx",
                first_page_no=(
                    "Hall - Doc Produced HALL 000001 to 002530 7-21-2023\n\nPg No: 17-"
                ),
            )
            parsed = parse_chronology_document(str(source))

        self.assertEqual(len(parsed.rows), 2)
        self.assertFalse(parsed.rows[0].extractable)
        self.assertIn("Could not parse record/pages from PAGE NO:", parsed.rows[0].warning)

    def test_parse_chronology_document_skips_decoy_table_before_real_chronology(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        with tempfile.TemporaryDirectory() as td:
            source = _build_chronology_docx(
                Path(td) / "chronology.docx",
                include_decoy_table=True,
            )
            parsed = parse_chronology_document(str(source))

        self.assertEqual(len(parsed.rows), 2)
        self.assertEqual(parsed.rows[0].date, "09/21/2020")
        self.assertEqual(
            parsed.rows[0].record_filename,
            "Hall - Doc Produced HALL 000001 to 002530 7-21-2023",
        )
        self.assertFalse(parsed.blocking_errors)

    def test_parse_chronology_document_warns_when_page_range_is_invalid(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        with tempfile.TemporaryDirectory() as td:
            source = _build_chronology_docx(
                Path(td) / "chronology.docx",
                first_page_no=(
                    "Hall - Doc Produced HALL 000001 to 002530 7-21-2023\n\n"
                    "Pg No: 10-5/20"
                ),
            )
            parsed = parse_chronology_document(str(source))

        self.assertEqual(parsed.rows[0].page_start, 10)
        self.assertEqual(parsed.rows[0].page_end, 5)
        self.assertFalse(parsed.rows[0].extractable)
        self.assertIn("Could not parse record/pages from PAGE NO:", parsed.rows[0].warning)

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


class TestSynopsisMatching(unittest.TestCase):
    def _parsed(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        td = tempfile.TemporaryDirectory()
        self.addCleanup(td.cleanup)
        source = _build_chronology_docx(Path(td.name) / "chronology.docx")
        return parse_chronology_document(str(source))

    def test_synopsis_match_confident_by_date_and_provider(self):
        from icharlotte_core.med_record_chronology import match_synopsis_to_rows

        parsed = self._parsed()
        result = match_synopsis_to_rows(parsed.synopsis_paragraphs[0], parsed.rows)

        self.assertEqual(result.status, "confident")
        self.assertEqual(result.row_ids, (parsed.rows[0].id,))
        self.assertEqual(result.candidate_row_ids, ())

    def test_synopsis_match_ambiguous_when_provider_is_not_distinct(self):
        from icharlotte_core.med_record_chronology import SynopsisParagraph, match_synopsis_to_rows

        parsed = self._parsed()
        paragraph = SynopsisParagraph(
            id="syn-x",
            order=0,
            text="On 09/21/2020, she presented to Kaiser Permanente for ankle care.",
        )
        result = match_synopsis_to_rows(paragraph, parsed.rows)

        self.assertEqual(result.status, "ambiguous")
        self.assertEqual(set(result.candidate_row_ids), {row.id for row in parsed.rows})

    def test_synopsis_match_none_without_same_date(self):
        from icharlotte_core.med_record_chronology import SynopsisParagraph, match_synopsis_to_rows

        parsed = self._parsed()
        paragraph = SynopsisParagraph(
            id="syn-x",
            order=0,
            text="On 12/31/2021, she saw Kaiser Permanente for unrelated care.",
        )
        result = match_synopsis_to_rows(paragraph, parsed.rows)

        self.assertEqual(result.status, "none")
        self.assertEqual(result.row_ids, ())

    def test_synopsis_match_accepts_two_digit_year(self):
        from icharlotte_core.med_record_chronology import SynopsisParagraph, match_synopsis_to_rows

        parsed = self._parsed()
        paragraph = SynopsisParagraph(
            id="syn-x",
            order=0,
            text="On 09/21/20, she was evaluated by Henry Louis Marr, DO.",
        )
        result = match_synopsis_to_rows(paragraph, parsed.rows)

        self.assertEqual(result.status, "confident")
        self.assertEqual(result.row_ids, (parsed.rows[0].id,))

    def test_synopsis_match_accepts_month_name_date(self):
        from icharlotte_core.med_record_chronology import SynopsisParagraph, match_synopsis_to_rows

        parsed = self._parsed()
        paragraph = SynopsisParagraph(
            id="syn-x",
            order=0,
            text="On September 21, 2020, she was evaluated by Henry Louis Marr, DO.",
        )
        result = match_synopsis_to_rows(paragraph, parsed.rows)

        self.assertEqual(result.status, "confident")
        self.assertEqual(result.row_ids, (parsed.rows[0].id,))

    def test_synopsis_match_accepts_doctor_before_verb(self):
        from icharlotte_core.med_record_chronology import SynopsisParagraph, match_synopsis_to_rows

        parsed = self._parsed()
        paragraph = SynopsisParagraph(
            id="syn-x",
            order=0,
            text="On 09/21/2020, Dr. Marr evaluated her for ankle pain.",
        )
        result = match_synopsis_to_rows(paragraph, parsed.rows)

        self.assertEqual(result.status, "confident")
        self.assertEqual(result.row_ids, (parsed.rows[0].id,))

    def test_synopsis_match_does_not_confidently_select_generic_provider_token(self):
        from dataclasses import replace

        from icharlotte_core.med_record_chronology import SynopsisParagraph, match_synopsis_to_rows

        parsed = self._parsed()
        rows = [
            replace(parsed.rows[0], provider="Kaiser Clinic\nHenry Louis Marr, DO"),
            parsed.rows[1],
        ]
        paragraph = SynopsisParagraph(
            id="syn-x",
            order=0,
            text="On 09/21/2020, she presented to clinic for follow-up care.",
        )
        result = match_synopsis_to_rows(paragraph, rows)

        self.assertEqual(result.status, "ambiguous")
        self.assertEqual(set(result.candidate_row_ids), {row.id for row in rows})


class TestSelectionState(unittest.TestCase):
    def test_selection_state_removes_only_paragraph_owned_rows(self):
        from icharlotte_core.med_record_chronology import SelectionState

        state = SelectionState()
        state.select_row("row-manual", source="manual")
        state.select_row("row-shared", source="syn-a")
        state.select_row("row-shared", source="syn-b")
        state.select_row("row-only-a", source="syn-a")

        state.clear_source("syn-a")

        self.assertTrue(state.is_row_selected("row-manual"))
        self.assertTrue(state.is_row_selected("row-shared"))
        self.assertFalse(state.is_row_selected("row-only-a"))
        self.assertEqual(state.selected_row_ids(), ["row-manual", "row-shared"])

    def test_selection_state_manual_deselect_preserves_synopsis_owned_row(self):
        from icharlotte_core.med_record_chronology import SelectionState

        state = SelectionState()
        state.select_row("row-shared", source="syn-a")
        state.select_row("row-shared", source="manual")

        state.deselect_row("row-shared", source="manual")

        self.assertTrue(state.is_row_selected("row-shared"))
        state.clear_source("syn-a")
        self.assertFalse(state.is_row_selected("row-shared"))


if __name__ == "__main__":
    unittest.main()
