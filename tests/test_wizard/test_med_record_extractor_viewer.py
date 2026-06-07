"""Tests for the Med Record Extractor chronology viewer."""
from __future__ import annotations

import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest
from docx import Document

pytest.importorskip("pytestqt")

from tests.test_med_record_extractor import _build_chronology_docx


def test_selection_page_loads_chronology_document(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    with tempfile.TemporaryDirectory() as td:
        source = _build_chronology_docx(Path(td) / "chronology.docx")
        page = MedChronologySelectionPage(
            case_path=td,
            file_number="5800.013",
            chronology_path=str(source),
        )
        qtbot.addWidget(page)

        assert page.selected_count_label.text() == "0 rows selected"
        assert page.synopsis_panel.count() == 2
        assert page.table_panel.count() == 2


def test_direct_row_selection_emits_selected_rows(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    with tempfile.TemporaryDirectory() as td:
        source = _build_chronology_docx(Path(td) / "chronology.docx")
        page = MedChronologySelectionPage(
            case_path=td,
            file_number="5800.013",
            chronology_path=str(source),
        )
        qtbot.addWidget(page)
        page.set_row_checked(page.document.rows[0].id, True)

        with qtbot.waitSignal(page.run_requested, timeout=500) as blocker:
            page.extract_btn.click()

        settings = blocker.args[0]
        assert settings["chronology_path"] == str(source)
        assert [row.id for row in settings["selected_rows"]] == [page.document.rows[0].id]


def test_synopsis_selection_auto_selects_confident_row(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    with tempfile.TemporaryDirectory() as td:
        source = _build_chronology_docx(Path(td) / "chronology.docx")
        page = MedChronologySelectionPage(
            case_path=td,
            file_number="5800.013",
            chronology_path=str(source),
        )
        qtbot.addWidget(page)
        page.set_paragraph_checked(page.document.synopsis_paragraphs[0].id, True)

        assert page.is_row_checked(page.document.rows[0].id)
        assert page.selected_count_label.text() == "1 row selected"


def test_deselecting_synopsis_owned_row_keeps_visual_selection_in_sync(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    with tempfile.TemporaryDirectory() as td:
        source = _build_chronology_docx(Path(td) / "chronology.docx")
        page = MedChronologySelectionPage(
            case_path=td,
            file_number="5800.013",
            chronology_path=str(source),
        )
        qtbot.addWidget(page)

        row_id = page.document.rows[0].id
        page.set_paragraph_checked(page.document.synopsis_paragraphs[0].id, True)
        page.set_row_checked(row_id, False)

        assert page.is_row_checked(row_id)
        assert page.selected_count_label.text() == "1 row selected"


def test_ambiguous_synopsis_selection_does_not_auto_select_rows(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    with tempfile.TemporaryDirectory() as td:
        source = _build_chronology_docx(
            Path(td) / "chronology.docx",
            synopsis_extra_paragraphs=(
                "On 09/21/2020, she presented to Kaiser Permanente for ankle care.",
            ),
        )
        page = MedChronologySelectionPage(
            case_path=td,
            file_number="5800.013",
            chronology_path=str(source),
        )
        qtbot.addWidget(page)

        ambiguous = page.document.synopsis_paragraphs[1]
        page.set_paragraph_checked(ambiguous.id, True)

        assert all(not page.is_row_checked(row.id) for row in page.document.rows)
        assert page.selected_count_label.text() == "0 rows selected"


def test_open_original_uses_os_startfile(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    with tempfile.TemporaryDirectory() as td:
        source = _build_chronology_docx(Path(td) / "chronology.docx")
        page = MedChronologySelectionPage(
            case_path=td,
            file_number="5800.013",
            chronology_path=str(source),
        )
        qtbot.addWidget(page)

        with patch(
            "icharlotte_core.ui.wizard.pages.med_record_extractor_page.os.startfile"
        ) as startfile:
            page.open_original_btn.click()

        startfile.assert_called_once_with(str(source))


def test_blocking_parse_errors_disable_extract(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    with tempfile.TemporaryDirectory() as td:
        source = Path(td) / "no_table.docx"
        doc = Document()
        doc.add_paragraph("BRIEF SYNOPSIS OF POST-INJURY MEDICAL RECORD:")
        doc.add_paragraph("On 09/21/2020, Test Plaintiff saw Kaiser Permanente.")
        doc.save(source)

        page = MedChronologySelectionPage(
            case_path=td,
            file_number="5800.013",
            chronology_path=str(source),
        )
        qtbot.addWidget(page)

        assert page.extract_btn.isEnabled() is False
        assert page.selected_count_label.text() == "0 rows selected"


def test_non_extractable_rows_are_not_emitted(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    with tempfile.TemporaryDirectory() as td:
        source = _build_chronology_docx(
            Path(td) / "chronology.docx",
            first_page_no=(
                "Hall - Doc Produced HALL 000001 to 002530 7-21-2023\n\nPg No: 17-"
            ),
        )
        page = MedChronologySelectionPage(
            case_path=td,
            file_number="5800.013",
            chronology_path=str(source),
        )
        qtbot.addWidget(page)

        page.set_row_checked(page.document.rows[0].id, True)
        page.set_row_checked(page.document.rows[1].id, True)

        with qtbot.waitSignal(page.run_requested, timeout=500) as blocker:
            page.extract_btn.click()

        assert [row.id for row in blocker.args[0]["selected_rows"]] == [
            page.document.rows[1].id
        ]


def test_non_extractable_rows_cannot_remain_checked(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    with tempfile.TemporaryDirectory() as td:
        source = _build_chronology_docx(
            Path(td) / "chronology.docx",
            first_page_no=(
                "Hall - Doc Produced HALL 000001 to 002530 7-21-2023\n\nPg No: 17-"
            ),
        )
        page = MedChronologySelectionPage(
            case_path=td,
            file_number="5800.013",
            chronology_path=str(source),
        )
        qtbot.addWidget(page)

        page.set_row_checked(page.document.rows[0].id, True)

        assert page.is_row_checked(page.document.rows[0].id) is False
        assert page.selected_count_label.text() == "0 rows selected"
        assert page.extract_btn.isEnabled() is False


def test_synopsis_owned_non_extractable_row_is_not_selected(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    with tempfile.TemporaryDirectory() as td:
        source = _build_chronology_docx(
            Path(td) / "chronology.docx",
            first_page_no=(
                "Hall - Doc Produced HALL 000001 to 002530 7-21-2023\n\nPg No: 17-"
            ),
        )
        page = MedChronologySelectionPage(
            case_path=td,
            file_number="5800.013",
            chronology_path=str(source),
        )
        qtbot.addWidget(page)

        page.set_paragraph_checked(page.document.synopsis_paragraphs[0].id, True)

        assert page.is_row_checked(page.document.rows[0].id) is False
        assert page.selected_count_label.text() == "0 rows selected"
        assert page.extract_btn.isEnabled() is False
        assert "not extractable" in page.match_status_label.text()


def test_failed_synopsis_match_updates_visible_status(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    with tempfile.TemporaryDirectory() as td:
        source = _build_chronology_docx(
            Path(td) / "chronology.docx",
            synopsis_extra_paragraphs=(
                "She returned home with crutches and follow-up instructions.",
            ),
        )
        page = MedChronologySelectionPage(
            case_path=td,
            file_number="5800.013",
            chronology_path=str(source),
        )
        qtbot.addWidget(page)

        no_date = page.document.synopsis_paragraphs[1]
        page.set_paragraph_checked(no_date.id, True)

        assert "No date found" in page.match_status_label.text()
        assert page.match_status_label.isHidden() is False


def test_build_med_extractor_tab_uses_docx_picker_and_summary_folder(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.in_process_task_tab import build_med_extractor_tab
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )
    from icharlotte_core.ui.wizard.registry import get_task

    summary_dir = tmp_path / "RECORDS" / "Medical Summary - DO NOT PRODUCE"
    summary_dir.mkdir(parents=True)
    source = _build_chronology_docx(summary_dir / "chronology.docx")

    with patch(
        "icharlotte_core.ui.wizard.in_process_task_tab.QFileDialog.getOpenFileName",
        return_value=(str(source), "Word Documents (*.docx)"),
    ) as picker:
        tab = build_med_extractor_tab(
            get_task("med_record_extractor"),
            case_path=str(tmp_path),
            file_number="5800.013",
            parent=None,
        )

    qtbot.addWidget(tab)
    assert isinstance(tab.settings_page, MedChronologySelectionPage)
    assert picker.call_args.args[2] == str(summary_dir)


def test_build_med_extractor_tab_cancel_returns_none(tmp_path):
    from icharlotte_core.ui.wizard.in_process_task_tab import build_med_extractor_tab
    from icharlotte_core.ui.wizard.registry import get_task

    with patch(
        "icharlotte_core.ui.wizard.in_process_task_tab.QFileDialog.getOpenFileName",
        return_value=("", ""),
    ):
        tab = build_med_extractor_tab(
            get_task("med_record_extractor"),
            case_path=str(tmp_path),
            file_number="5800.013",
            parent=None,
        )

    assert tab is None


def test_match_status_reflects_checked_failed_synopsis_paragraphs(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    with tempfile.TemporaryDirectory() as td:
        source = _build_chronology_docx(
            Path(td) / "chronology.docx",
            synopsis_extra_paragraphs=(
                "She returned home with crutches and follow-up instructions.",
            ),
        )
        page = MedChronologySelectionPage(
            case_path=td,
            file_number="5800.013",
            chronology_path=str(source),
        )
        qtbot.addWidget(page)

        failed = page.document.synopsis_paragraphs[1]
        confident = page.document.synopsis_paragraphs[0]
        page.set_paragraph_checked(failed.id, True)
        page.set_paragraph_checked(confident.id, True)

        assert "No date found" in page.match_status_label.text()

        page.set_paragraph_checked(failed.id, False)

        assert page.match_status_label.text() == ""
        assert page.match_status_label.isHidden()
