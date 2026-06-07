"""Smoke tests for InProcessTaskTab and its settings widgets."""
import json
import os
import pytest
from unittest.mock import patch

pytest.importorskip("pytestqt")
from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import QWidget

from icharlotte_core.ui.wizard.registry import get_task
from icharlotte_core.ui.wizard.in_process_task_tab import (
    InProcessTaskTab,
    MedExtractorOutputPage,
    SubpoenaSettingsPage,
    PAGE_OUTPUT,
    PAGE_SETTINGS,
    PAGE_STATUS,
)


class _SettingsWidget(QWidget):
    run_requested = Signal(dict)


class _FakeWorker(QThread):
    """Mimics the worker contract used by InProcessTaskTab."""

    progress = Signal(str)
    warning = Signal(str)
    finished_result = Signal(bool, str)

    def __init__(self, *, ok: bool, payload: str, parent=None):
        super().__init__(parent)
        self._ok = ok
        self._payload = payload

    def run(self):
        self.progress.emit("working")
        self.finished_result.emit(self._ok, self._payload)


def _make_tab(qtbot, *, settings_widget, output_widget, ok, payload, auto_run=False):
    spec = get_task("subpoena_tracker")  # arbitrary; only used for title/task_id

    def factory(cp, fn, settings, p):
        return _FakeWorker(ok=ok, payload=payload, parent=p)

    tab = InProcessTaskTab(
        spec=spec,
        case_path="/tmp/case",
        file_number="0000.000",
        settings_widget=settings_widget,
        output_widget=output_widget,
        worker_factory=factory,
        auto_run=auto_run,
    )
    qtbot.addWidget(tab)
    return tab


def test_subpoena_settings_emits_run_requested(qtbot):
    w = SubpoenaSettingsPage()
    qtbot.addWidget(w)
    with qtbot.waitSignal(w.run_requested, timeout=500) as blocker:
        w.run_btn.click()
    assert blocker.args[0] == {}


def test_success_path_transitions_to_output(qtbot, tmp_path):
    output_file = tmp_path / "report.docx"
    output_file.write_bytes(b"x")
    from icharlotte_core.ui.wizard.pages.output_page import OutputPage
    out = OutputPage()
    tab = _make_tab(
        qtbot,
        settings_widget=SubpoenaSettingsPage(),
        output_widget=out,
        ok=True,
        payload=str(output_file),
    )
    assert tab.currentIndex() == PAGE_SETTINGS

    with qtbot.waitSignal(tab.task_completed, timeout=2000):
        tab.settings_page.run_btn.click()

    assert tab.currentIndex() == PAGE_OUTPUT
    assert out.output_path == str(output_file)


def test_success_path_applies_save_as_defaults_from_settings(qtbot, tmp_path):
    output_file = tmp_path / "preview.docx"
    output_file.write_bytes(b"x")
    from icharlotte_core.ui.wizard.pages.output_page import OutputPage
    out = OutputPage()
    settings = _SettingsWidget()
    default_dir = str(tmp_path / "DISCOVERY" / "RESPONSES")
    suggested_filename = "Def Jones's Resp to SI(1).docx"
    tab = _make_tab(
        qtbot,
        settings_widget=settings,
        output_widget=out,
        ok=True,
        payload=str(output_file),
    )

    with qtbot.waitSignal(tab.task_completed, timeout=2000):
        settings.run_requested.emit(
            {
                "save_default_dir": default_dir,
                "suggested_filename": suggested_filename,
            }
        )

    chosen = tmp_path / "DISCOVERY" / "RESPONSES" / "final.docx"
    with patch(
        "icharlotte_core.ui.wizard.pages.output_page.QFileDialog.getSaveFileName",
        return_value=(str(chosen), "Word Documents (*.docx)"),
    ) as mock_dialog:
        with patch("icharlotte_core.ui.wizard.pages.output_page.shutil.copyfile"):
            with patch(
                "icharlotte_core.ui.wizard.pages.output_page.QMessageBox.information"
            ):
                out.save_btn.click()

    assert mock_dialog.call_args.args[2] == str(
        tmp_path / "DISCOVERY" / "RESPONSES" / suggested_filename
    )


def test_success_path_records_settings_files_in_completed_entry(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.output_page import OutputPage

    output_file = tmp_path / "preview.docx"
    output_file.write_bytes(b"x")
    chronology_path = tmp_path / "RECORDS" / "summary.docx"
    chronology_path.parent.mkdir()
    chronology_path.write_bytes(b"x")
    out = OutputPage()
    settings = _SettingsWidget()
    tab = _make_tab(
        qtbot,
        settings_widget=settings,
        output_widget=out,
        ok=True,
        payload=str(output_file),
    )

    with qtbot.waitSignal(tab.task_completed, timeout=2000) as blocker:
        settings.run_requested.emit(
            {
                "chronology_path": str(chronology_path),
                "files": [str(tmp_path / "other.pdf")],
            }
        )

    entry = blocker.args[0]
    assert entry["files"] == [str(tmp_path / "other.pdf"), str(chronology_path)]


def test_success_path_serializes_selected_rows_in_completed_settings(qtbot, tmp_path):
    from icharlotte_core.med_record_chronology import SelectableChronologyRow
    from icharlotte_core.ui.wizard.pages.output_page import OutputPage

    output_file = tmp_path / "preview.docx"
    output_file.write_bytes(b"x")
    row = SelectableChronologyRow(
        id="row-1",
        order=0,
        date="09/21/2020",
        page_no="source\n\nPg No: 2/3",
        provider="Kaiser Permanente",
        description="Emergency department note",
        flags="",
        record_filename="source",
        page_start=2,
        page_end=2,
    )
    out = OutputPage()
    settings = _SettingsWidget()
    tab = _make_tab(
        qtbot,
        settings_widget=settings,
        output_widget=out,
        ok=True,
        payload=str(output_file),
    )

    with qtbot.waitSignal(tab.task_completed, timeout=2000) as blocker:
        settings.run_requested.emit(
            {
                "chronology_path": str(tmp_path / "summary.docx"),
                "selected_rows": [row],
            }
        )

    entry = blocker.args[0]
    json.dumps(entry)
    assert entry["settings"]["selected_rows"] == [
        {
            "id": "row-1",
            "order": 0,
            "date": "09/21/2020",
            "page_no": "source\n\nPg No: 2/3",
            "provider": "Kaiser Permanente",
            "description": "Emergency department note",
            "flags": "",
            "record_filename": "source",
            "page_start": 2,
            "page_end": 2,
            "warning": "",
        }
    ]


def test_success_path_merges_settings_page_state_in_completed_entry(qtbot, tmp_path):
    from icharlotte_core.med_record_chronology import SelectableChronologyRow
    from icharlotte_core.ui.wizard.pages.output_page import OutputPage

    class StatefulSettings(_SettingsWidget):
        def to_dict(self):
            return {
                "selected_row_sources": {"row-1": ["manual", "syn-1"]},
                "selected_paragraph_ids": ["syn-1"],
            }

    output_file = tmp_path / "preview.docx"
    output_file.write_bytes(b"x")
    row = SelectableChronologyRow(
        id="row-1",
        order=0,
        date="09/21/2020",
        page_no="source\n\nPg No: 2/3",
        provider="Kaiser Permanente",
        description="Emergency department note",
        flags="",
        record_filename="source",
        page_start=2,
        page_end=2,
    )
    out = OutputPage()
    settings = StatefulSettings()
    tab = _make_tab(
        qtbot,
        settings_widget=settings,
        output_widget=out,
        ok=True,
        payload=str(output_file),
    )

    with qtbot.waitSignal(tab.task_completed, timeout=2000) as blocker:
        settings.run_requested.emit(
            {
                "chronology_path": str(tmp_path / "summary.docx"),
                "selected_rows": [row],
            }
        )

    entry = blocker.args[0]
    assert entry["settings"]["selected_rows"][0]["id"] == "row-1"
    assert entry["settings"]["selected_row_sources"] == {
        "row-1": ["manual", "syn-1"],
    }
    assert entry["settings"]["selected_paragraph_ids"] == ["syn-1"]


def test_failure_path_stays_on_status(qtbot):
    from icharlotte_core.ui.wizard.pages.output_page import OutputPage
    out = OutputPage()
    tab = _make_tab(
        qtbot,
        settings_widget=SubpoenaSettingsPage(),
        output_widget=out,
        ok=False,
        payload="something went wrong",
    )
    tab.settings_page.run_btn.click()
    qtbot.waitUntil(lambda: "FAILED" in tab.status_page.status_label.text(), timeout=2000)
    assert tab.currentIndex() == PAGE_STATUS
    assert "something went wrong" in tab.status_page.status_label.text()


def test_med_extractor_output_show_result(qtbot, tmp_path):
    page = MedExtractorOutputPage(str(tmp_path))
    qtbot.addWidget(page)
    page.show_result("3 record(s) extracted")
    assert "3 record(s) extracted" in page.summary_view.toPlainText()
