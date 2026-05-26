"""Smoke tests for InProcessTaskTab and its settings widgets."""
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
    MedExtractorSettingsPage,
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


def test_med_extractor_settings_rejects_empty(qtbot):
    w = MedExtractorSettingsPage()
    qtbot.addWidget(w)
    # No text → run_requested must NOT fire. (A QMessageBox pops; pytest-qt
    # auto-dismisses if we patch QMessageBox, but easier: just wait briefly and
    # confirm no signal.)
    from unittest.mock import patch
    with patch("icharlotte_core.ui.wizard.in_process_task_tab.QMessageBox.warning"):
        signal_received = []
        w.run_requested.connect(lambda d: signal_received.append(d))
        w.extract_btn.click()
    assert signal_received == []


def test_med_extractor_settings_emits_text(qtbot):
    w = MedExtractorSettingsPage()
    qtbot.addWidget(w)
    w.text_input.setPlainText("On Jan 1 2025, Plaintiff saw Dr. Smith at City Hospital.")
    with qtbot.waitSignal(w.run_requested, timeout=500) as blocker:
        w.extract_btn.click()
    assert blocker.args[0]["text"].startswith("On Jan 1 2025")


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
