"""Restore/reopen persistence for Case Intake & Docket wizard tabs."""
import os
import sys

import pytest

pytest.importorskip("pytestqt")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from PySide6.QtWidgets import QTabWidget, QWidget


class _Stub(QWidget):
    """Minimal MainWindow stand-in for wizard tab restore/reopen methods."""

    def __init__(self, case_path, file_number="0000.000"):
        super().__init__()
        self.tabs = QTabWidget()
        self.case_path = case_path
        self.file_number = file_number

    def _on_task_completed(self, entry):
        pass

    def _hide_fixed_close_buttons(self):
        pass


def _bind(stub, *method_names):
    import iCharlotte as ich

    for method_name in method_names:
        method = getattr(ich.MainWindow, method_name)
        setattr(stub, method_name, method.__get__(stub, type(stub)))


def _write_outputs(tmp_path):
    out_dir = tmp_path / "NOTES" / "AI OUTPUT"
    out_dir.mkdir(parents=True)
    docket_pdf = out_dir / "docket.pdf"
    variables_docx = out_dir / "variables.docx"
    docket_pdf.write_bytes(b"%PDF-1.4\n")
    variables_docx.write_bytes(b"variables")
    return docket_pdf, variables_docx


def _metadata():
    return {
        "case_number": "30-2026-00000001-CU-PO-CJC",
        "venue_county": "Orange",
    }


def _summary(docket_pdf, variables_docx):
    return {
        "success": True,
        "state": "success",
        "status": "Docket finished.",
        "warning": "",
        "docket_pdf": str(docket_pdf),
        "variables_docx": str(variables_docx),
        "trial_date": "2026-09-01",
        "other_hearings": "",
        "procedural_history": "Complaint filed.",
        "recent_lines": ["docket complete"],
    }


def test_restore_reloads_case_intake_docket_output_state(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.case_intake_docket_page import TASK_PAGE_OUTPUT
    from icharlotte_core.ui.wizard.persistence import WizardStatePersistence

    case_path = str(tmp_path)
    docket_pdf, variables_docx = _write_outputs(tmp_path)
    summary = _summary(
        os.path.relpath(str(docket_pdf), case_path),
        os.path.relpath(str(variables_docx), case_path),
    )
    metadata = _metadata()

    p = WizardStatePersistence(case_path)
    p.set_open_tabs([{
        "task_id": "case_intake_docket",
        "instance_suffix": "",
        "files": [],
        "settings": {},
        "metadata": metadata,
        "summary": summary,
        "page": "output",
        "output_path": os.path.relpath(str(docket_pdf), case_path),
    }])
    p.save()

    stub = _Stub(case_path)
    qtbot.addWidget(stub)
    _bind(stub, "_restore_task_tabs_for_case")

    stub._restore_task_tabs_for_case()

    assert stub.tabs.count() == 1
    tab = stub.tabs.widget(0)
    assert type(tab).__name__ == "CaseIntakeDocketTaskTab"
    assert tab.currentIndex() == TASK_PAGE_OUTPUT
    assert os.path.basename(tab.output_page.output_path) == "docket.pdf"
    reviewed = tab.review_page.to_dict()
    assert reviewed["case_number"] == metadata["case_number"]
    assert reviewed["venue_county"] == metadata["venue_county"]


def test_reopen_recent_reloads_case_intake_docket_output_state(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.case_intake_docket_page import TASK_PAGE_OUTPUT

    case_path = str(tmp_path)
    docket_pdf, variables_docx = _write_outputs(tmp_path)
    summary = _summary(
        os.path.relpath(str(docket_pdf), case_path),
        os.path.relpath(str(variables_docx), case_path),
    )
    metadata = _metadata()

    stub = _Stub(case_path)
    qtbot.addWidget(stub)
    _bind(stub, "_on_reopen_recent_task")

    stub._on_reopen_recent_task({
        "task_id": "case_intake_docket",
        "files": [],
        "metadata": metadata,
        "summary": summary,
        "output_path": os.path.relpath(str(docket_pdf), case_path),
    })

    assert stub.tabs.count() == 1
    tab = stub.tabs.widget(0)
    assert type(tab).__name__ == "CaseIntakeDocketTaskTab"
    assert tab.currentIndex() == TASK_PAGE_OUTPUT
    assert os.path.basename(tab.output_page.output_path) == "docket.pdf"
    reviewed = tab.review_page.to_dict()
    assert reviewed["case_number"] == metadata["case_number"]
    assert reviewed["venue_county"] == metadata["venue_county"]


def test_snapshot_includes_case_intake_docket_summary_and_metadata(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.case_intake_docket_page import (
        TASK_PAGE_OUTPUT,
        CaseIntakeDocketTaskTab,
    )
    from icharlotte_core.ui.wizard.registry import get_task

    case_path = str(tmp_path)
    docket_pdf, variables_docx = _write_outputs(tmp_path)
    metadata = _metadata()
    summary = _summary(docket_pdf, variables_docx)

    stub = _Stub(case_path)
    qtbot.addWidget(stub)
    tab = CaseIntakeDocketTaskTab(
        get_task("case_intake_docket"),
        case_path,
        stub.file_number,
        parent=stub,
    )
    tab.setProperty("wizard_task_id", "case_intake_docket")
    tab.setProperty("wizard_instance_suffix", "")
    tab.load_output_summary(summary, metadata=metadata)
    stub.tabs.addTab(tab, "Case Intake & Docket")
    _bind(stub, "_iter_task_tabs", "_relpath_under", "_snapshot_open_task_tabs")

    snapshots = stub._snapshot_open_task_tabs(cancel_running=False)

    assert len(snapshots) == 1
    snap = snapshots[0]
    assert tab.currentIndex() == TASK_PAGE_OUTPUT
    assert snap["task_id"] == "case_intake_docket"
    assert snap["page"] == "output"
    assert os.path.basename(snap["output_path"]) == "docket.pdf"
    assert snap["summary"]["docket_pdf"] == str(docket_pdf)
    assert snap["summary"]["variables_docx"] == str(variables_docx)
    assert snap["metadata"]["case_number"] == metadata["case_number"]
    assert snap["metadata"]["venue_county"] == metadata["venue_county"]
    assert snap["settings"]["case_number"] == metadata["case_number"]
    assert snap["settings"]["venue_county"] == metadata["venue_county"]
