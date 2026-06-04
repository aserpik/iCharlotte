"""Regression: a Generate Motion tab left open on its Output page must restore
the generated motion after an app restart / recent-task reopen — not come back
on the blank intake Settings page.

Mirrors test_mediation_brief_restore.py: exercises the real
MainWindow._restore_task_tabs_for_case / _on_reopen_recent_task routing via a
lightweight QWidget stub, so it catches the in-process-builder branch that
previously skipped output restore for generate_motion.
"""
import os
import sys

import pytest

pytest.importorskip("pytestqt")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QTabWidget, QWidget


class _Stub(QWidget):
    """MainWindow stand-in with just enough surface for the restore/reopen
    methods (they build task tabs with ``parent=self``)."""

    def __init__(self, case_path, file_number="0000.000"):
        super().__init__()
        self.tabs = QTabWidget()
        self.case_path = case_path
        self.file_number = file_number

    def _on_task_completed(self, entry):
        pass

    def _hide_fixed_close_buttons(self):
        pass


def _bind(stub, method_name):
    import iCharlotte as ich

    method = getattr(ich.MainWindow, method_name)
    setattr(stub, method_name, method.__get__(stub, type(stub)))


def _write_motion(tmp_path):
    """Create a real .docx motion preview; return its path."""
    from docx import Document

    out_dir = tmp_path / "NOTES" / "AI OUTPUT" / ".icharlotte" / "wizard_previews" / "generate_motion"
    out_dir.mkdir(parents=True)
    preview = out_dir / "Motion Preview.docx"
    doc = Document()
    doc.add_paragraph("restored motion body")
    doc.save(str(preview))
    return preview


def _settings():
    return {
        "motion_type_id": "generic",
        "motion_type_name": "Motion in Limine to Exclude Witnesses",
        "target_files": [],
        "metadata": {
            "motion_type": "Motion in Limine to Exclude Witnesses",
            "relief_requested": "Exclude the witnesses",
            "principal_arguments": ["Undisclosed witness"],
        },
        "outline": [],
    }


def test_restore_reloads_generate_motion_output(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.persistence import WizardStatePersistence
    from icharlotte_core.ui.wizard.pages.generate_motion_page import TASK_PAGE_OUTPUT

    case_path = str(tmp_path)
    preview = _write_motion(tmp_path)

    p = WizardStatePersistence(case_path)
    p.set_open_tabs([{
        "task_id": "generate_motion",
        "instance_suffix": "",
        "files": [],
        "settings": _settings(),
        "page": "output",
        "output_path": os.path.relpath(str(preview), case_path),
    }])
    p.save()

    stub = _Stub(case_path)
    qtbot.addWidget(stub)
    _bind(stub, "_restore_task_tabs_for_case")
    stub._restore_task_tabs_for_case()

    assert stub.tabs.count() == 1
    tab = stub.tabs.widget(0)
    assert type(tab).__name__ == "GenerateMotionTaskTab"
    # Must restore on the OUTPUT page with the motion loaded, not blank intake.
    assert tab.currentIndex() == TASK_PAGE_OUTPUT
    assert tab.output_page.output_path
    assert os.path.basename(tab.output_page.output_path) == os.path.basename(str(preview))


def test_reopen_recent_reloads_generate_motion_output(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.generate_motion_page import TASK_PAGE_OUTPUT

    case_path = str(tmp_path)
    preview = _write_motion(tmp_path)

    stub = _Stub(case_path)
    qtbot.addWidget(stub)
    _bind(stub, "_on_reopen_recent_task")
    stub._on_reopen_recent_task({
        "task_id": "generate_motion",
        "instance_suffix": "",
        "files": [],
        "settings": _settings(),
        "output_path": os.path.relpath(str(preview), case_path),
    })

    assert stub.tabs.count() == 1
    tab = stub.tabs.widget(0)
    assert type(tab).__name__ == "GenerateMotionTaskTab"
    assert tab.currentIndex() == TASK_PAGE_OUTPUT
    assert os.path.basename(tab.output_page.output_path) == os.path.basename(str(preview))
