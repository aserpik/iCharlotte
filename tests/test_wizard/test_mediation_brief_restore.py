"""Regression: a Mediation Brief tab left open on its Output page must restore
the generated brief after an app restart — not come back blank on Settings.

Exercises the real MainWindow._restore_task_tabs_for_case /
_on_reopen_recent_task routing via a lightweight QWidget stub (the established
pattern from test_chat_task_tab_spawn.py), so it catches the in-process-builder
branch that previously skipped output restore for mediation_brief.
"""
import os
import sys

import pytest

pytest.importorskip("pytestqt")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QTabWidget, QWidget


class _Stub(QWidget):
    """MainWindow stand-in with just enough surface for the restore/reopen
    methods. Must be a real QWidget because those methods construct task tabs
    with ``parent=self``."""

    def __init__(self, case_path, file_number="0000.000"):
        super().__init__()
        self.tabs = QTabWidget()
        self.case_path = case_path
        self.file_number = file_number

    def _on_task_completed(self, entry):  # connected to tab.task_completed
        pass

    def _hide_fixed_close_buttons(self):
        pass


def _bind(stub, method_name):
    import iCharlotte as ich

    method = getattr(ich.MainWindow, method_name)
    setattr(stub, method_name, method.__get__(stub, type(stub)))


def _write_brief(tmp_path):
    """Create a real .docx brief under NOTES/AI OUTPUT/MEDIATION; return its path."""
    from docx import Document

    out_dir = tmp_path / "NOTES" / "AI OUTPUT" / "MEDIATION"
    out_dir.mkdir(parents=True)
    brief = out_dir / "Defendant's Confidential Mediation Brief.docx"
    doc = Document()
    doc.add_paragraph("restored brief body")
    doc.save(str(brief))
    return brief


def test_restore_reloads_mediation_brief_output(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.persistence import WizardStatePersistence
    from icharlotte_core.ui.wizard.pages.mediation_brief_page import TASK_PAGE_OUTPUT

    case_path = str(tmp_path)
    brief = _write_brief(tmp_path)

    p = WizardStatePersistence(case_path)
    p.set_open_tabs([{
        "task_id": "mediation_brief",
        "instance_suffix": "",
        "files": [],
        "settings": {"files": [], "caption_path": ""},
        "page": "output",
        "output_path": os.path.relpath(str(brief), case_path),
    }])
    p.save()

    stub = _Stub(case_path)
    qtbot.addWidget(stub)
    _bind(stub, "_restore_task_tabs_for_case")
    stub._restore_task_tabs_for_case()

    assert stub.tabs.count() == 1
    tab = stub.tabs.widget(0)
    assert type(tab).__name__ == "MediationBriefTaskTab"
    # Must come back on the OUTPUT page with the brief loaded, not blank Settings.
    assert tab.currentIndex() == TASK_PAGE_OUTPUT
    assert tab.output_page.output_path
    assert os.path.basename(tab.output_page.output_path) == os.path.basename(str(brief))


def test_reopen_recent_reloads_mediation_brief_output(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.mediation_brief_page import TASK_PAGE_OUTPUT

    case_path = str(tmp_path)
    brief = _write_brief(tmp_path)

    stub = _Stub(case_path)
    qtbot.addWidget(stub)
    _bind(stub, "_on_reopen_recent_task")
    stub._on_reopen_recent_task({
        "task_id": "mediation_brief",
        "instance_suffix": "",
        "files": [],
        "settings": {"files": [], "caption_path": ""},
        "output_path": os.path.relpath(str(brief), case_path),
    })

    assert stub.tabs.count() == 1
    tab = stub.tabs.widget(0)
    assert type(tab).__name__ == "MediationBriefTaskTab"
    assert tab.currentIndex() == TASK_PAGE_OUTPUT
    assert os.path.basename(tab.output_page.output_path) == os.path.basename(str(brief))
