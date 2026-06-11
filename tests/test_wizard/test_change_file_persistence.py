"""Regression coverage for wizard task-tab persistence through Change File.

The dialog/hotkey path (`MainWindow.change_file`) is separate from the
Master List path (`load_case_by_number`). It must still snapshot the old
case's wizard tabs, remove them, and restore the newly selected case's tabs.
"""
import os
import sys
import types

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from PySide6.QtWidgets import QApplication, QDialog


@pytest.fixture(scope="module")
def qapp():
    return QApplication.instance() or QApplication([])


def test_change_file_saves_old_case_and_restores_new_case_wizard_tabs(
    qapp, monkeypatch, tmp_path
):
    import iCharlotte as ich

    old_case = str(tmp_path / "old")
    new_case = str(tmp_path / "new")
    os.makedirs(old_case)
    os.makedirs(new_case)

    class FakeDialog:
        def __init__(self, parent):
            self.parent = parent

        def exec(self):
            return QDialog.DialogCode.Accepted

        def get_file_number(self):
            return "2222.000"

    stub = types.SimpleNamespace()
    stub.file_number = "1111.000"
    stub.case_path = old_case
    stub.agent_buttons = {}
    stub.running_agents = {}

    calls = []
    stub._save_wizard_state_for_current_case = lambda: calls.append(
        ("save_wizard", stub.case_path)
    )
    stub._remove_all_task_tabs = lambda: calls.append(("remove_tabs", stub.case_path))
    stub._restore_task_tabs_for_case = lambda: calls.append(("restore_tabs", stub.case_path))
    stub._iter_task_tabs = lambda: []
    stub.save_status_history = lambda: calls.append(("save_status", stub.case_path))
    stub.setWindowTitle = lambda title: calls.append(("title", title))
    stub.populate_tree = lambda: calls.append(("populate_tree", stub.case_path))
    stub.clear_all_status = lambda: calls.append(("clear_status", stub.case_path))
    stub.load_status_history = lambda: calls.append(("load_status", stub.case_path))
    stub._default_to_wizard_mode = lambda: calls.append(("default_wizard", stub.case_path))

    monkeypatch.setattr(ich, "FileNumberDialog", FakeDialog)
    monkeypatch.setattr(ich, "get_case_path", lambda file_number: new_case)

    stub.change_file = ich.MainWindow.change_file.__get__(stub, type(stub))
    stub.change_file()

    assert ("save_wizard", old_case) in calls
    assert ("remove_tabs", old_case) in calls
    assert ("restore_tabs", new_case) in calls
    assert calls.index(("save_wizard", old_case)) < calls.index(("remove_tabs", old_case))
    assert calls.index(("remove_tabs", old_case)) < calls.index(("restore_tabs", new_case))
