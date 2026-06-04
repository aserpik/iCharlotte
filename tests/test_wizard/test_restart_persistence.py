"""Restart must persist wizard session state before quitting.

Root cause this pins: the in-app red **Restart** button calls
``restart_app() -> QApplication.quit()``. Unlike a window-manager close,
``QApplication.quit()`` does NOT dispatch ``closeEvent``, so the open wizard
task tabs (and their contents) were never saved when restarting via the
button — the most common "restart iCharlotte" path. State only survived if the
user happened to switch cases or close via the OS window control after opening
the tabs, which is why persistence was intermittent.

The fix routes both shutdown paths (closeEvent and restart_app) through a
single ``_persist_session_state`` method, and calls it in restart_app before
quitting.

Uses the lightweight stub pattern from ``test_main_window_restart.py``: bind
the real method onto a barebones object so we exercise the logic without
constructing the whole MainWindow UI tree.
"""
import os
import sys
import types

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from PySide6.QtWidgets import QApplication, QTabWidget, QWidget


@pytest.fixture(scope="module")
def qapp():
    return QApplication.instance() or QApplication([])


class _FakeWorker:
    def __init__(self):
        self.cancelled = False

    def cancel(self):
        self.cancelled = True


class _FakeTaskTab(QWidget):
    """Minimal stand-in for a wizard TaskTab for snapshot tests.

    page_idx: 1 == PAGE_STATUS (running), 2 == PAGE_OUTPUT, else settings.
    """

    def __init__(self, task_id="summarize_documents", page_idx=1, worker=None):
        super().__init__()
        self.setProperty("wizard_task_id", task_id)
        self.setProperty("wizard_instance_suffix", "")
        self.spec = types.SimpleNamespace(task_id=task_id)
        self._page_idx = page_idx
        self._worker = worker
        self.files = []
        self.settings_page = types.SimpleNamespace(to_dict=lambda: {})
        self.output_page = types.SimpleNamespace(output_path=None)

    def currentIndex(self):
        return self._page_idx


def _snapshot_stub(tmp_path, fake_tab):
    import iCharlotte as ich
    stub = types.SimpleNamespace()
    stub.case_path = str(tmp_path)
    stub.tabs = QTabWidget()
    stub.tabs.addTab(fake_tab, "Summarize")
    for m in ("_iter_task_tabs", "_relpath_under", "_snapshot_open_task_tabs"):
        setattr(stub, m, getattr(ich.MainWindow, m).__get__(stub, type(stub)))
    return stub


def _bind(method_name, stub):
    import iCharlotte as ich
    setattr(stub, method_name,
            getattr(ich.MainWindow, method_name).__get__(stub, type(stub)))
    return ich


def test_restart_app_persists_session_state_before_quit(qapp, monkeypatch):
    stub = types.SimpleNamespace()
    stub.agent_runners = []
    stub.file_number = "1234"
    stub.case_path = r"C:\cases\1234"
    stub.tabs = QTabWidget()
    ich = _bind("restart_app", stub)

    order = []
    stub._persist_session_state = lambda: order.append("persist")
    monkeypatch.setattr(ich.subprocess, "Popen", lambda *a, **k: order.append("popen"))
    # Replace the module-global QApplication so .quit() is observable and inert.
    monkeypatch.setattr(ich, "QApplication",
                        types.SimpleNamespace(quit=lambda: order.append("quit")))

    stub.restart_app()

    assert "persist" in order, "restart_app must persist session state before quitting"
    assert "quit" in order, "restart_app must still quit"
    assert order.index("persist") < order.index("quit"), \
        "session state must be persisted BEFORE QApplication.quit()"


def test_restart_writes_open_tabs_to_disk(qapp, tmp_path, monkeypatch):
    """End-to-end: clicking Restart writes the open task tabs to
    wizard_state.json via the REAL snapshot/persistence chain.

    This is the exact scenario the user reported — before the fix, restart_app
    quit without saving and the file was never written for this session."""
    import iCharlotte as ich
    from icharlotte_core.ui.wizard.persistence import WizardStatePersistence

    stub = types.SimpleNamespace()
    stub.case_path = str(tmp_path)
    stub.file_number = "1234"
    stub.agent_runners = []
    stub.tabs = QTabWidget()
    stub.tabs.addTab(_FakeTaskTab(task_id="summarize_documents", page_idx=0),
                     "Summarize Documents")
    for m in ("restart_app", "_persist_session_state",
              "_save_wizard_state_for_current_case", "_snapshot_open_task_tabs",
              "_iter_task_tabs", "_relpath_under"):
        setattr(stub, m, getattr(ich.MainWindow, m).__get__(stub, type(stub)))
    # Non-wizard parts of the persistence path → harmless no-ops.
    stub.save_status_history = lambda: None
    stub.chat_tab = None
    stub.discovery_tab = None
    # Don't actually spawn a new process or quit Qt.
    monkeypatch.setattr(ich.subprocess, "Popen", lambda *a, **k: None)
    monkeypatch.setattr(ich, "QApplication", types.SimpleNamespace(quit=lambda: None))

    stub.restart_app()

    open_tabs = WizardStatePersistence(str(tmp_path)).get_open_tabs()
    assert [t["task_id"] for t in open_tabs] == ["summarize_documents"], \
        "restart must write the open task tabs to disk"


def test_persist_session_state_saves_wizard_tabs(qapp):
    """The shared persistence method writes the open-tab snapshot to disk."""
    stub = types.SimpleNamespace()
    saved = []
    stub._save_wizard_state_for_current_case = lambda: saved.append(True)
    stub.save_status_history = lambda: None
    stub.chat_tab = None
    stub.discovery_tab = None
    stub._iter_task_tabs = lambda: []
    _bind("_persist_session_state", stub)

    stub._persist_session_state()

    assert saved == [True], "_persist_session_state must save the wizard tab snapshot"


# --- Defense in depth: opportunistic, non-destructive snapshots --------------
# Force-kills/freezes bypass BOTH closeEvent and restart_app, so nothing is
# saved. The fix saves the open-tab snapshot whenever the tab set or its
# contents change (open/close/complete). That snapshot must be NON-destructive:
# unlike the shutdown path, it must not cancel a tab's running worker.


def test_snapshot_does_not_cancel_running_worker_when_flag_false(qapp, tmp_path):
    worker = _FakeWorker()
    tab = _FakeTaskTab(page_idx=1, worker=worker)  # PAGE_STATUS, mid-run
    stub = _snapshot_stub(tmp_path, tab)

    snaps = stub._snapshot_open_task_tabs(cancel_running=False)

    assert worker.cancelled is False, \
        "opportunistic snapshot must NOT cancel a running worker"
    assert snaps and snaps[0]["page"] == "settings"


def test_snapshot_cancels_running_worker_by_default(qapp, tmp_path):
    """Shutdown path keeps cancelling running workers (least surprise)."""
    worker = _FakeWorker()
    tab = _FakeTaskTab(page_idx=1, worker=worker)
    stub = _snapshot_stub(tmp_path, tab)

    stub._snapshot_open_task_tabs()  # default cancel_running=True

    assert worker.cancelled is True


def test_snapshot_captures_output_page_for_contents_restore(qapp, tmp_path):
    """A completed tab (on the output page) is captured with its output_path,
    so its contents restore after a force-kill."""
    tab = _FakeTaskTab(page_idx=2)  # PAGE_OUTPUT
    tab.output_page.output_path = os.path.join(str(tmp_path), "out.docx")
    stub = _snapshot_stub(tmp_path, tab)

    snaps = stub._snapshot_open_task_tabs(cancel_running=False)

    assert snaps[0]["page"] == "output"
    assert snaps[0]["output_path"] == "out.docx"  # case-relative


def test_persist_open_tabs_is_non_destructive(qapp):
    """The opportunistic-save seam saves with cancel_running=False."""
    import iCharlotte as ich
    stub = types.SimpleNamespace()
    saved = []
    stub._save_wizard_state_for_current_case = \
        lambda cancel_running=True: saved.append(cancel_running)
    stub._persist_open_tabs = \
        ich.MainWindow._persist_open_tabs.__get__(stub, type(stub))

    stub._persist_open_tabs()

    assert saved == [False], \
        "_persist_open_tabs must save without cancelling running workers"


def test_persist_open_tabs_soon_defers_non_destructive_save(qapp, monkeypatch):
    """Completion uses a DEFERRED save: the tab emits task_completed before it
    switches to its output page, so deferring one event-loop turn lets the
    snapshot capture the output page (and its output_path)."""
    import iCharlotte as ich
    stub = types.SimpleNamespace()
    saved = []
    stub._save_wizard_state_for_current_case = \
        lambda cancel_running=True: saved.append(cancel_running)
    for m in ("_persist_open_tabs", "_persist_open_tabs_soon"):
        setattr(stub, m, getattr(ich.MainWindow, m).__get__(stub, type(stub)))

    captured = []
    monkeypatch.setattr(ich.QTimer, "singleShot", lambda ms, cb: captured.append((ms, cb)))

    stub._persist_open_tabs_soon()

    assert captured, "completion save must be deferred via QTimer.singleShot"
    assert captured[0][0] == 0, "deferral should be one event-loop turn (0 ms)"
    assert saved == [], "nothing saved until the deferred callback fires"
    captured[0][1]()  # fire the deferred callback
    assert saved == [False], "deferred save must be non-destructive"


def test_closing_task_tab_persists_open_tabs(qapp):
    import iCharlotte as ich
    stub = types.SimpleNamespace()
    stub.tabs = QTabWidget()
    tab = QWidget()
    tab.setProperty("wizard_task_id", "summarize_documents")
    stub.tabs.addTab(tab, "Summarize")
    calls = []
    stub._persist_open_tabs = lambda: calls.append(True)
    stub._on_tab_close_requested = \
        ich.MainWindow._on_tab_close_requested.__get__(stub, type(stub))

    stub._on_tab_close_requested(0)

    assert calls == [True], "closing a task tab must persist the open-tab snapshot"
