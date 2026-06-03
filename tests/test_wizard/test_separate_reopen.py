"""Reopen routes the Separate task through its builder, not a generic TaskTab."""
import pytest

pytest.importorskip("pytestqt")


def test_reopen_separate_uses_builder(qtbot, monkeypatch):
    import iCharlotte
    from icharlotte_core.ui.wizard import in_process_task_tab

    win = iCharlotte.MainWindow.__new__(iCharlotte.MainWindow)
    # Minimal stand-ins so _on_reopen_recent_task can run without full app init.
    from PySide6.QtWidgets import QTabWidget
    win.tabs = QTabWidget()
    qtbot.addWidget(win.tabs)
    win.case_path = "C:/case"
    win.file_number = "1234.001"
    win._on_task_completed = lambda *a, **k: None
    win._hide_fixed_close_buttons = lambda *a, **k: None

    calls = {}

    # The real tab exposes a task_completed signal that reopen connects to.
    from PySide6.QtCore import Signal

    class FakeTab(QTabWidget):
        task_completed = Signal()

    sentinel = FakeTab()

    def fake_builder(spec, case_path, file_number, parent):
        calls["spec_id"] = spec.task_id
        return sentinel

    monkeypatch.setattr(in_process_task_tab, "build_separate_tab", fake_builder)

    win._on_reopen_recent_task({"task_id": "separate", "files": [], "settings": {}})

    assert calls.get("spec_id") == "separate"
    assert win.tabs.count() == 1
    assert win.tabs.widget(0) is sentinel


def test_reopen_separate_cancelled_picker_adds_nothing(qtbot, monkeypatch):
    import iCharlotte
    from icharlotte_core.ui.wizard import in_process_task_tab
    from PySide6.QtWidgets import QTabWidget

    win = iCharlotte.MainWindow.__new__(iCharlotte.MainWindow)
    win.tabs = QTabWidget()
    qtbot.addWidget(win.tabs)
    win.case_path = "C:/case"
    win.file_number = "1234.001"
    win._on_task_completed = lambda *a, **k: None
    win._hide_fixed_close_buttons = lambda *a, **k: None

    monkeypatch.setattr(in_process_task_tab, "build_separate_tab",
                        lambda spec, case_path, file_number, parent: None)
    win._on_reopen_recent_task({"task_id": "separate", "files": [], "settings": {}})
    assert win.tabs.count() == 0
