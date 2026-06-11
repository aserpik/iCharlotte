import os
from types import SimpleNamespace

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

pytest.importorskip("pytestqt")

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QTabBar, QTabWidget, QWidget

import iCharlotte
from iCharlotte import MainWindow


class _FakeSummaryBrowser(QWidget):
    open_requested = Signal(str)

    def __init__(self, case_path, file_number, task_id, parent=None):
        super().__init__()
        self.case_path = case_path
        self.file_number = file_number
        self.task_id = task_id
        self.refresh_count = 0

    def refresh(self):
        self.refresh_count += 1


def _make(qtbot):
    mw = MainWindow.__new__(MainWindow)
    tabs = QTabWidget()
    tabs.setTabsClosable(True)
    qtbot.addWidget(tabs)
    tabs.addTab(QWidget(), "Master List")
    tabs.addTab(QWidget(), "Wizard")
    mw.tabs = tabs
    mw.case_path = "C:/case"
    mw.file_number = "1234.001"
    mw.mode_controller = SimpleNamespace(is_wizard=True)
    mw.index_tab = QWidget()
    return mw, tabs


def _close_btn(tabs, idx):
    bar = tabs.tabBar()
    return (
        bar.tabButton(idx, QTabBar.ButtonPosition.RightSide)
        or bar.tabButton(idx, QTabBar.ButtonPosition.LeftSide)
    )


def test_summary_card_action_opens_task_specific_browser_tab(qtbot, monkeypatch):
    monkeypatch.setattr(iCharlotte, "SummaryBrowserTab", _FakeSummaryBrowser)
    mw, tabs = _make(qtbot)

    mw._on_card_action("open_medical_records_outputs")

    assert tabs.count() == 3
    browser = tabs.widget(2)
    assert isinstance(browser, _FakeSummaryBrowser)
    assert browser.task_id == "medical_records"
    assert browser.property("summary_browser_task_id") == "medical_records"
    assert tabs.tabText(2) == "Summary Browser - Medical Records"
    assert tabs.currentIndex() == 2


def test_summary_card_action_reuses_existing_browser_tab(qtbot, monkeypatch):
    monkeypatch.setattr(iCharlotte, "SummaryBrowserTab", _FakeSummaryBrowser)
    mw, tabs = _make(qtbot)

    mw._on_card_action("open_summarize_documents_outputs")
    first = tabs.widget(2)
    mw._on_card_action("open_summarize_documents_outputs")

    assert tabs.count() == 3
    assert tabs.widget(2) is first
    assert first.refresh_count == 1


def test_summary_browser_close_button_removes_tab(qtbot, monkeypatch):
    monkeypatch.setattr(iCharlotte, "SummaryBrowserTab", _FakeSummaryBrowser)
    mw, tabs = _make(qtbot)
    mw._on_card_action("open_summarize_depositions_outputs")
    idx = tabs.currentIndex()

    mw._hide_fixed_close_buttons()
    btn = _close_btn(tabs, idx)
    assert btn is not None and not btn.isHidden()

    mw._on_tab_close_requested(idx)

    assert tabs.count() == 2
