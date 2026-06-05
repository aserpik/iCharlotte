"""Behavioral tests for the Index-tab reveal / re-hide wiring (Task 6).

These exercise the REAL MainWindow methods on a minimal fixture built via
__new__ — the full MainWindow.__init__ starts Outlook/docket monitors that can
hang a headless run, so we bypass it and inject only the attributes the methods
under test actually use (self.tabs, self.index_tab, self.case_path,
self.file_number, self.mode_controller).
"""
import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

pytest.importorskip("pytestqt")

from types import SimpleNamespace

from PySide6.QtWidgets import QTabBar, QTabWidget, QWidget

import iCharlotte
from iCharlotte import MainWindow


def _make(qtbot, *, is_wizard, file_number=""):
    mw = MainWindow.__new__(MainWindow)  # skip the heavy __init__
    tabs = QTabWidget()
    tabs.setTabsClosable(True)
    qtbot.addWidget(tabs)
    tabs.addTab(QWidget(), "Master List")
    tabs.addTab(QWidget(), "Wizard")
    index_tab = QWidget()
    tabs.addTab(index_tab, "Index")
    mw.tabs = tabs
    mw.index_tab = index_tab
    mw.case_path = "C:/case"
    mw.file_number = file_number
    mw.mode_controller = SimpleNamespace(is_wizard=is_wizard)
    return mw, tabs, index_tab


def _close_btn(tabs, idx):
    bar = tabs.tabBar()
    return (bar.tabButton(idx, QTabBar.ButtonPosition.RightSide)
            or bar.tabButton(idx, QTabBar.ButtonPosition.LeftSide))


def test_reveal_shows_and_selects_index(qtbot):
    mw, tabs, index_tab = _make(qtbot, is_wizard=True)
    idx = tabs.indexOf(index_tab)
    tabs.setTabVisible(idx, False)
    mw._reveal_index_tab()
    assert tabs.isTabVisible(idx)
    assert tabs.currentIndex() == idx


def test_reveal_without_case_noops(qtbot, monkeypatch):
    mw, tabs, index_tab = _make(qtbot, is_wizard=True)
    mw.case_path = ""
    # The no-case guard pops a modal dialog; stub it so the test can't block.
    monkeypatch.setattr(iCharlotte.QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    idx = tabs.indexOf(index_tab)
    tabs.setTabVisible(idx, False)
    mw._reveal_index_tab()
    assert not tabs.isTabVisible(idx)


def test_index_x_rehides_not_destroys_in_wizard(qtbot):
    mw, tabs, index_tab = _make(qtbot, is_wizard=True)
    idx = tabs.indexOf(index_tab)
    tabs.setTabVisible(idx, True)
    n_before = tabs.count()
    mw._on_tab_close_requested(idx)
    assert tabs.count() == n_before              # singleton not removed
    assert not tabs.isTabVisible(idx)            # just hidden
    assert tabs.tabText(tabs.currentIndex()) == "Wizard"


def test_index_close_button_only_visible_in_wizard(qtbot):
    mw, tabs, index_tab = _make(qtbot, is_wizard=True)
    idx = tabs.indexOf(index_tab)
    tabs.setTabVisible(idx, True)
    mw._hide_fixed_close_buttons()
    btn = _close_btn(tabs, idx)
    # isHidden() reflects the explicit setVisible() flag (independent of whether
    # the unshown QTabWidget's ancestor chain is on screen).
    assert btn is not None and not btn.isHidden()

    # Advanced mode: the Index tab is permanent — no "x".
    mw.mode_controller = SimpleNamespace(is_wizard=False)
    mw._hide_fixed_close_buttons()
    btn2 = _close_btn(tabs, idx)
    assert btn2 is None or btn2.isHidden()
