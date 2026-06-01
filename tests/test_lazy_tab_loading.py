"""Tests for lazy per-tab loading on case switch (iCharlotte.MainWindow).

The lazy-load helpers (`_schedule_lazy_tab_reloads`, `_maybe_load_current_tab`,
`_on_tab_changed`) only touch ``self.tabs`` (a QTabWidget), ``self._lazy_loaders``,
``self._tabs_pending_reload`` and ``self.file_number``. We bind the real methods
from MainWindow onto a lightweight stub so we exercise the actual implementation
without standing up the full main window / a real case folder.
"""
import types

import pytest

pytestqt = pytest.importorskip("pytestqt")  # noqa: F841  (ensures a QApplication)
from PySide6.QtWidgets import QTabWidget, QWidget  # noqa: E402

import iCharlotte  # noqa: E402


class _Stub:
    """Minimal stand-in carrying the attributes the lazy-load methods use."""

    def __init__(self):
        self.tabs = QTabWidget()
        self.file_number = "12345"
        self._lazy_loaders = {}
        self._tabs_pending_reload = set()
        self.load_calls = []  # (widget_label, file_number) per loader invocation

    # bind the real methods under test
    _register_lazy_tabs = iCharlotte.MainWindow._register_lazy_tabs
    _schedule_lazy_tab_reloads = iCharlotte.MainWindow._schedule_lazy_tab_reloads
    _maybe_load_current_tab = iCharlotte.MainWindow._maybe_load_current_tab
    _on_tab_changed = iCharlotte.MainWindow._on_tab_changed


def _make_stub_with_tabs(qtbot, lazy_labels, plain_labels=()):
    """Build a stub whose tab widget has lazy tabs (loaders recorded) and
    optional plain tabs (no loader). Returns (stub, {label: widget})."""
    stub = _Stub()
    qtbot.addWidget(stub.tabs)
    widgets = {}

    for label in plain_labels:
        w = QWidget()
        widgets[label] = w
        stub.tabs.addTab(w, label)

    for label in lazy_labels:
        w = QWidget()
        widgets[label] = w
        stub.tabs.addTab(w, label)

        def _loader(fn, _lbl=label):
            stub.load_calls.append((_lbl, fn))

        stub._lazy_loaders[w] = _loader

    # wire currentChanged exactly as the real app does
    stub.tabs.currentChanged.connect(stub._on_tab_changed)
    return stub, widgets


def test_only_visible_tab_loads_on_switch(qtbot):
    """Scheduling a reload loads only the currently-visible lazy tab, not the
    hidden ones."""
    stub, widgets = _make_stub_with_tabs(
        qtbot, lazy_labels=["Index", "Chat", "Email"], plain_labels=["Case View"]
    )
    # Case View (a non-lazy tab) is active, mimicking a real case switch.
    stub.tabs.setCurrentWidget(widgets["Case View"])
    stub.file_number = "99999"  # caller points at new case before scheduling
    stub.load_calls.clear()

    stub._schedule_lazy_tab_reloads()

    assert stub.load_calls == [], "no lazy tab is visible, so none should load"
    # all lazy tabs remain pending
    assert stub._tabs_pending_reload == {
        widgets["Index"], widgets["Chat"], widgets["Email"]
    }


def test_visible_lazy_tab_loads_immediately(qtbot):
    """If a lazy tab happens to be the active tab when the case switches, it
    loads right away."""
    stub, widgets = _make_stub_with_tabs(qtbot, lazy_labels=["Index", "Chat"])
    stub.tabs.setCurrentWidget(widgets["Chat"])
    stub.file_number = "99999"
    stub.load_calls.clear()

    stub._schedule_lazy_tab_reloads()

    assert stub.load_calls == [("Chat", "99999")]
    assert widgets["Chat"] not in stub._tabs_pending_reload
    assert widgets["Index"] in stub._tabs_pending_reload


def test_tab_loads_when_shown_and_not_again(qtbot):
    """A deferred tab loads the first time it is shown, and not on re-visits."""
    stub, widgets = _make_stub_with_tabs(
        qtbot, lazy_labels=["Index", "Chat"], plain_labels=["Case View"]
    )
    stub.tabs.setCurrentWidget(widgets["Case View"])
    stub.file_number = "99999"
    stub._schedule_lazy_tab_reloads()
    stub.load_calls.clear()

    # Show Chat -> loads once.
    stub.tabs.setCurrentWidget(widgets["Chat"])
    assert stub.load_calls == [("Chat", "99999")]

    # Leave and come back -> no reload.
    stub.tabs.setCurrentWidget(widgets["Case View"])
    stub.tabs.setCurrentWidget(widgets["Chat"])
    assert stub.load_calls == [("Chat", "99999")]


def test_each_pending_tab_loads_once(qtbot):
    """Visiting several deferred tabs loads each exactly once, in order."""
    stub, widgets = _make_stub_with_tabs(
        qtbot, lazy_labels=["Index", "Chat", "Email"], plain_labels=["Case View"]
    )
    stub.tabs.setCurrentWidget(widgets["Case View"])
    stub.file_number = "99999"
    stub._schedule_lazy_tab_reloads()
    stub.load_calls.clear()

    stub.tabs.setCurrentWidget(widgets["Index"])
    stub.tabs.setCurrentWidget(widgets["Email"])
    stub.tabs.setCurrentWidget(widgets["Chat"])

    assert stub.load_calls == [
        ("Index", "99999"), ("Email", "99999"), ("Chat", "99999")
    ]
    assert stub._tabs_pending_reload == set()


def test_loader_exception_does_not_propagate(qtbot):
    """A failing tab loader is logged, not raised, so the switch survives."""
    stub, widgets = _make_stub_with_tabs(qtbot, lazy_labels=["Boom"])

    def _boom(fn):
        raise RuntimeError("kaboom")

    stub._lazy_loaders[widgets["Boom"]] = _boom
    stub.tabs.setCurrentWidget(widgets["Boom"])

    # Should not raise.
    stub._schedule_lazy_tab_reloads()
    assert widgets["Boom"] not in stub._tabs_pending_reload


# --- Wizard-mode default on case load --------------------------------------

class _FakeModeController:
    """Records set_mode calls without touching the real QSettings."""

    def __init__(self, mode="advanced"):
        self.mode = mode
        self.calls = []

    @property
    def is_wizard(self):
        return self.mode == "wizard"

    def set_mode(self, mode):
        self.calls.append(mode)
        self.mode = mode


class _ModeStub:
    _index_of_tab = iCharlotte.MainWindow._index_of_tab
    _default_to_wizard_mode = iCharlotte.MainWindow._default_to_wizard_mode

    def __init__(self):
        self.tabs = QTabWidget()
        self.mode_controller = _FakeModeController()


def _mode_stub(qtbot, tab_labels):
    stub = _ModeStub()
    qtbot.addWidget(stub.tabs)
    widgets = {}
    for label in tab_labels:
        w = QWidget()
        widgets[label] = w
        stub.tabs.addTab(w, label)
    return stub, widgets


def test_default_to_wizard_sets_mode_and_selects_wizard_tab(qtbot):
    stub, widgets = _mode_stub(qtbot, ["Master List", "Wizard", "Case View"])
    # Start on a non-wizard tab in advanced mode (as after a fresh switch).
    stub.tabs.setCurrentWidget(widgets["Case View"])

    stub._default_to_wizard_mode()

    assert stub.mode_controller.calls == ["wizard"]
    assert stub.mode_controller.mode == "wizard"
    assert stub.tabs.currentWidget() is widgets["Wizard"]


def test_default_to_wizard_when_already_wizard(qtbot):
    stub, widgets = _mode_stub(qtbot, ["Master List", "Wizard", "Case View"])
    stub.mode_controller.mode = "wizard"
    stub.tabs.setCurrentWidget(widgets["Master List"])

    stub._default_to_wizard_mode()

    assert stub.mode_controller.mode == "wizard"
    assert stub.tabs.currentWidget() is widgets["Wizard"]


def test_default_to_wizard_noop_when_no_wizard_tab(qtbot):
    # Defensive: if the Wizard tab is missing, mode still flips but no crash.
    stub, widgets = _mode_stub(qtbot, ["Master List", "Case View"])
    stub.tabs.setCurrentWidget(widgets["Case View"])

    stub._default_to_wizard_mode()

    assert stub.mode_controller.mode == "wizard"
    assert stub.tabs.currentWidget() is widgets["Case View"]
