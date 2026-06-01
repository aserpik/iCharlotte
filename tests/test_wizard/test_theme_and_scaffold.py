"""Tests for the wizard design system (theme) and shared scaffold
(TaskHeader / WizardTaskContainer)."""
import pytest

pytest.importorskip("pytestqt")
from PySide6.QtWidgets import QLabel, QPushButton, QWidget  # noqa: E402

from icharlotte_core.ui.wizard import theme  # noqa: E402
from icharlotte_core.ui.wizard.task_scaffold import (  # noqa: E402
    TaskHeader,
    WizardTaskContainer,
)
from icharlotte_core.ui.wizard.registry import get_task  # noqa: E402


# --- theme ---------------------------------------------------------------- #

def test_wizard_stylesheet_is_nonempty_str():
    qss = theme.wizard_stylesheet()
    assert isinstance(qss, str) and "QPushButton" in qss and theme.PRIMARY in qss


def test_button_builders_return_styled_buttons(qtbot):
    for build, needle in [
        (theme.primary_button, theme.PRIMARY),
        (theme.danger_button, theme.ERROR),
        (theme.secondary_button, theme.BORDER),
    ]:
        btn = build("Go")
        qtbot.addWidget(btn)
        assert isinstance(btn, QPushButton)
        assert btn.text() == "Go"
        assert needle in btn.styleSheet()


def test_label_builders(qtbot):
    for build in (theme.page_title, theme.section_header, theme.helper_text, theme.caption):
        lbl = build("Hello")
        qtbot.addWidget(lbl)
        assert isinstance(lbl, QLabel)
        assert lbl.text() == "Hello"


# --- TaskHeader ----------------------------------------------------------- #

def test_task_header_has_step_per_default_step(qtbot):
    header = TaskHeader(get_task("summarize_documents"))
    qtbot.addWidget(header)
    assert len(header._step_labels) == 3
    # Default current step is 0 → primary colored.
    assert theme.PRIMARY in header._step_labels[0].styleSheet()


def test_task_header_set_step_highlights_active_and_completes_earlier(qtbot):
    header = TaskHeader(get_task("summarize_documents"))
    qtbot.addWidget(header)
    header.set_step(2)
    assert theme.PRIMARY in header._step_labels[2].styleSheet()      # active
    assert theme.SUCCESS in header._step_labels[0].styleSheet()      # done
    assert theme.SUCCESS in header._step_labels[1].styleSheet()      # done


def test_task_header_set_step_clamps_out_of_range(qtbot):
    header = TaskHeader(get_task("summarize_documents"))
    qtbot.addWidget(header)
    header.set_step(99)
    assert header._current == 2  # clamped to last step


# --- WizardTaskContainer -------------------------------------------------- #

def test_container_proxies_stack_api(qtbot):
    c = WizardTaskContainer(get_task("summarize_documents"))
    qtbot.addWidget(c)
    a, b = QWidget(), QWidget()
    assert c.addWidget(a) == 0
    assert c.addWidget(b) == 1
    assert c.count() == 2
    assert c.currentIndex() == 0
    c.setCurrentIndex(1)
    assert c.currentIndex() == 1
    assert c.currentWidget() is b
    assert c.indexOf(a) == 0
    assert c.widget(1) is b


def test_container_page_change_updates_header_step(qtbot):
    c = WizardTaskContainer(get_task("summarize_documents"))
    qtbot.addWidget(c)
    c.addWidget(QWidget())  # settings
    c.addWidget(QWidget())  # status
    c.addWidget(QWidget())  # output
    c.setCurrentIndex(2)
    assert c.header._current == 2
    assert theme.PRIMARY in c.header._step_labels[2].styleSheet()


def test_container_applies_wizard_stylesheet(qtbot):
    c = WizardTaskContainer(get_task("summarize_documents"))
    qtbot.addWidget(c)
    assert "QPushButton" in c.styleSheet()
