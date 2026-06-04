"""TaskCard corner action button: presence, emission, and click isolation."""
import pytest

pytest.importorskip("pytestqt")
from PySide6.QtCore import Qt

from icharlotte_core.ui.wizard.registry import get_task
from icharlotte_core.ui.wizard.task_card import TaskCard


def test_separate_card_has_action_button(qtbot):
    card = TaskCard(get_task("separate"))
    qtbot.addWidget(card)
    assert card.action_btn is not None


def test_other_card_has_no_action_button(qtbot):
    card = TaskCard(get_task("chat"))
    qtbot.addWidget(card)
    assert card.action_btn is None


def test_button_emits_action_requested(qtbot):
    card = TaskCard(get_task("separate"))
    qtbot.addWidget(card)
    with qtbot.waitSignal(card.action_requested, timeout=500) as blocker:
        card.action_btn.click()
    assert blocker.args == ["open_separate_index"]


def test_button_click_does_not_launch_task(qtbot):
    card = TaskCard(get_task("separate"))
    qtbot.addWidget(card)
    fired = []
    card.clicked.connect(fired.append)
    qtbot.mouseClick(card.action_btn, Qt.MouseButton.LeftButton)
    assert fired == []
