"""Smoke tests for WizardTab."""
import pytest

pytest.importorskip("pytestqt")
from PySide6.QtCore import Qt

from icharlotte_core.ui.wizard.wizard_tab import WizardTab


def test_renders_four_cards(qtbot):
    tab = WizardTab()
    qtbot.addWidget(tab)
    assert len(tab.cards) == 4


def test_card_click_emits_task_requested(qtbot):
    tab = WizardTab()
    qtbot.addWidget(tab)
    # Pick the first card and click it.
    first_card = tab.cards[0]
    with qtbot.waitSignal(tab.task_requested, timeout=500) as blocker:
        qtbot.mouseClick(first_card, Qt.MouseButton.LeftButton)
    assert blocker.args[0] == first_card.task_id
