"""Smoke tests for TaskCard."""
import pytest

pytest.importorskip("pytestqt")
from PySide6.QtCore import Qt

from icharlotte_core.ui.wizard.registry import get_task
from icharlotte_core.ui.wizard.task_card import TaskCard


def test_card_displays_task_metadata(qtbot):
    spec = get_task("summarize_documents")
    card = TaskCard(spec)
    qtbot.addWidget(card)
    assert spec.title in card.title_label.text()
    assert spec.description in card.description_label.text()


def test_clicking_card_emits_signal(qtbot):
    spec = get_task("medical_records")
    card = TaskCard(spec)
    qtbot.addWidget(card)
    with qtbot.waitSignal(card.clicked, timeout=500) as blocker:
        qtbot.mouseClick(card, Qt.MouseButton.LeftButton)
    assert blocker.args == ["medical_records"]
