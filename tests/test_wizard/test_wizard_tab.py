"""Smoke tests for WizardTab."""
import pytest

pytest.importorskip("pytestqt")
from PySide6.QtCore import Qt

from icharlotte_core.ui.wizard.wizard_tab import WizardTab


def test_renders_all_cards(qtbot):
    from icharlotte_core.ui.wizard.registry import list_tasks
    tab = WizardTab()
    qtbot.addWidget(tab)
    assert len(tab.cards) == len(list_tasks())


def test_card_click_emits_task_requested(qtbot):
    tab = WizardTab()
    qtbot.addWidget(tab)
    # Pick the first card and click it.
    first_card = next(iter(tab.cards.values()))
    with qtbot.waitSignal(tab.task_requested, timeout=500) as blocker:
        qtbot.mouseClick(first_card, Qt.MouseButton.LeftButton)
    assert blocker.args[0] == first_card.task_id


def test_chat_card_emits_chat_task_id(qtbot):
    tab = WizardTab()
    qtbot.addWidget(tab)
    assert "chat" in tab.cards
    chat_card = tab.cards["chat"]
    with qtbot.waitSignal(tab.task_requested, timeout=500) as blocker:
        qtbot.mouseClick(chat_card, Qt.MouseButton.LeftButton)
    assert blocker.args[0] == "chat"


def test_recent_tasks_hidden_by_default(qtbot):
    tab = WizardTab()
    qtbot.addWidget(tab)
    assert tab._recent_container.isVisible() is False
    assert tab._recent_toggle_btn.text() == "Show Recent Tasks"


def test_toggle_button_shows_and_hides_recent_tasks(qtbot):
    tab = WizardTab()
    qtbot.addWidget(tab)
    tab.show()

    qtbot.mouseClick(tab._recent_toggle_btn, Qt.MouseButton.LeftButton)
    assert tab._recent_container.isVisible() is True
    assert tab._recent_toggle_btn.text() == "Hide Recent Tasks"

    qtbot.mouseClick(tab._recent_toggle_btn, Qt.MouseButton.LeftButton)
    assert tab._recent_container.isVisible() is False
    assert tab._recent_toggle_btn.text() == "Show Recent Tasks"


def test_refresh_recent_tasks_does_not_force_visibility(qtbot):
    tab = WizardTab()
    qtbot.addWidget(tab)
    tab.refresh_recent_tasks([
        {"title": "Demo", "completed_at": "2026-05-19 10:00", "task_id": "demo"},
    ])
    assert tab._recent_container.isVisible() is False
    assert tab._recent_layout.count() == 1


def test_splitter_separates_cards_from_recent_section(qtbot):
    """Cards (top) and the recent toggle/list (bottom) live in a vertical QSplitter
    so the user can drag the boundary."""
    from PySide6.QtWidgets import QSplitter
    tab = WizardTab()
    qtbot.addWidget(tab)
    assert isinstance(tab._splitter, QSplitter)
    assert tab._splitter.orientation() == Qt.Orientation.Vertical
    assert tab._splitter.count() == 2
    assert tab._splitter.childrenCollapsible() is False


def test_expanding_recent_allocates_visible_space(qtbot):
    tab = WizardTab()
    qtbot.addWidget(tab)
    tab.resize(900, 700)
    tab.show()
    qtbot.mouseClick(tab._recent_toggle_btn, Qt.MouseButton.LeftButton)
    sizes = tab._splitter.sizes()
    assert sizes[1] >= 200, f"recent section should have ~220+ px when expanded, got {sizes}"
    assert sizes[0] > 0
