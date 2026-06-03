"""UI tests for the Wizard launcher grouped sections + search (spec #5)."""
import pytest

pytest.importorskip("pytestqt")
from PySide6.QtCore import Qt  # noqa: E402

from icharlotte_core.ui.wizard.registry import CATEGORY_ORDER, list_tasks  # noqa: E402
from icharlotte_core.ui.wizard.wizard_tab import WizardTab  # noqa: E402


def test_has_search_box(qtbot):
    w = WizardTab()
    qtbot.addWidget(w)
    assert w.search_box is not None


def test_default_shows_all_categories_in_order(qtbot):
    w = WizardTab()
    qtbot.addWidget(w)
    assert w.visible_categories() == CATEGORY_ORDER
    assert len(w.visible_task_ids()) == len(list_tasks())


def test_search_filters_to_matching_cards(qtbot):
    w = WizardTab()
    qtbot.addWidget(w)
    w.search_box.setText("rfp")
    ids = w.visible_task_ids()
    assert "respond_to_discovery" in ids
    assert "medical_records" not in ids


def test_search_no_match_shows_empty_state(qtbot):
    w = WizardTab()
    qtbot.addWidget(w)
    w.search_box.setText("zzzznope")
    assert w.visible_task_ids() == set()
    assert w.visible_categories() == []


def test_clearing_search_restores_all(qtbot):
    w = WizardTab()
    qtbot.addWidget(w)
    w.search_box.setText("rfp")
    w.search_box.clear()
    assert len(w.visible_task_ids()) == len(list_tasks())


def test_escape_clears_search(qtbot):
    w = WizardTab()
    qtbot.addWidget(w)
    w.search_box.setText("rfp")
    qtbot.keyPress(w.search_box, Qt.Key.Key_Escape)
    assert w.search_box.text() == ""


def test_category_header_shows_count(qtbot):
    w = WizardTab()
    qtbot.addWidget(w)
    # Medical has 4 tasks; its header text should include the count.
    assert "4" in w.category_header_text("Medical")
