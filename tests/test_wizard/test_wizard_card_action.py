"""WizardTab re-emits a card's action_requested as card_action_requested."""
import pytest

pytest.importorskip("pytestqt")

from icharlotte_core.ui.wizard.wizard_tab import WizardTab


def test_reemits_card_action(qtbot):
    w = WizardTab()
    qtbot.addWidget(w)
    with qtbot.waitSignal(w.card_action_requested, timeout=500) as blocker:
        w.cards["separate"].action_requested.emit("open_separate_index")
    assert blocker.args == ["open_separate_index"]
