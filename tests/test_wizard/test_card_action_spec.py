"""TaskSpec optional launcher-card corner-action fields."""
from icharlotte_core.ui.wizard.registry import get_task


def test_separate_spec_declares_index_action():
    spec = get_task("separate")
    assert spec.card_action_id == "open_separate_index"
    assert spec.card_action_glyph
    assert spec.card_action_tooltip


def test_other_spec_has_no_action():
    assert get_task("chat").card_action_id is None
