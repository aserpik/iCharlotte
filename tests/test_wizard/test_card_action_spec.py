"""TaskSpec optional launcher-card corner-action fields."""
from icharlotte_core.ui.wizard.registry import get_task


def test_separate_spec_declares_index_action():
    spec = get_task("separate")
    assert spec.card_action_id == "open_separate_index"
    assert spec.card_action_glyph
    assert spec.card_action_tooltip


def test_summary_specs_declare_output_browser_actions():
    expected = {
        "summarize_documents": "open_summarize_documents_outputs",
        "summarize_discovery": "open_summarize_discovery_outputs",
        "summarize_depositions": "open_summarize_depositions_outputs",
        "medical_records": "open_medical_records_outputs",
    }
    for task_id, action_id in expected.items():
        spec = get_task(task_id)
        assert spec.card_action_id == action_id
        assert spec.card_action_glyph
        assert spec.card_action_tooltip


def test_other_spec_has_no_action():
    assert get_task("chat").card_action_id is None
