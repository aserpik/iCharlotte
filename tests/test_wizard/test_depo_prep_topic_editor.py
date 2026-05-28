import pytest

pytest.importorskip("pytestqt")
pytest.importorskip("PySide6")

from PySide6.QtCore import Qt


def _topics():
    return [
        {"id": "t01", "title": "Pre-existing", "strategic_note": "Establish baseline",
         "relevant_digest_refs": [], "default_checked": True, "lawyer_added": False},
        {"id": "t02", "title": "Treatment", "strategic_note": "Highlight gaps",
         "relevant_digest_refs": [], "default_checked": True, "lawyer_added": False},
    ]


def test_topic_editor_populates_and_returns(qtbot):
    from icharlotte_core.ui.wizard.pages.depo_prep_topic_editor import TopicEditor
    w = TopicEditor()
    qtbot.addWidget(w)
    w.set_topics(_topics())
    out = w.get_topics()
    assert [t["id"] for t in out] == ["t01", "t02"]
    assert all(t["default_checked"] for t in out)


def test_topic_editor_emits_topics_changed_on_uncheck(qtbot):
    from icharlotte_core.ui.wizard.pages.depo_prep_topic_editor import TopicEditor
    w = TopicEditor()
    qtbot.addWidget(w)
    w.set_topics(_topics())
    with qtbot.waitSignal(w.topics_changed, timeout=1000):
        w.set_checked(0, False)
    out = w.get_topics()
    assert out[0]["default_checked"] is False


def test_topic_editor_add_topic_appends_lawyer_added(qtbot):
    from icharlotte_core.ui.wizard.pages.depo_prep_topic_editor import TopicEditor
    w = TopicEditor()
    qtbot.addWidget(w)
    w.set_topics(_topics())
    w.add_topic(title="Custom", strategic_note="My note")
    out = w.get_topics()
    assert len(out) == 3
    assert out[2]["title"] == "Custom"
    assert out[2]["lawyer_added"] is True
    assert out[2]["id"].startswith("t")  # auto-generated id


def test_topic_editor_remove_topic(qtbot):
    from icharlotte_core.ui.wizard.pages.depo_prep_topic_editor import TopicEditor
    w = TopicEditor()
    qtbot.addWidget(w)
    w.set_topics(_topics())
    w.remove_topic_at(0)
    out = w.get_topics()
    assert len(out) == 1
    assert out[0]["id"] == "t02"


def test_topic_editor_does_not_use_setitemwidget(qtbot):
    """MEMORY.md rule: setItemWidget breaks drag-reorder; we must use item flags."""
    from icharlotte_core.ui.wizard.pages.depo_prep_topic_editor import TopicEditor
    w = TopicEditor()
    qtbot.addWidget(w)
    w.set_topics(_topics())
    # Drag-and-drop is enabled at the QListWidget level.
    lw = w._list  # internal handle for testing
    assert lw.dragDropMode() == lw.DragDropMode.InternalMove
    # No items have widgets attached.
    for i in range(lw.count()):
        assert lw.itemWidget(lw.item(i)) is None
