"""Chat streaming should follow the bottom only when the user is parked there.

Regression test for the "scrolling up during streaming yanks me back to the
bottom" bug. The desired behavior (every modern chat UI) is:

* If the view is parked at the bottom, new streamed tokens keep it pinned to
  the bottom (the response "follows").
* If the user has scrolled up to read earlier text, streamed tokens must NOT
  move the viewport — the user stays where they scrolled to.
* Scrolling back down to the bottom re-arms the follow behavior.

These run headless via the offscreen Qt platform so no real window appears.
"""
import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QApplication


def _app():
    return QApplication.instance() or QApplication([])


def _make_scrollable_chat():
    """Build a ChatTab whose chat_history has a real, large scroll range."""
    from icharlotte_core.ui.tabs import ChatTab

    tab = ChatTab()
    tab.resize(640, 320)
    tab.show()
    ch = tab.chat_history
    # Pour in enough content that the document is far taller than the viewport.
    for i in range(150):
        ch.append(f"filler line {i} ............................................")
    QApplication.instance().processEvents()
    return tab, ch


class TestChatStreamFollow(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = _app()

    def test_scrollable_precondition(self):
        tab, ch = _make_scrollable_chat()
        self.assertGreater(
            ch.verticalScrollBar().maximum(), 50,
            "Test setup failed to create a scrollable chat history",
        )

    def test_follows_when_parked_at_bottom(self):
        tab, ch = _make_scrollable_chat()
        sb = ch.verticalScrollBar()

        sb.setValue(sb.maximum())  # park at the bottom
        self.app.processEvents()

        tab.on_streaming_token("hello world ")
        self.app.processEvents()

        # Still pinned to the (now larger) bottom.
        self.assertGreaterEqual(sb.value(), sb.maximum() - 4)
        self.assertIn("hello world", ch.toPlainText())

    def test_does_not_yank_when_scrolled_up(self):
        tab, ch = _make_scrollable_chat()
        sb = ch.verticalScrollBar()

        sb.setValue(0)  # scroll to the very top
        self.app.processEvents()
        self.assertEqual(sb.value(), 0)

        for _ in range(6):
            tab.on_streaming_token("streamed-word ")
        self.app.processEvents()

        # The viewport must NOT have jumped to the bottom.
        self.assertEqual(sb.value(), 0, "Streaming yanked the scrolled-up user to the bottom")
        # ...but the text was still inserted into the document.
        self.assertIn("streamed-word streamed-word", ch.toPlainText())

    def test_scrolling_back_to_bottom_rearms_follow(self):
        tab, ch = _make_scrollable_chat()
        sb = ch.verticalScrollBar()

        # Scroll up: tokens should be ignored for scrolling.
        sb.setValue(0)
        self.app.processEvents()
        tab.on_streaming_token("aaa ")
        self.app.processEvents()
        self.assertEqual(sb.value(), 0)

        # User scrolls back to the bottom -> following re-arms.
        sb.setValue(sb.maximum())
        self.app.processEvents()
        tab.on_streaming_token("bbb ")
        self.app.processEvents()
        self.assertGreaterEqual(sb.value(), sb.maximum() - 4)


if __name__ == "__main__":
    unittest.main()
