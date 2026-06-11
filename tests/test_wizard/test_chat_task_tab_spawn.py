"""Tests for wizard-mode multi-chat-tab spawn and snapshot/remove safety.

The wizard "Chat" task card spawns a new ChatTab each click (titled "Chat",
"Chat (2)", ...). Each one is treated as a task tab (has wizard_task_id
property) so it's closable, gets cleared on case-switch, and is included in
the wizard snapshot/restore set.

These tests exercise the helper methods that have to handle ChatTab safely:
- `_snapshot_open_task_tabs` must include wizard-spawned ChatTab.
- `_restore_task_tabs_for_case` must rebuild wizard-spawned ChatTab.
- `_remove_all_task_tabs` must call `save_current_state` before delete.
"""
import os
import sys
import unittest
from unittest.mock import MagicMock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QApplication, QTabWidget, QWidget


class _Stub:
    """Minimal MainWindow stand-in with just enough surface for the helpers."""
    def __init__(self):
        self.tabs = QTabWidget()
        self.case_path = "C:/fake/case"
        self.file_number = "0000.000"

    def _iter_task_tabs(self):
        out = []
        for i in range(self.tabs.count()):
            w = self.tabs.widget(i)
            if w is not None and w.property("wizard_task_id") is not None:
                out.append((i, w))
        return out

    def _relpath_under(self, root, path):
        try:
            return os.path.relpath(path, root)
        except ValueError:
            return path


def _bind(method_name, stub):
    import iCharlotte as _ich
    method = getattr(_ich.MainWindow, method_name)
    setattr(stub, method_name, method.__get__(stub, _Stub))


class _FakeTaskTab(QWidget):
    """Stand-in for a regular TaskTab — exposes the attrs the snapshotter reads."""
    def __init__(self, task_id="fake", files=None, output_path=None):
        super().__init__()
        self.spec = MagicMock(task_id=task_id)
        self.files = files or []
        self.output_page = MagicMock(output_path=output_path)
        self.settings_page = MagicMock()
        self.settings_page.to_dict = MagicMock(return_value={"k": "v"})
        self._worker = None
        self.setProperty("wizard_task_id", task_id)
        self.setProperty("wizard_instance_suffix", "")

    def currentIndex(self):
        return 0  # PAGE_SETTINGS


class TestSnapshotIncludesChatTab(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_snapshot_includes_wizard_chat_task_tabs(self):
        from icharlotte_core.ui.tabs import ChatTab

        stub = _Stub()
        _bind("_snapshot_open_task_tabs", stub)

        regular = _FakeTaskTab(task_id="summarize")
        chat = ChatTab()
        chat.setProperty("wizard_task_id", "chat")
        chat.setProperty("wizard_instance_suffix", "(2)")
        chat.current_conversation_id = "conv-1"

        stub.tabs.addTab(regular, "Summarize")
        stub.tabs.addTab(chat, "Chat (2)")

        snapshots = stub._snapshot_open_task_tabs()

        self.assertEqual(len(snapshots), 2)
        self.assertEqual(snapshots[0]["task_id"], "summarize")
        self.assertEqual(snapshots[1]["task_id"], "chat")
        self.assertEqual(snapshots[1]["instance_suffix"], "(2)")
        self.assertEqual(snapshots[1]["files"], [])
        self.assertEqual(snapshots[1]["settings"], {"current_conversation_id": "conv-1"})
        self.assertEqual(snapshots[1]["page"], "settings")
        self.assertIsNone(snapshots[1]["output_path"])


class _RestoreStub(QWidget):
    """QWidget stand-in for restore helpers that construct child widgets."""

    def __init__(self, case_path):
        super().__init__()
        self.tabs = QTabWidget(self)
        self.case_path = case_path
        self.file_number = "0000.000"

    def _on_task_completed(self, entry):
        pass

    def _hide_fixed_close_buttons(self):
        pass


class TestRestoreChatTab(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_restore_rebuilds_wizard_chat_task_tab(self):
        import tempfile
        from icharlotte_core.ui.tabs import ChatTab
        from icharlotte_core.ui.wizard.persistence import WizardStatePersistence

        with tempfile.TemporaryDirectory() as case_path:
            p = WizardStatePersistence(case_path)
            p.set_open_tabs([{
                "task_id": "chat",
                "instance_suffix": "(2)",
                "files": [],
                "settings": {"current_conversation_id": "conv-1"},
                "page": "settings",
                "output_path": None,
            }])
            p.save()

            stub = _RestoreStub(case_path)
            _bind("_restore_task_tabs_for_case", stub)
            stub._restore_task_tabs_for_case()

            self.assertEqual(stub.tabs.count(), 1)
            tab = stub.tabs.widget(0)
            self.assertIsInstance(tab, ChatTab)
            self.assertEqual(tab.property("wizard_task_id"), "chat")
            self.assertEqual(tab.property("wizard_instance_suffix"), "(2)")
            self.assertEqual(stub.tabs.tabText(0), "Chat (2)")
            self.assertEqual(tab.file_number, "0000.000")
            self.assertEqual(tab.current_conversation_id, "conv-1")


class TestRemoveAllTaskTabsSavesChatState(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_chat_save_current_state_called_before_delete(self):
        from icharlotte_core.ui.tabs import ChatTab

        stub = _Stub()
        _bind("_remove_all_task_tabs", stub)

        chat_a = ChatTab()
        chat_a.setProperty("wizard_task_id", "chat")
        chat_a.setProperty("wizard_instance_suffix", "")
        chat_a.save_current_state = MagicMock()

        chat_b = ChatTab()
        chat_b.setProperty("wizard_task_id", "chat")
        chat_b.setProperty("wizard_instance_suffix", "(2)")
        chat_b.save_current_state = MagicMock()

        regular = _FakeTaskTab(task_id="summarize")

        stub.tabs.addTab(chat_a, "Chat")
        stub.tabs.addTab(regular, "Summarize")
        stub.tabs.addTab(chat_b, "Chat (2)")

        stub._remove_all_task_tabs()

        chat_a.save_current_state.assert_called_once()
        chat_b.save_current_state.assert_called_once()
        self.assertEqual(stub.tabs.count(), 0)


class TestNextInstanceSuffixForChat(unittest.TestCase):
    """Ensures the wizard chat spawn uses next_instance_suffix correctly.

    The persistent singleton 'Chat' tab always exists (visible in Advanced
    mode, hidden in Wizard mode). _open_task_tab computes the suffix against
    the full tab-bar title list, so wizard-spawned chats start at 'Chat (2)'
    — never duplicating the singleton's title.
    """

    def test_singleton_present_first_wizard_click_yields_chat_2(self):
        from icharlotte_core.ui.wizard.instance_naming import next_instance_suffix
        # Tab bar at time of first wizard click:
        titles = ["Master List", "Case View", "Status", "Index", "Chat",
                  "Email", "Email Update", "Depositions", "Wizard"]
        self.assertEqual(next_instance_suffix("Chat", titles), "(2)")

    def test_progressive_suffixes_with_singleton(self):
        from icharlotte_core.ui.wizard.instance_naming import next_instance_suffix
        titles = ["Chat"]  # singleton
        self.assertEqual(next_instance_suffix("Chat", titles), "(2)")
        titles.append("Chat (2)")
        self.assertEqual(next_instance_suffix("Chat", titles), "(3)")
        titles.append("Chat (3)")
        self.assertEqual(next_instance_suffix("Chat", titles), "(4)")

    def test_closing_wizard_chat_reuses_its_slot(self):
        """Singleton stays at 'Chat'. If 'Chat (2)' is closed but '(3)'
        remains, the next wizard click reuses '(2)'."""
        from icharlotte_core.ui.wizard.instance_naming import next_instance_suffix
        titles = ["Chat", "Chat (3)"]  # singleton + a leftover wizard chat
        self.assertEqual(next_instance_suffix("Chat", titles), "(2)")


if __name__ == "__main__":
    unittest.main()
