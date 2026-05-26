"""Tests for MainWindow restart tab restoration.

Verifies that `_restore_initial_tab` selects the right tab using a stable
identifier (wizard_task_id+suffix, then tab name) and only falls back to a
numeric index when no identifier matches. Numeric indices shift across
restarts because the Wizard tab insert and post-init task tabs move the
mapping, so the identifier path is the critical one.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QApplication, QTabWidget, QWidget


class _Stub:
    """Pulls `_restore_initial_tab` onto a barebones object with a real
    QTabWidget so we can exercise the lookup logic without spinning up the
    full MainWindow (which would require the whole UI tree)."""

    def __init__(self):
        self.tabs = QTabWidget()

    def _index_of_tab(self, tab_text):
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == tab_text:
                return i
        return -1


def _build_stub():
    # Bind the real method onto a stub instance.
    import iCharlotte as _ich
    stub = _Stub()
    stub._restore_initial_tab = _ich.MainWindow._restore_initial_tab.__get__(stub, _Stub)
    return stub


def _add_tab(stub, label, task_id=None, suffix=""):
    w = QWidget()
    if task_id is not None:
        w.setProperty("wizard_task_id", task_id)
        w.setProperty("wizard_instance_suffix", suffix)
    stub.tabs.addTab(w, label)
    return w


class TestRestoreInitialTab(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_task_id_match_wins_over_index(self):
        stub = _build_stub()
        _add_tab(stub, "Master List")             # 0
        _add_tab(stub, "Wizard")                  # 1
        _add_tab(stub, "Case View")               # 2
        _add_tab(stub, "Summarize", task_id="summarize", suffix="")  # 3

        # Caller passes a stale fallback index pointing at Master List —
        # task_id match should override it.
        stub._restore_initial_tab(task_id="summarize", instance_suffix="",
                                  tab_name=None, fallback_index=0)
        self.assertEqual(stub.tabs.currentIndex(), 3)

    def test_task_id_with_matching_suffix(self):
        stub = _build_stub()
        _add_tab(stub, "Master List")
        _add_tab(stub, "Summarize",   task_id="summarize", suffix="")
        _add_tab(stub, "Summarize (2)", task_id="summarize", suffix="(2)")

        stub._restore_initial_tab(task_id="summarize", instance_suffix="(2)",
                                  tab_name=None, fallback_index=None)
        self.assertEqual(stub.tabs.currentIndex(), 2)

    def test_task_id_no_match_falls_through_to_name(self):
        stub = _build_stub()
        _add_tab(stub, "Master List")
        _add_tab(stub, "Case View")

        # task_id "ghost" has no widget, but "Case View" exists by name.
        stub._restore_initial_tab(task_id="ghost", instance_suffix="",
                                  tab_name="Case View", fallback_index=None)
        self.assertEqual(stub.tabs.currentIndex(), 1)

    def test_name_match(self):
        stub = _build_stub()
        _add_tab(stub, "Master List")
        _add_tab(stub, "Wizard")
        _add_tab(stub, "Case View")

        stub._restore_initial_tab(task_id=None, instance_suffix=None,
                                  tab_name="Case View", fallback_index=None)
        self.assertEqual(stub.tabs.currentIndex(), 2)

    def test_fallback_index_when_no_identifiers(self):
        stub = _build_stub()
        _add_tab(stub, "Master List")
        _add_tab(stub, "Wizard")
        _add_tab(stub, "Case View")

        stub._restore_initial_tab(task_id=None, instance_suffix=None,
                                  tab_name=None, fallback_index=2)
        self.assertEqual(stub.tabs.currentIndex(), 2)

    def test_no_match_leaves_default(self):
        stub = _build_stub()
        _add_tab(stub, "Master List")
        _add_tab(stub, "Wizard")
        stub.tabs.setCurrentIndex(0)

        # No identifier matches, index out of bounds → leave current alone.
        stub._restore_initial_tab(task_id="ghost", instance_suffix="",
                                  tab_name="Nope", fallback_index=99)
        self.assertEqual(stub.tabs.currentIndex(), 0)

    def test_hidden_tab_not_activated(self):
        stub = _build_stub()
        _add_tab(stub, "Master List")
        _add_tab(stub, "Case View")
        stub.tabs.setTabVisible(1, False)
        stub.tabs.setCurrentIndex(0)

        stub._restore_initial_tab(task_id=None, instance_suffix=None,
                                  tab_name="Case View", fallback_index=None)
        # Visibility check should prevent activation.
        self.assertEqual(stub.tabs.currentIndex(), 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
