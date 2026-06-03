import os
import pytest
pytest.importorskip("pytestqt")
from PySide6.QtWidgets import QApplication  # noqa: E402

from icharlotte_core.doc_library.library import DocumentLibrary  # noqa: E402
from icharlotte_core.doc_library.extract import Extracted  # noqa: E402


@pytest.fixture(scope="module")
def app():
    return QApplication.instance() or QApplication([])


def _seed_library(case_root):
    src = os.path.join(case_root, "depo.pdf")
    with open(src, "wb") as f:
        f.write(b"bytes")
    lib = DocumentLibrary(case_root)
    lib.add_entry("summarize_depositions", [src], {"party": "Plaintiff"},
                  extractor=lambda p: Extracted("DEPO BODY TEXT", 1, "pdf_native", None))
    return lib


def test_tree_populates_from_library(app, tmp_path):
    from icharlotte_core.ui.tabs import ChatTab
    _seed_library(str(tmp_path))
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)   # injected accessor
    tab._refresh_library_tree()
    assert tab.library_tree.topLevelItemCount() == 1
    top = tab.library_tree.topLevelItem(0)
    assert top.text(0) == "Plaintiff's Deposition Transcript"
    assert top.childCount() == 1  # one member


def test_add_to_library_adds_entry(app, tmp_path, qtbot, monkeypatch):
    from icharlotte_core.ui.tabs import ChatTab
    src = tmp_path / "Traffic Collision Report.txt"
    src.write_text("Officer responded to a two-vehicle collision.", encoding="utf-8")
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    # Stub the file picker so no dialog opens.
    monkeypatch.setattr(tab, "_pick_files_for_library", lambda: [str(src)])
    tab.add_to_library()
    # Wait for the background QThread to finish the add.
    qtbot.waitUntil(lambda: bool(tab._library().list_entries()), timeout=5000)
    labels = [e.label for e in tab._library().list_entries()]
    assert "Traffic Collision Report" in labels


def test_read_library_content_includes_checked_text(app, tmp_path):
    from PySide6.QtCore import Qt
    from icharlotte_core.ui.tabs import ChatTab
    _seed_library(str(tmp_path))
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    tab._refresh_library_tree()
    tab.library_tree.topLevelItem(0).setCheckState(0, Qt.CheckState.Checked)
    out = tab.read_library_content()
    assert "DEPO BODY TEXT" in out
    assert "--- FILE:" in out


def test_read_library_content_empty_when_nothing_checked(app, tmp_path):
    from icharlotte_core.ui.tabs import ChatTab
    _seed_library(str(tmp_path))
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    tab._refresh_library_tree()
    assert tab.read_library_content() == ""


def test_budget_warning_fires_when_over_limit(app, tmp_path):
    from icharlotte_core.ui.tabs import ChatTab
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    tab._context_limit_for_test = 100
    huge = "x" * 100_000  # ~25k tokens
    warn = tab._library_budget_warning(huge, history_tokens=0)
    assert warn is not None
    assert "context" in warn.lower()


def test_budget_warning_silent_when_under_limit(app, tmp_path):
    from icharlotte_core.ui.tabs import ChatTab
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    tab._context_limit_for_test = 1_000_000
    assert tab._library_budget_warning("small text", history_tokens=0) is None


def test_inline_rename_persists(app, tmp_path):
    from PySide6.QtCore import Qt
    from icharlotte_core.ui.tabs import ChatTab
    _seed_library(str(tmp_path))
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    tab._refresh_library_tree()
    top = tab.library_tree.topLevelItem(0)
    top.setText(0, "Renamed Depo")
    tab._on_library_item_changed(top, 0)
    labels = [e.label for e in tab._library().list_entries()]
    assert "Renamed Depo" in labels


def test_selection_roundtrip(app, tmp_path):
    from PySide6.QtCore import Qt
    from icharlotte_core.ui.tabs import ChatTab
    lib = _seed_library(str(tmp_path))
    entry_id = lib.list_entries()[0].id
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    tab._refresh_library_tree()
    tab.library_tree.topLevelItem(0).setCheckState(0, Qt.CheckState.Checked)
    saved = tab._collect_checked_entry_ids()
    assert entry_id in saved
    tab2 = ChatTab()
    tab2._case_root_for_library = str(tmp_path)
    tab2._refresh_library_tree()
    tab2._restore_checked_entry_ids(saved)
    assert tab2.library_tree.topLevelItem(0).checkState(0) == Qt.CheckState.Checked


def test_reset_library_entry_action(app, tmp_path):
    from PySide6.QtCore import Qt
    from icharlotte_core.ui.tabs import ChatTab
    lib = _seed_library(str(tmp_path))
    entry_id = lib.list_entries()[0].id
    lib.rename_entry(entry_id, "Custom Name")
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    tab._refresh_library_tree()
    tab._reset_library_entry(entry_id)
    assert lib.list_entries()[0].label == "Plaintiff's Deposition Transcript"
    # tree reflects the reset
    assert tab.library_tree.topLevelItem(0).text(0) == "Plaintiff's Deposition Transcript"


def test_delete_library_entry_action(app, tmp_path):
    from icharlotte_core.ui.tabs import ChatTab
    lib = _seed_library(str(tmp_path))
    entry_id = lib.list_entries()[0].id
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    tab._refresh_library_tree()
    tab._delete_library_entry(entry_id, confirm=False)
    assert lib.list_entries() == []
    assert tab.library_tree.topLevelItemCount() == 0


def test_refresh_preserves_checked_selection(app, tmp_path):
    from PySide6.QtCore import Qt
    from icharlotte_core.ui.tabs import ChatTab
    lib = _seed_library(str(tmp_path))
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    tab._refresh_library_tree()
    tab.library_tree.topLevelItem(0).setCheckState(0, Qt.CheckState.Checked)
    checked_before = tab._collect_checked_entry_ids()
    tab._refresh_library_tree()  # rebuild
    assert tab._collect_checked_entry_ids() == checked_before  # still checked


def test_refresh_open_library_trees_updates_already_open_tab(app, tmp_path):
    # Simulates a background task-completion capture landing while the Chat tab
    # is already open: the tree must refresh to show the newly-captured entry.
    from PySide6.QtWidgets import QTabWidget
    from icharlotte_core.ui.tabs import ChatTab
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    tab._refresh_library_tree()
    assert tab.library_tree.topLevelItemCount() == 0  # empty: no captures yet
    container = QTabWidget()
    container.addTab(tab, "Chat")
    _seed_library(str(tmp_path))  # capture writes an entry in the background
    ChatTab.refresh_open_library_trees(container)
    assert tab.library_tree.topLevelItemCount() == 1  # now visible without manual Refresh
