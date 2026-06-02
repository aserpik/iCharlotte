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
