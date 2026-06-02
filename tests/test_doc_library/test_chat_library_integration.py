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
