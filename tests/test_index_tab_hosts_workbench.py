import pytest
pytest.importorskip("pytestqt")
from icharlotte_core.ui.tabs import IndexTab

def test_indextab_embeds_workbench(qtbot):
    tab = IndexTab()
    qtbot.addWidget(tab)
    assert hasattr(tab, "workbench")
    # The workbench owns the table now; IndexTab should not duplicate it.
    assert tab.workbench.doc_table.columnCount() == 6
