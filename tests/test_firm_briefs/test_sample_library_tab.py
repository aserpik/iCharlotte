# tests/test_firm_briefs/test_sample_library_tab.py
import pytest
pytest.importorskip("PySide6")
from icharlotte_core.ui.dialogs_sample_library import SampleLibraryTab


def test_roots_add_remove_persist(tmp_path, qtbot, monkeypatch):
    # A fresh tab seeds from config.FIRM_BRIEFS_ROOTS; force it empty so this test
    # is independent of the machine's real configured roots.
    monkeypatch.setattr("icharlotte_core.config.FIRM_BRIEFS_ROOTS", [])
    cfg = str(tmp_path / "roots.json")
    tab = SampleLibraryTab(roots_config_path=cfg)
    qtbot.addWidget(tab)
    tab.add_root_programmatic(r"C:\lib\5800")
    tab.add_root_programmatic(r"C:\lib\3800")
    assert set(tab.roots()) == {r"C:\lib\5800", r"C:\lib\3800"}
    # persisted -> a fresh tab reads them back
    tab2 = SampleLibraryTab(roots_config_path=cfg)
    qtbot.addWidget(tab2)
    assert set(tab2.roots()) == {r"C:\lib\5800", r"C:\lib\3800"}
    tab.remove_root_programmatic(r"C:\lib\3800")
    assert tab.roots() == [r"C:\lib\5800"]


def test_stats_text_from_index(tmp_path, qtbot):
    tab = SampleLibraryTab(roots_config_path=str(tmp_path / "roots.json"))
    qtbot.addWidget(tab)
    class FakeIdx:
        def stats(self): return {"briefs": 547, "citations": 3349}
    txt = tab.stats_text(FakeIdx())
    assert "547" in txt and "3349" in txt


def test_stats_text_no_index(tmp_path, qtbot):
    tab = SampleLibraryTab(roots_config_path=str(tmp_path / "roots.json"))
    qtbot.addWidget(tab)
    assert "not built" in tab.stats_text(None).lower()
