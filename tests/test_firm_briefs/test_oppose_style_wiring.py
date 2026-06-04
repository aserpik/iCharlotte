from types import SimpleNamespace
from icharlotte_core.ui.wizard.pages import oppose_motion_page as omp


def test_firm_style_exemplars_helper_guarded(monkeypatch):
    monkeypatch.setattr("icharlotte_core.firm_briefs.style.select_exemplars",
                        lambda mt, side, meta, **k: ["FIRM STYLE EXCERPT"])
    meta = SimpleNamespace(motion_type="compel", relief_requested="x", principal_arguments=["y"])
    out = omp._firm_style_exemplars("compel", "opposition", meta)
    assert out == ["FIRM STYLE EXCERPT"]


def test_firm_style_exemplars_swallows_errors(monkeypatch):
    def boom(*a, **k): raise RuntimeError("no index")
    monkeypatch.setattr("icharlotte_core.firm_briefs.style.select_exemplars", boom)
    out = omp._firm_style_exemplars("compel", "opposition", SimpleNamespace())
    assert out == []
