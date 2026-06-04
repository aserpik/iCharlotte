from types import SimpleNamespace
from icharlotte_core.ui.wizard.pages import generate_motion_page as gmp


def test_generate_firm_style_helper(monkeypatch):
    monkeypatch.setattr("icharlotte_core.firm_briefs.style.select_exemplars",
                        lambda mt, side, meta, **k: ["MOVING STYLE EXCERPT"])
    out = gmp._firm_style_exemplars("msj", "moving", SimpleNamespace(relief_requested="x", principal_arguments=[]))
    assert out == ["MOVING STYLE EXCERPT"]
