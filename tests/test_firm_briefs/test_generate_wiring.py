from icharlotte_core.ui.wizard.pages import generate_motion_page as gmp


def test_generate_make_firm_provider_present(monkeypatch):
    class FakeIndex: ...
    monkeypatch.setattr("icharlotte_core.firm_briefs.factory.make_index",
                        lambda **k: FakeIndex())
    prov = gmp._make_firm_provider(corpus="C")
    assert prov is not None and prov.corpus == "C"
