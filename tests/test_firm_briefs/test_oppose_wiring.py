from icharlotte_core.ui.wizard.pages import oppose_motion_page as omp


def test_make_firm_provider_none_when_index_absent(monkeypatch):
    # Force the "no built index" condition regardless of the machine's real index.
    monkeypatch.setattr("icharlotte_core.firm_briefs.factory.make_index", lambda **k: None)
    prov = omp._make_firm_provider(corpus=None)
    assert prov is None


def test_make_firm_provider_builds_when_index_present(monkeypatch):
    class FakeIndex: ...
    monkeypatch.setattr("icharlotte_core.firm_briefs.factory.make_index",
                        lambda **k: FakeIndex())
    prov = omp._make_firm_provider(corpus="CORPUS")
    assert prov is not None
    assert prov.corpus == "CORPUS"
