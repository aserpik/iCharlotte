"""Lock the production extraction path: _default_extract must return the
ExtractResult's .text string, not the ExtractResult object."""
import types
import icharlotte_core.firm_briefs.ingest as ingest


def test_default_extract_returns_text_string(monkeypatch):
    class FakeResult:
        text = "Townsend v. Superior Court (1998) 61 Cal.App.4th 1431."

    class FakeProcessor:
        def extract_text(self, path, *a, **k):
            return FakeResult()

    fake_mod = types.SimpleNamespace(DocumentProcessor=FakeProcessor)
    monkeypatch.setitem(__import__("sys").modules, "icharlotte_core.document_processor", fake_mod)

    out = ingest._default_extract("anything.pdf")
    assert isinstance(out, str)
    assert "Townsend" in out


def test_default_extract_swallows_errors(monkeypatch):
    class BoomProcessor:
        def extract_text(self, path, *a, **k):
            raise RuntimeError("boom")

    fake_mod = types.SimpleNamespace(DocumentProcessor=BoomProcessor)
    monkeypatch.setitem(__import__("sys").modules, "icharlotte_core.document_processor", fake_mod)

    assert ingest._default_extract("x.pdf") == ""
