from types import SimpleNamespace
from icharlotte_core.firm_briefs import style


class FakeIndex:
    def __init__(self, cands): self._c = cands
    def style_candidates(self, qv, *, motion_type, side, k=3): return self._c


class FakeEmb:
    dim = 384
    def encode(self, texts):
        import numpy as np
        return np.ones((len(texts), 384), dtype=np.float32)


META = SimpleNamespace(relief_requested="compel further responses",
                       principal_arguments=["failed to meet and confer"])


def test_select_returns_trimmed_excerpts(tmp_path):
    idx = FakeIndex([{"path": "a.pdf"}, {"path": "b.pdf"}])
    texts = {"a.pdf": "CAPTION...\nARGUMENT\nThe motion fails because " + "x" * 50,
             "b.pdf": "no heading here just prose " + "y" * 50}
    out = style.select_exemplars("compel", "opposition", META, index=idx, embedder=FakeEmb(),
                                 extract_fn=lambda p: texts[p],
                                 cache_dir=str(tmp_path), max_chars=200)
    assert len(out) == 2
    # 'a' is trimmed to start at the ARGUMENT heading
    assert out[0].startswith("ARGUMENT")
    assert all(len(t) <= 210 for t in out)


def test_select_empty_when_no_index(monkeypatch):
    # index=None falls back to factory.make_index(); force it absent so the test
    # doesn't depend on whether a real index is built on this machine.
    monkeypatch.setattr("icharlotte_core.firm_briefs.factory.make_index", lambda **k: None)
    assert style.select_exemplars("compel", "opposition", META, index=None) == []


def test_select_empty_when_no_profile(tmp_path):
    idx = FakeIndex([{"path": "a.pdf"}])
    blank = SimpleNamespace(relief_requested="", principal_arguments=[])
    out = style.select_exemplars("compel", "opposition", blank, index=idx, embedder=FakeEmb(),
                                 extract_fn=lambda p: "text", cache_dir=str(tmp_path))
    assert out == []   # no query profile -> no selection
