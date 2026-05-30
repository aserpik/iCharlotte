import io, json, zipfile

from icharlotte_core.ui.wizard.pages import oppose_motion_page as omp


def test_corpus_available_true_when_both_files_exist(tmp_path, monkeypatch):
    (tmp_path / "corpus.db").write_text("x")
    (tmp_path / "vectors.f16").write_bytes(b"\x00\x00")
    monkeypatch.setattr(omp, "CASELAW_DATA_DIR", str(tmp_path))
    assert omp._corpus_available() is True


def test_corpus_available_false_when_missing(tmp_path, monkeypatch):
    monkeypatch.setattr(omp, "CASELAW_DATA_DIR", str(tmp_path))
    assert omp._corpus_available() is False


def test_corpus_available_false_during_partial_build(tmp_path, monkeypatch):
    # corpus.db exists but vectors.f16 not yet written (mid-build) -> not ready.
    (tmp_path / "corpus.db").write_text("x")
    monkeypatch.setattr(omp, "CASELAW_DATA_DIR", str(tmp_path))
    assert omp._corpus_available() is False


def test_make_local_corpus_returns_corpus(tmp_path, monkeypatch):
    # Build a tiny real corpus so the constructor path is exercised.
    from icharlotte_core.legal_research.local_corpus import build
    from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder
    case = {"id": 1, "name": "A v. B", "decision_date": "2003-01-01",
            "citations": [{"type": "official", "cite": "30 Cal. 4th 43"}],
            "jurisdiction": {"name_long": "California"}, "cites_to": [],
            "casebody": {"opinions": [{"type": "majority", "text": "duty of care."}]}}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("json/0043-01.json", json.dumps(case))
        zf.writestr("html/0043-01.html", "<p>duty of care.</p>")
    z = tmp_path / "v.zip"; z.write_bytes(buf.getvalue())
    build.build_from_cap_zips([str(z)], db_path=str(tmp_path / "corpus.db"),
                              vectors_path=str(tmp_path / "vectors.f16"), embedder=FakeEmbedder(dim=32))

    monkeypatch.setattr(omp, "CASELAW_DATA_DIR", str(tmp_path))
    monkeypatch.setattr(omp, "_corpus_embedder", lambda: FakeEmbedder(dim=32))
    corpus = omp._make_local_corpus()
    assert corpus is not None
    assert corpus.get_opinion_text("cap:1")
