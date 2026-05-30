import io, json, zipfile

from icharlotte_core.legal_research.local_corpus import build
from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder
from icharlotte_core.legal_research.local_corpus.corpus import LocalCaseCorpus


def _cap_zip():
    case = {"id": 1, "name": "Duty v. Care", "name_abbreviation": "Duty v. Care",
            "decision_date": "2003-01-01", "citations": [{"type": "official", "cite": "30 Cal. 4th 43"}],
            "court": {"name_abbreviation": "Cal."}, "jurisdiction": {"name_long": "California"},
            "cites_to": [], "casebody": {"opinions": [{"type": "majority",
            "text": "The duty of care in negligence is established."}]}}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("json/0043-01.json", json.dumps(case))
        zf.writestr("html/0043-01.html", '<p>The duty of care in negligence is established.</p>')
    return buf.getvalue()


def test_build_from_cap_then_search(tmp_path):
    db = str(tmp_path / "c.db"); vec = str(tmp_path / "v.f16")
    emb = FakeEmbedder(dim=48)
    zpath = tmp_path / "cal-4th-1.zip"; zpath.write_bytes(_cap_zip())

    summary = build.build_from_cap_zips([str(zpath)], db_path=db, vectors_path=vec, embedder=emb)
    assert summary["cases"] == 1

    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    results = corpus.search_opinions("negligence duty", semantic=True, max_results=5)
    assert results and results[0].cluster_id == "cap:1"
