import io, json, zipfile
import pytest

fastembed = pytest.importorskip("fastembed")  # skip if model deps absent


def test_real_embedder_end_to_end(tmp_path):
    from icharlotte_core.legal_research.local_corpus import build
    from icharlotte_core.legal_research.local_corpus.embedder import OnnxEmbedder
    from icharlotte_core.legal_research.local_corpus.corpus import LocalCaseCorpus

    def _zip(cid, cite, text):
        case = {"id": cid, "name": f"Case {cid}", "decision_date": "2010-01-01",
                "citations": [{"type": "official", "cite": cite}],
                "jurisdiction": {"name_long": "California"}, "cites_to": [],
                "casebody": {"opinions": [{"type": "majority", "text": text}]}}
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr(f"json/{cid:04d}-01.json", json.dumps(case))
            zf.writestr(f"html/{cid:04d}-01.html", f"<p>{text}</p>")
        p = tmp_path / f"z{cid}.zip"; p.write_bytes(buf.getvalue()); return str(p)

    zips = [
        _zip(1, "30 Cal. 4th 43", "The duty of care in negligence requires foreseeability."),
        _zip(2, "10 Cal. 5th 1", "Constitutional privacy constrains the scope of civil discovery."),
    ]
    db = str(tmp_path / "c.db"); vec = str(tmp_path / "v.f16")
    build.build_from_cap_zips(zips, db_path=db, vectors_path=vec, embedder=OnnxEmbedder())

    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=OnnxEmbedder())
    # Semantic query phrased differently from the opinion text should still hit case 2.
    results = corpus.search_opinions("can the other side inspect private records",
                                     semantic=True, max_results=2)
    assert any(r.cluster_id == "cap:2" for r in results)
