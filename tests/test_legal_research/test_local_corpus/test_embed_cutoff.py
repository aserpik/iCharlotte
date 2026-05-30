"""embed_year_cutoff: old cases stay keyword-searchable with zero vectors;
recent cases get real embeddings and are semantically searchable."""
import numpy as np

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord
from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder
from icharlotte_core.legal_research.local_corpus.indexer import CorpusIndexer
from icharlotte_core.legal_research.local_corpus.corpus import LocalCaseCorpus


def _case(uid, cite, year, text):
    rec = CaseRecord(case_uid=uid, source="cap", name=uid, citation=cite,
                     decision_date=f"{year}-01-01", year=str(year), full_text=text)
    return rec, [PassageRecord(passage_uid=f"{uid}#0", case_uid=uid, ordinal=0, text=text)]


def test_cutoff_zeros_old_embeds_recent(tmp_path):
    db = str(tmp_path / "c.db"); vec = str(tmp_path / "v.f16")
    con = schema.connect(db); schema.create_schema(con)
    emb = FakeEmbedder(dim=32)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, embed_year_cutoff=1950)
    idx.add(*_case("cap:old", "1 Cal. 9", 1880, "ancient negligence duty of care"))
    idx.add(*_case("cap:new", "10 Cal. 5th 1", 2020, "modern privacy discovery limits"))
    idx.finalize()
    con.close()

    arr = np.memmap(vec, dtype=np.float16, mode="r").reshape(-1, 32)
    assert arr.shape[0] == 2                       # one vector row per passage (alignment)
    # Old case row is the zero placeholder; new case row is a real (nonzero) embedding.
    con = schema.connect(db)
    rows = {r["case_uid"]: r["vec_row"] for r in con.execute("SELECT case_uid, vec_row FROM passages")}
    con.close()
    assert np.allclose(arr[rows["cap:old"]], 0.0)
    assert not np.allclose(arr[rows["cap:new"]], 0.0)

    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    # BOTH findable by keyword.
    assert any(r.cluster_id == "cap:old" for r in corpus.search_opinions("ancient negligence", semantic=False))
    assert any(r.cluster_id == "cap:new" for r in corpus.search_opinions("modern privacy", semantic=False))
    # Semantic surfaces the recent (embedded) case, never the zero-vector old one.
    sem = corpus.search_opinions("privacy discovery", semantic=True, max_results=5)
    assert any(r.cluster_id == "cap:new" for r in sem)
