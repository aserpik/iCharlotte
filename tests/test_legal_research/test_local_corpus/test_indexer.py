import numpy as np

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord
from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder
from icharlotte_core.legal_research.local_corpus.indexer import CorpusIndexer


def _case(uid, cite, text):
    rec = CaseRecord(case_uid=uid, source="cap", name=uid, citation=cite,
                     decision_date="2003-01-01", year="2003", full_text=text)
    passages = [PassageRecord(passage_uid=f"{uid}#0", case_uid=uid, ordinal=0, text=text)]
    return rec, passages


def test_indexer_writes_rows_fts_and_vectors(tmp_path):
    db = str(tmp_path / "corpus.db")
    vec = str(tmp_path / "vectors.f16")
    con = schema.connect(db)
    schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=FakeEmbedder(dim=32))
    idx.add(*_case("cap:1", "30 Cal. 4th 43", "duty of care and negligence"))
    idx.add(*_case("cap:2", "10 Cal. 5th 1", "privacy and discovery limits"))
    idx.finalize()

    assert con.execute("SELECT COUNT(*) FROM cases").fetchone()[0] == 2
    assert con.execute("SELECT COUNT(*) FROM passages").fetchone()[0] == 2
    fts = con.execute("SELECT COUNT(*) FROM passages_fts WHERE passages_fts MATCH 'privacy'").fetchone()[0]
    assert fts == 1
    # vectors.f16 has 2 rows of dim 32
    arr = np.memmap(vec, dtype=np.float16, mode="r").reshape(-1, 32)
    assert arr.shape == (2, 32)
    # vec_row assigned on passages
    rows = {r["passage_uid"]: r["vec_row"] for r in con.execute("SELECT passage_uid, vec_row FROM passages")}
    assert set(rows.values()) == {0, 1}
