"""Resumable, volume-checkpointed indexing: a crash mid-build resumes cleanly."""
import numpy as np

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord
from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder
from icharlotte_core.legal_research.local_corpus.indexer import CorpusIndexer
from icharlotte_core.legal_research.local_corpus.corpus import LocalCaseCorpus


def _case(uid, cite, text):
    rec = CaseRecord(case_uid=uid, source="cap", name=uid, citation=cite,
                     decision_date="2010-01-01", year="2010", full_text=text)
    passages = [PassageRecord(passage_uid=f"{uid}#0", case_uid=uid, ordinal=0, text=text)]
    return rec, passages


def test_resume_skips_done_volumes_keeps_alignment(tmp_path):
    db = str(tmp_path / "c.db"); vec = str(tmp_path / "v.f16")
    emb = FakeEmbedder(dim=32)

    # Run 1: ingest volume A, commit it, then simulate a crash (close, no finalize).
    con = schema.connect(db); schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, resume=False)
    idx.add(*_case("cap:1", "30 Cal. 4th 43", "negligence duty of care"))
    idx.commit_volume("volA.zip")
    idx._vec_fh.close(); con.close()        # crash: no finalize/rename

    # Run 2: resume against the same files; volume A must be skipped.
    con = schema.connect(db); schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, resume=True)
    assert idx.is_volume_done("volA.zip")
    assert not idx.is_volume_done("volB.zip")
    idx.add(*_case("cap:2", "10 Cal. 5th 1", "privacy limits discovery"))
    idx.commit_volume("volB.zip")
    idx.finalize()

    n_cases = con.execute("SELECT COUNT(*) FROM cases").fetchone()[0]
    n_pass = con.execute("SELECT COUNT(*) FROM passages").fetchone()[0]
    con.close()
    assert n_cases == 2          # both cases present, no duplication
    arr = np.memmap(vec, dtype=np.float16, mode="r").reshape(-1, 32)
    assert arr.shape[0] == n_pass   # vector rows aligned with passages across resume

    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    assert any(r.cluster_id == "cap:1" for r in corpus.search_opinions("negligence", semantic=True))
    assert any(r.cluster_id == "cap:2" for r in corpus.search_opinions("privacy discovery", semantic=True))


def test_abort_volume_rolls_back_partial_work(tmp_path):
    db = str(tmp_path / "c.db"); vec = str(tmp_path / "v.f16")
    emb = FakeEmbedder(dim=32)
    con = schema.connect(db); schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, resume=False)
    idx.add(*_case("cap:1", "30 Cal. 4th 43", "good volume"))
    idx.commit_volume("volA.zip")
    # Start a second volume, then abort it (simulate a bad zip mid-volume).
    idx.add(*_case("cap:bad", "99 Cal. 4th 1", "partial work"))
    idx.abort_volume()
    idx.finalize()
    n_cases = con.execute("SELECT COUNT(*) FROM cases").fetchone()[0]
    con.close()
    assert n_cases == 1          # aborted case left no orphan row
    arr = np.memmap(vec, dtype=np.float16, mode="r").reshape(-1, 32)
    assert arr.shape[0] == 1     # and no orphan vector
