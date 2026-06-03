from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord
from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder
from icharlotte_core.legal_research.local_corpus.indexer import CorpusIndexer
from icharlotte_core.legal_research.local_corpus.corpus import LocalCaseCorpus
from icharlotte_core.legal_research.models import CaseResult


def _build(tmp_path):
    db = str(tmp_path / "c.db"); vec = str(tmp_path / "v.f16")
    con = schema.connect(db); schema.create_schema(con)
    emb = FakeEmbedder(dim=64)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)
    idx.add(
        CaseRecord(case_uid="cap:1", source="cap", name="Duty v. Care",
                   citation="30 Cal. 4th 43", decision_date="2003-01-01", year="2003",
                   full_text="The duty of care in negligence is well established."),
        [PassageRecord(passage_uid="cap:1#0", case_uid="cap:1", ordinal=0, page_label="44",
                       text="The duty of care in negligence is well established.")],
    )
    idx.add(
        CaseRecord(case_uid="cap:2", source="cap", name="Privacy v. Discovery",
                   citation="10 Cal. 5th 1", decision_date="2020-01-01", year="2020",
                   full_text="Constitutional privacy limits civil discovery scope."),
        [PassageRecord(passage_uid="cap:2#0", case_uid="cap:2", ordinal=0, page_label="2",
                       text="Constitutional privacy limits civil discovery scope.")],
    )
    idx.finalize()
    con.close()
    return db, vec, emb


def test_search_returns_caseresults_for_keyword(tmp_path):
    db, vec, emb = _build(tmp_path)
    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    results = corpus.search_opinions("privacy discovery", semantic=False, max_results=5)
    assert results and isinstance(results[0], CaseResult)
    assert results[0].cluster_id == "cap:2"
    assert results[0].citation == "10 Cal. 5th 1"


def test_search_semantic_path_runs(tmp_path):
    db, vec, emb = _build(tmp_path)
    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    results = corpus.search_opinions("negligence duty", semantic=True, max_results=5)
    assert any(r.cluster_id == "cap:1" for r in results)


def test_get_opinion_text_and_lookup(tmp_path):
    db, vec, emb = _build(tmp_path)
    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    assert "duty of care" in corpus.get_opinion_text("cap:1")
    hit = corpus.lookup_by_citation("30 Cal. 4th 43")
    assert hit is not None and hit["case_uid"] == "cap:1"
    assert "negligence" in hit["full_text"]


def test_search_result_uses_short_name_abbreviation(tmp_path):
    """The corpus must return the Bluebook short name (name_abbreviation), not
    the full party caption — the caption reads badly in a brief and is
    unparseable by the citation parser (breaks the output panel's selectable
    cites)."""
    db = str(tmp_path / "c.db"); vec = str(tmp_path / "v.f16")
    con = schema.connect(db); schema.create_schema(con)
    emb = FakeEmbedder(dim=64)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)
    idx.add(
        CaseRecord(
            case_uid="cap:9", source="cap",
            name="NIDA ENGALLA et al., Plaintiffs and Respondents, v. PERMANENTE "
                 "MEDICAL GROUP, INC., et al., Defendants and Appellants",
            name_abbreviation="Engalla v. Permanente Medical Group, Inc.",
            citation="15 Cal. 4th 951", decision_date="1997-06-30", year="1997",
            full_text="Arbitration agreement and discovery scheduling cooperation."),
        [PassageRecord(passage_uid="cap:9#0", case_uid="cap:9", ordinal=0, page_label="951",
                       text="Arbitration agreement and discovery scheduling cooperation.")],
    )
    idx.finalize(); con.close()
    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    res = corpus.search_opinions("arbitration discovery cooperation", semantic=False, max_results=3)
    assert res and res[0].name == "Engalla v. Permanente Medical Group, Inc."
    assert "Plaintiffs and Respondents" not in res[0].name  # not the full caption


def test_search_works_across_threads(tmp_path):
    """Opposition research/verify fan out over a ThreadPoolExecutor. A single
    shared SQLite connection raises 'created in a thread can only be used in that
    same thread' and sinks every search to zero results — regression guard for
    thread-local connections."""
    import concurrent.futures

    db, vec, emb = _build(tmp_path)
    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    # Prime the connection on the main thread (as the wizard does before fan-out).
    assert corpus.search_opinions("privacy discovery", semantic=True, max_results=5)

    def _worker(_i):
        res = corpus.search_opinions("privacy discovery", semantic=True, max_results=5)
        return bool(res) and res[0].cluster_id == "cap:2"

    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
        outcomes = list(pool.map(_worker, range(8)))
    assert all(outcomes)  # every worker thread retrieved, none hit the cross-thread error
