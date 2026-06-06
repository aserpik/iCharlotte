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


def test_search_finds_case_through_parenthetical_passage(tmp_path):
    db = str(tmp_path / "c.db")
    vec = str(tmp_path / "v.f16")
    con = schema.connect(db)
    schema.create_schema(con)
    emb = FakeEmbedder(dim=64)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)
    idx.add(
        CaseRecord(
            case_uid="cap:aguilar",
            source="cap",
            name="Aguilar v. Atlantic Richfield Co.",
            citation="25 Cal. 4th 826",
            court="Cal.",
            decision_date="2001-06-14",
            year="2001",
            full_text="The opinion discusses asbestos and procedure.",
        ),
        [
            PassageRecord(
                passage_uid="cap:aguilar#0",
                case_uid="cap:aguilar",
                ordinal=0,
                text="The opinion discusses asbestos and procedure.",
            )
        ],
    )
    idx.finalize()
    con.close()

    con = schema.connect(db)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, resume=True)
    idx.add_passages(
        [
            PassageRecord(
                passage_uid="cap:aguilar#parenthetical:900",
                case_uid="cap:aguilar",
                ordinal=1_000_000,
                text="describing Aguilar as allocating the summary judgment burden",
                passage_type="parenthetical",
                source="courtlistener_parenthetical",
                parenthetical_id="900",
                parenthetical_score=0.95,
            )
        ],
        embed=False,
    )
    idx.finalize()
    con.close()

    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    results = corpus.search_opinions("summary judgment burden", semantic=False, max_results=5)

    assert results
    assert results[0].cluster_id == "cap:aguilar"
    assert "summary judgment burden" in results[0].snippet
    assert results[0].snippet_source == "parenthetical"
    assert results[0].snippet_parenthetical_id == "900"


def test_get_opinion_text_excludes_parenthetical_passages(tmp_path):
    db = str(tmp_path / "c.db")
    vec = str(tmp_path / "v.f16")
    con = schema.connect(db)
    schema.create_schema(con)
    emb = FakeEmbedder(dim=64)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)
    idx.add(
        CaseRecord(
            case_uid="cap:aguilar",
            source="cap",
            citation="25 Cal. 4th 826",
            full_text="primary opinion text only",
        ),
        [PassageRecord(passage_uid="cap:aguilar#0", case_uid="cap:aguilar", ordinal=0, text="primary opinion text only")],
    )
    idx.finalize()
    con.close()
    con = schema.connect(db)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, resume=True)
    idx.add_passages(
        [
            PassageRecord(
                passage_uid="cap:aguilar#parenthetical:900",
                case_uid="cap:aguilar",
                ordinal=1_000_000,
                text="external parenthetical text",
                passage_type="parenthetical",
                parenthetical_id="900",
            )
        ],
        embed=False,
    )
    idx.finalize()
    con.close()

    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)

    assert corpus.get_opinion_text("cap:aguilar") == "primary opinion text only"
    assert "external parenthetical" not in corpus.get_opinion_text("cap:aguilar")


def test_opinion_text_hit_ranks_before_parenthetical_only_hit(tmp_path):
    db = str(tmp_path / "c.db")
    vec = str(tmp_path / "v.f16")
    con = schema.connect(db)
    schema.create_schema(con)
    emb = FakeEmbedder(dim=64)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)
    idx.add(
        CaseRecord(
            case_uid="cap:opinion",
            source="cap",
            name="Opinion Hit",
            citation="1 Cal. 5th 1",
            full_text="summary judgment burden applies in the opinion text.",
        ),
        [
            PassageRecord(
                passage_uid="cap:opinion#0",
                case_uid="cap:opinion",
                ordinal=0,
                text="summary judgment burden applies in the opinion text.",
            )
        ],
    )
    idx.add(
        CaseRecord(
            case_uid="cap:parenthetical",
            source="cap",
            name="Parenthetical Hit",
            citation="2 Cal. 5th 2",
            full_text="This opinion discusses unrelated procedure.",
        ),
        [
            PassageRecord(
                passage_uid="cap:parenthetical#0",
                case_uid="cap:parenthetical",
                ordinal=0,
                text="This opinion discusses unrelated procedure.",
            )
        ],
    )
    idx.finalize()
    con.close()
    con = schema.connect(db)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, resume=True)
    idx.add_passages(
        [
            PassageRecord(
                passage_uid="cap:parenthetical#parenthetical:901",
                case_uid="cap:parenthetical",
                ordinal=1_000_000,
                text="summary judgment burden summary judgment burden summary judgment burden",
                passage_type="parenthetical",
                parenthetical_id="901",
            )
        ],
        embed=False,
    )
    idx.finalize()
    con.close()

    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    results = corpus.search_opinions("summary judgment burden", semantic=False, max_results=5)

    assert [r.cluster_id for r in results[:2]] == ["cap:opinion", "cap:parenthetical"]
    assert results[0].snippet_source == "opinion"
    assert results[1].snippet_source == "parenthetical"


def test_parenthetical_does_not_add_duplicate_vote_for_existing_keyword_hit(tmp_path):
    db = str(tmp_path / "c.db")
    vec = str(tmp_path / "v.f16")
    con = schema.connect(db)
    schema.create_schema(con)
    emb = FakeEmbedder(dim=64)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)
    idx.add(
        CaseRecord(
            case_uid="cap:strong",
            source="cap",
            name="Strong Opinion Hit",
            citation="4 Cal. 5th 4",
            full_text="summary judgment burden causation allocation appears in the opinion text.",
        ),
        [
            PassageRecord(
                passage_uid="cap:strong#0",
                case_uid="cap:strong",
                ordinal=0,
                text="summary judgment burden causation allocation appears in the opinion text.",
            )
        ],
    )
    idx.add(
        CaseRecord(
            case_uid="cap:duplicate",
            source="cap",
            name="Duplicate Parenthetical Hit",
            citation="5 Cal. 5th 5",
            full_text="summary judgment appears in the opinion text.",
        ),
        [
            PassageRecord(
                passage_uid="cap:duplicate#0",
                case_uid="cap:duplicate",
                ordinal=0,
                text="summary judgment appears in the opinion text.",
            )
        ],
    )
    idx.finalize()
    con.close()
    con = schema.connect(db)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, resume=True)
    idx.add_passages(
        [
            PassageRecord(
                passage_uid="cap:duplicate#parenthetical:902",
                case_uid="cap:duplicate",
                ordinal=1_000_000,
                text=(
                    "summary judgment burden causation allocation "
                    "summary judgment burden causation allocation"
                ),
                passage_type="parenthetical",
                parenthetical_id="902",
            )
        ],
        embed=False,
    )
    idx.finalize()
    con.close()

    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    results = corpus.search_opinions(
        "summary judgment burden causation allocation",
        semantic=False,
        max_results=5,
    )

    assert [r.cluster_id for r in results[:2]] == ["cap:strong", "cap:duplicate"]


def test_search_opinions_migrates_pre_parenthetical_schema(tmp_path):
    db = str(tmp_path / "legacy.db")
    vec = str(tmp_path / "legacy.f16")
    con = schema.connect(db)
    con.executescript(
        """
        CREATE TABLE cases (
            case_uid TEXT PRIMARY KEY,
            source TEXT NOT NULL,
            name TEXT,
            name_abbreviation TEXT,
            citation TEXT,
            parallel_citations TEXT,
            court TEXT,
            decision_date TEXT,
            year TEXT,
            docket_number TEXT,
            url TEXT,
            full_text TEXT,
            citation_count INTEGER,
            latest_citing_year TEXT,
            cites_to TEXT
        );
        CREATE TABLE passages (
            passage_uid TEXT PRIMARY KEY,
            case_uid TEXT NOT NULL,
            ordinal INTEGER NOT NULL,
            text TEXT NOT NULL,
            page_label TEXT,
            vec_row INTEGER
        );
        CREATE VIRTUAL TABLE passages_fts USING fts5(text, content='');
        """
    )
    con.execute(
        "INSERT INTO cases (case_uid, source, name, citation, full_text) VALUES (?, ?, ?, ?, ?)",
        ("cap:legacy", "cap", "Legacy Case", "3 Cal. 5th 3", "legacy summary judgment text"),
    )
    con.execute(
        "INSERT INTO passages (passage_uid, case_uid, ordinal, text, vec_row) VALUES (?, ?, ?, ?, ?)",
        ("cap:legacy#0", "cap:legacy", 0, "legacy summary judgment text", 0),
    )
    con.execute("INSERT INTO passages_fts(rowid, text) VALUES (?, ?)", (1, "legacy summary judgment text"))
    con.commit()
    con.close()
    open(vec, "wb").close()

    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=FakeEmbedder(dim=64))
    results = corpus.search_opinions("summary judgment", semantic=False, max_results=3)

    assert results
    assert results[0].cluster_id == "cap:legacy"
