import sqlite3

import numpy as np

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder
from icharlotte_core.legal_research.local_corpus.indexer import CorpusIndexer
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord


def test_schema_creates_parenthetical_passage_columns_and_opinion_map():
    con = sqlite3.connect(":memory:")
    schema.create_schema(con)

    passage_columns = {
        row[1] for row in con.execute("PRAGMA table_info(passages)").fetchall()
    }
    assert {
        "passage_type",
        "source",
        "parenthetical_id",
        "parenthetical_score",
        "described_opinion_id",
        "describing_opinion_id",
        "describing_cluster_id",
    }.issubset(passage_columns)

    tables = {
        row[0]
        for row in con.execute(
            "SELECT name FROM sqlite_master WHERE type IN ('table','view')"
        ).fetchall()
    }
    assert "courtlistener_opinion_map" in tables


def test_passage_record_accepts_parenthetical_metadata():
    passage = PassageRecord(
        passage_uid="cap:1#parenthetical:900",
        case_uid="cap:1",
        ordinal=1_000_000,
        text="describing the case as adopting a burden-shifting standard",
        passage_type="parenthetical",
        source="courtlistener_parenthetical",
        parenthetical_id="900",
        parenthetical_score=0.88,
        described_opinion_id="10",
        describing_opinion_id="20",
        describing_cluster_id="200",
    )

    assert passage.passage_type == "parenthetical"
    assert passage.parenthetical_id == "900"
    assert passage.parenthetical_score == 0.88
    assert passage.describing_cluster_id == "200"


def test_indexer_appends_parenthetical_passage_to_existing_case(tmp_path):
    db = str(tmp_path / "corpus.db")
    vec = str(tmp_path / "vectors.f16")
    emb = FakeEmbedder(dim=16)
    con = schema.connect(db)
    schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)
    idx.add(
        CaseRecord(
            case_uid="cap:1",
            source="cap",
            name="Aguilar v. Atlantic Richfield Co.",
            citation="25 Cal. 4th 826",
            decision_date="2001-06-14",
            year="2001",
            full_text="primary opinion text",
        ),
        [
            PassageRecord(
                passage_uid="cap:1#0",
                case_uid="cap:1",
                ordinal=0,
                text="primary opinion text",
            )
        ],
    )
    idx.finalize()
    con.close()

    con = schema.connect(db)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, resume=True)
    added = idx.add_passages(
        [
            PassageRecord(
                passage_uid="cap:1#parenthetical:900",
                case_uid="cap:1",
                ordinal=1_000_000,
                text="describing Aguilar as setting the summary judgment burden",
                passage_type="parenthetical",
                source="courtlistener_parenthetical",
                parenthetical_id="900",
                parenthetical_score=0.91,
                described_opinion_id="10",
                describing_opinion_id="20",
                describing_cluster_id="200",
            )
        ],
        embed=False,
    )
    idx.finalize()

    assert added == 1
    row = con.execute(
        "SELECT * FROM passages WHERE passage_uid=?",
        ("cap:1#parenthetical:900",),
    ).fetchone()
    assert row["passage_type"] == "parenthetical"
    assert row["parenthetical_id"] == "900"
    assert row["parenthetical_score"] == 0.91
    assert row["described_opinion_id"] == "10"
    assert row["describing_opinion_id"] == "20"
    assert row["describing_cluster_id"] == "200"

    arr = np.memmap(vec, dtype=np.float16, mode="r").reshape(-1, 16)
    assert arr.shape[0] == 2
    assert np.allclose(arr[int(row["vec_row"])], 0.0)


def test_indexer_skips_duplicate_parenthetical_passage_uid(tmp_path):
    db = str(tmp_path / "corpus.db")
    vec = str(tmp_path / "vectors.f16")
    emb = FakeEmbedder(dim=16)
    con = schema.connect(db)
    schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)
    idx.add(
        CaseRecord(case_uid="cap:1", source="cap", citation="25 Cal. 4th 826"),
        [PassageRecord(passage_uid="cap:1#0", case_uid="cap:1", ordinal=0, text="x")],
    )
    idx.finalize()
    con.close()

    passage = PassageRecord(
        passage_uid="cap:1#parenthetical:900",
        case_uid="cap:1",
        ordinal=1_000_000,
        text="summary judgment burden",
        passage_type="parenthetical",
        parenthetical_id="900",
    )
    con = schema.connect(db)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, resume=True)
    assert idx.add_passages([passage], embed=False) == 1
    idx.commit_volume("parentheticals-batch-1")
    assert idx.add_passages([passage], embed=False) == 0
    idx.finalize()

    assert con.execute(
        "SELECT COUNT(*) FROM passages WHERE parenthetical_id='900'"
    ).fetchone()[0] == 1
