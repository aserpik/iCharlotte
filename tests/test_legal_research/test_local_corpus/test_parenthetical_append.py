import csv
import io

import numpy as np

from icharlotte_core.legal_research.local_corpus import build, schema
from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder
from icharlotte_core.legal_research.local_corpus.indexer import CorpusIndexer
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord


def _csv(rows):
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=list(rows[0].keys()))
    writer.writeheader()
    for row in rows:
        writer.writerow(row)
    buf.seek(0)
    return buf


def _seed_corpus(db, vec, emb):
    con = schema.connect(db)
    schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)
    idx.add(
        CaseRecord(
            case_uid="cap:aguilar",
            source="cap",
            name="Aguilar v. Atlantic Richfield Co.",
            citation="25 Cal. 4th 826",
            decision_date="2001-06-14",
            year="2001",
            full_text="primary Aguilar opinion text",
        ),
        [
            PassageRecord(
                passage_uid="cap:aguilar#0",
                case_uid="cap:aguilar",
                ordinal=0,
                text="primary Aguilar opinion text",
            )
        ],
    )
    idx.finalize()
    con.close()


def test_append_parentheticals_adds_rows_and_metadata(tmp_path):
    db = str(tmp_path / "corpus.db")
    vec = str(tmp_path / "vectors.f16")
    emb = FakeEmbedder(dim=16)
    _seed_corpus(db, vec, emb)
    con = schema.connect(db)
    schema.set_meta(con, cl_snapshot_date="2026-03-31")
    con.close()

    opinions = _csv([
        {"id": "10", "cluster_id": "300", "plain_text": ""},
        {"id": "20", "cluster_id": "200", "plain_text": ""},
    ])
    citations = _csv([
        {"id": "1", "volume": "25", "reporter": "Cal. 4th", "page": "826", "type": "8", "cluster_id": "300"},
    ])
    parentheticals = _csv([
        {
            "id": "900",
            "text": "describing Aguilar as allocating the summary judgment burden",
            "score": "0.95",
            "described_opinion_id": "10",
            "describing_opinion_id": "20",
            "group_id": "",
        }
    ])

    summary = build.append_parentheticals_to_corpus(
        parentheticals_stream=parentheticals,
        opinions_stream=opinions,
        citations_stream=citations,
        db_path=db,
        vectors_path=vec,
        embedder=emb,
        snapshot_date="2026-03-31",
        min_score=0.5,
        max_per_case=25,
        embed=False,
        refresh_opinion_map=True,
    )

    assert summary == {"added": 1}
    con = schema.connect(db)
    row = con.execute(
        "SELECT * FROM passages WHERE parenthetical_id='900'"
    ).fetchone()
    assert row["case_uid"] == "cap:aguilar"
    assert row["passage_type"] == "parenthetical"
    assert "summary judgment burden" in row["text"]
    meta = schema.get_meta(con)
    assert meta["parentheticals_snapshot_date"] == "2026-03-31"
    assert meta["parentheticals_count"] == "1"
    assert meta["parentheticals_min_score"] == "0.5"
    assert meta["parentheticals_max_per_case"] == "25"
    assert meta["cl_snapshot_date"] == "2026-03-31"
    arr = np.memmap(vec, dtype=np.float16, mode="r").reshape(-1, 16)
    assert arr.shape[0] == con.execute("SELECT COUNT(*) FROM passages").fetchone()[0]
    assert np.allclose(arr[int(row["vec_row"])], 0.0)


def test_append_parentheticals_is_idempotent(tmp_path):
    db = str(tmp_path / "corpus.db")
    vec = str(tmp_path / "vectors.f16")
    emb = FakeEmbedder(dim=16)
    _seed_corpus(db, vec, emb)
    opinions = _csv([{"id": "10", "cluster_id": "300", "plain_text": ""}])
    citations = _csv([
        {"id": "1", "volume": "25", "reporter": "Cal. 4th", "page": "826", "type": "8", "cluster_id": "300"},
    ])
    parentheticals = [
        {
            "id": "900",
            "text": "describing Aguilar as allocating the summary judgment burden",
            "score": "0.95",
            "described_opinion_id": "10",
            "describing_opinion_id": "10",
            "group_id": "",
        }
    ]

    first = build.append_parentheticals_to_corpus(
        parentheticals_stream=_csv(parentheticals),
        opinions_stream=opinions,
        citations_stream=citations,
        db_path=db,
        vectors_path=vec,
        embedder=emb,
        snapshot_date="2026-03-31",
        refresh_opinion_map=True,
    )
    second = build.append_parentheticals_to_corpus(
        parentheticals_stream=_csv(parentheticals),
        opinions_stream=None,
        citations_stream=_csv([
            {"id": "1", "volume": "25", "reporter": "Cal. 4th", "page": "826", "type": "8", "cluster_id": "300"},
        ]),
        db_path=db,
        vectors_path=vec,
        embedder=emb,
        snapshot_date="2026-03-31",
        refresh_opinion_map=False,
    )

    assert first == {"added": 1}
    assert second == {"added": 0}
    con = schema.connect(db)
    assert con.execute(
        "SELECT COUNT(*) FROM passages WHERE parenthetical_id='900'"
    ).fetchone()[0] == 1
