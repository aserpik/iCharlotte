"""Append CL recent cases into an existing (published) corpus, in place."""
import csv, io, sqlite3, sys
import numpy as np

from icharlotte_core.legal_research.local_corpus import build, schema
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord
from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder
from icharlotte_core.legal_research.local_corpus.indexer import CorpusIndexer
from icharlotte_core.legal_research.local_corpus.corpus import LocalCaseCorpus


def _csv(rows):
    buf = io.StringIO(); w = csv.DictWriter(buf, fieldnames=list(rows[0].keys()))
    w.writeheader(); [w.writerow(r) for r in rows]; buf.seek(0); return buf


def _seed_cap_corpus(db, vec, emb):
    """Build a tiny 'CAP' corpus (embedded) to append onto."""
    con = schema.connect(db); schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)  # embed all
    idx.add(CaseRecord(case_uid="cap:1", source="cap", name="Old Duty v. Care",
                       citation="30 Cal. 4th 43", decision_date="2003-01-01", year="2003",
                       full_text="duty of care negligence"),
            [PassageRecord(passage_uid="cap:1#0", case_uid="cap:1", ordinal=0, text="duty of care negligence")])
    idx.commit_volume("cal-4th-30.zip"); idx.finalize(); con.close()


def test_append_cl_keyword_only_preserves_cap(tmp_path):
    db = str(tmp_path / "corpus.db"); vec = str(tmp_path / "vectors.f16")
    emb = FakeEmbedder(dim=32)
    _seed_cap_corpus(db, vec, emb)
    base_pass = int(np.memmap(vec, dtype=np.float16, mode="r").reshape(-1, 32).shape[0])
    assert base_pass == 1

    citations = _csv([{"id": "1", "volume": "15", "reporter": "Cal. 5th", "page": "1",
                       "type": "8", "cluster_id": "500"}])
    clusters = _csv([{"id": "500", "date_filed": "2023-06-01", "case_name": "Recent Arb v. CA",
                      "case_name_short": "Recent", "citation_count": "3",
                      "precedential_status": "Published", "docket_id": "9"}])
    opinions = _csv([{"cluster_id": "500", "plain_text": "modern arbitration unconscionability holding.",
                      "html": "", "type": "020lead"}])

    summary = build.append_cl_to_corpus(
        citations_stream=citations, clusters_stream=clusters, opinions_stream=opinions,
        db_path=db, vectors_path=vec, embedder=emb, cutoff_date="2017-01-01",
        embed=False, snapshot_date="2026-03-31",
    )
    assert summary["added"] == 1

    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    # CAP case still there + searchable.
    assert any(r.cluster_id == "cap:1" for r in corpus.search_opinions("negligence duty", semantic=False))
    # New CL case keyword-searchable.
    assert any(r.cluster_id == "cl:500" for r in corpus.search_opinions("arbitration unconscionability", semantic=False))
    # Vectors still aligned (CL passage got a zero placeholder row).
    arr = np.memmap(vec, dtype=np.float16, mode="r").reshape(-1, 32)
    import sqlite3
    npass = sqlite3.connect(db).execute("SELECT COUNT(*) FROM passages").fetchone()[0]
    assert arr.shape[0] == npass
    # CL row is a zero placeholder (keyword-only); CAP row is a real embedding.
    rows = {r[0]: r[1] for r in sqlite3.connect(db).execute("SELECT case_uid, vec_row FROM passages")}
    assert np.allclose(arr[rows["cl:500"]], 0.0)
    assert not np.allclose(arr[rows["cap:1"]], 0.0)
    meta = schema.get_meta(sqlite3.connect(db))
    assert meta["cl_snapshot_date"] == "2026-03-31"


def test_backfill_cl_embeddings_rewrites_only_zero_cl_opinion_vectors(tmp_path):
    db = str(tmp_path / "corpus.db"); vec = str(tmp_path / "vectors.f16")
    emb = FakeEmbedder(dim=32)
    _seed_cap_corpus(db, vec, emb)

    citations = _csv([{"id": "1", "volume": "15", "reporter": "Cal. 5th", "page": "1",
                       "type": "8", "cluster_id": "500"}])
    clusters = _csv([{"id": "500", "date_filed": "2023-06-01", "case_name": "Recent Arb v. CA",
                      "case_name_short": "Recent", "citation_count": "3",
                      "precedential_status": "Published", "docket_id": "9"}])
    opinions = _csv([{"cluster_id": "500", "plain_text": "modern arbitration unconscionability holding.",
                      "html": "", "type": "020lead"}])
    build.append_cl_to_corpus(
        citations_stream=citations, clusters_stream=clusters, opinions_stream=opinions,
        db_path=db, vectors_path=vec, embedder=emb, cutoff_date="2017-01-01",
        embed=False, snapshot_date="2026-03-31",
    )
    con = schema.connect(db)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, embed=True, resume=True)
    added = idx.add_passages([
        PassageRecord(
            passage_uid="cl:500#parenthetical:900",
            case_uid="cl:500",
            ordinal=1_000_000,
            text="describing Recent as an arbitration unconscionability case",
            passage_type="parenthetical",
            parenthetical_id="900",
        )
    ], embed=False)
    assert added == 1
    idx.finalize(); con.close()

    con = sqlite3.connect(db)
    rows = {
        r[0]: r[1]
        for r in con.execute("SELECT passage_uid, vec_row FROM passages ORDER BY vec_row")
    }
    con.close()
    arr = np.memmap(vec, dtype=np.float16, mode="r").reshape(-1, 32)
    cap_before = np.array(arr[rows["cap:1#0"]])
    assert np.allclose(arr[rows["cl:500#0"]], 0.0)
    assert np.allclose(arr[rows["cl:500#parenthetical:900"]], 0.0)

    summary = build.backfill_cl_embeddings(
        db_path=db,
        vectors_path=vec,
        embedder=emb,
        batch_size=2,
    )

    assert summary["eligible"] == 1
    assert summary["backfilled"] == 1
    assert summary["already_embedded"] == 0
    assert summary["remaining_zero"] == 0
    arr = np.memmap(vec, dtype=np.float16, mode="r").reshape(-1, 32)
    assert np.array_equal(arr[rows["cap:1#0"]], cap_before)
    assert not np.allclose(arr[rows["cl:500#0"]], 0.0)
    assert np.allclose(arr[rows["cl:500#parenthetical:900"]], 0.0)
    meta = schema.get_meta(sqlite3.connect(db))
    assert meta["cl_embeddings_backfilled_count"] == "1"
    assert meta["cl_embeddings_zero_remaining"] == "0"

    rerun = build.backfill_cl_embeddings(
        db_path=db,
        vectors_path=vec,
        embedder=emb,
        batch_size=2,
    )

    assert rerun["backfilled"] == 0
    assert rerun["already_embedded"] == 1


def test_backfill_cl_embeddings_can_include_parentheticals(tmp_path):
    db = str(tmp_path / "corpus.db"); vec = str(tmp_path / "vectors.f16")
    emb = FakeEmbedder(dim=32)
    _seed_cap_corpus(db, vec, emb)
    con = schema.connect(db)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, embed=True, resume=True)
    idx.add(
        CaseRecord(
            case_uid="cl:500",
            source="cl",
            name="Recent Arb v. CA",
            citation="15 Cal. 5th 1",
            decision_date="2023-06-01",
            year="2023",
            full_text="modern arbitration unconscionability holding.",
        ),
        [
            PassageRecord(
                passage_uid="cl:500#0",
                case_uid="cl:500",
                ordinal=0,
                text="modern arbitration unconscionability holding.",
            )
        ],
    )
    idx.commit_volume("manual-cl")
    idx.add_passages([
        PassageRecord(
            passage_uid="cl:500#parenthetical:900",
            case_uid="cl:500",
            ordinal=1_000_000,
            text="describing Recent as an arbitration unconscionability case",
            passage_type="parenthetical",
            parenthetical_id="900",
        )
    ], embed=False)
    idx.finalize(); con.close()

    con = sqlite3.connect(db)
    rows = {
        r[0]: r[1]
        for r in con.execute("SELECT passage_uid, vec_row FROM passages ORDER BY vec_row")
    }
    con.close()
    arr = np.memmap(vec, dtype=np.float16, mode="r").reshape(-1, 32)
    assert np.allclose(arr[rows["cl:500#parenthetical:900"]], 0.0)

    summary = build.backfill_cl_embeddings(
        db_path=db,
        vectors_path=vec,
        embedder=emb,
        batch_size=2,
        include_parentheticals=True,
    )

    assert summary["eligible"] == 2
    assert summary["backfilled"] == 1
    assert summary["already_embedded"] == 1
    arr = np.memmap(vec, dtype=np.float16, mode="r").reshape(-1, 32)
    assert not np.allclose(arr[rows["cl:500#parenthetical:900"]], 0.0)


def test_main_source_cl_embeddings_runs_backfill_only(monkeypatch, tmp_path):
    calls = []

    monkeypatch.setattr(sys, "argv", [
        "build.py",
        "--source",
        "cl-embeddings",
        "--data-dir",
        str(tmp_path),
    ])
    monkeypatch.setattr(build, "OnnxEmbedder", lambda: FakeEmbedder(dim=16))
    monkeypatch.setattr(
        build,
        "backfill_cl_embeddings",
        lambda **kwargs: calls.append(kwargs) or {
            "eligible": 1,
            "backfilled": 1,
            "already_embedded": 0,
            "remaining_zero": 0,
        },
    )
    monkeypatch.setattr(
        build,
        "run_cl_append",
        lambda **kwargs: (_ for _ in ()).throw(AssertionError("bulk CL append should not run")),
    )
    monkeypatch.setattr(
        build,
        "_download_cap_volumes",
        lambda scratch: (_ for _ in ()).throw(AssertionError("CAP download should not run")),
    )

    build.main()

    assert len(calls) == 1
    assert calls[0]["db_path"].endswith("corpus.db")
    assert calls[0]["vectors_path"].endswith("vectors.f16")
