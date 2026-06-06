import csv
import io
import sys

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


class _TrackedStream(io.StringIO):
    def __init__(self, name):
        super().__init__("")
        self.name = name


def test_run_parenthetical_append_without_refresh_skips_opinions_and_closes_streams(monkeypatch):
    opened = []
    captured = {}

    def fake_stream(name, date):
        stream = _TrackedStream(name)
        opened.append(stream)
        return stream

    def fake_append(**kwargs):
        captured.update(kwargs)
        return {"added": 3}

    monkeypatch.setattr(build, "_stream_cl_bulk", fake_stream)
    monkeypatch.setattr(build, "append_parentheticals_to_corpus", fake_append)

    summary = build.run_parenthetical_append(
        db_path="corpus.db",
        vectors_path="vectors.f16",
        embedder=FakeEmbedder(dim=16),
        date="2026-03-31",
        refresh_opinion_map=False,
    )

    assert summary == {"added": 3}
    assert [stream.name for stream in opened] == ["parentheticals", "citations"]
    assert captured["opinions_stream"] is None
    assert all(stream.closed for stream in opened)


def test_run_parenthetical_append_with_refresh_opens_opinions_and_closes_streams(monkeypatch):
    opened = []
    captured = {}

    def fake_stream(name, date):
        stream = _TrackedStream(name)
        opened.append(stream)
        return stream

    def fake_append(**kwargs):
        captured.update(kwargs)
        return {"added": 2}

    monkeypatch.setattr(build, "_stream_cl_bulk", fake_stream)
    monkeypatch.setattr(build, "append_parentheticals_to_corpus", fake_append)

    summary = build.run_parenthetical_append(
        db_path="corpus.db",
        vectors_path="vectors.f16",
        embedder=FakeEmbedder(dim=16),
        date="2026-03-31",
        refresh_opinion_map=True,
    )

    assert summary == {"added": 2}
    assert [stream.name for stream in opened] == ["parentheticals", "opinions", "citations"]
    assert captured["opinions_stream"].name == "opinions"
    assert all(stream.closed for stream in opened)


def test_main_source_all_does_not_run_parenthetical_append(monkeypatch, tmp_path):
    calls = []

    monkeypatch.setattr(sys, "argv", [
        "build.py",
        "--source",
        "all",
        "--data-dir",
        str(tmp_path),
        "--fts-only",
    ])
    monkeypatch.setattr(build, "OnnxEmbedder", lambda: FakeEmbedder(dim=16))
    monkeypatch.setattr(build, "_download_cap_volumes", lambda scratch: ["cap.zip"])
    monkeypatch.setattr(
        build,
        "build_from_cap_zips",
        lambda *args, **kwargs: calls.append("cap") or {"cases": 1},
    )
    monkeypatch.setattr(
        build,
        "run_cl_append",
        lambda *args, **kwargs: calls.append("cl") or {"added": 1},
    )
    monkeypatch.setattr(
        build,
        "run_parenthetical_append",
        lambda *args, **kwargs: calls.append("parentheticals") or {"added": 1},
    )

    build.main()

    assert calls == ["cap", "cl"]


def test_append_parentheticals_closes_indexer_and_connection_on_error(monkeypatch, tmp_path):
    db = str(tmp_path / "corpus.db")
    vec = str(tmp_path / "vectors.f16")
    emb = FakeEmbedder(dim=16)
    _seed_corpus(db, vec, emb)
    real_connect = schema.connect
    wrappers = []
    indexers = []

    class ConnectionWrapper:
        def __init__(self, con):
            self._con = con
            self.closed = False

        def close(self):
            self.closed = True
            self._con.close()

        def __getattr__(self, name):
            return getattr(self._con, name)

    class FailingIndexer:
        def __init__(self, con, *, vectors_path, embedder, embed, resume):
            self._vec_fh = open(vectors_path, "ab")
            indexers.append(self)

        def add_passages(self, passages, *, embed):
            raise RuntimeError("boom")

    def tracking_connect(path):
        wrapper = ConnectionWrapper(real_connect(path))
        wrappers.append(wrapper)
        return wrapper

    monkeypatch.setattr(build.schema, "connect", tracking_connect)
    monkeypatch.setattr(build.parenthetical_loader, "load_opinion_cluster_map", lambda *args, **kwargs: {})
    monkeypatch.setattr(build.parenthetical_loader, "build_cluster_case_map", lambda *args, **kwargs: {})
    monkeypatch.setattr(
        build.parenthetical_loader,
        "iter_parenthetical_passages",
        lambda **kwargs: [
            PassageRecord(
                passage_uid="cap:aguilar#parenthetical:900",
                case_uid="cap:aguilar",
                ordinal=1_000_000,
                text="summary judgment burden",
                passage_type="parenthetical",
                parenthetical_id="900",
            )
        ],
    )
    monkeypatch.setattr(build, "CorpusIndexer", FailingIndexer)

    try:
        build.append_parentheticals_to_corpus(
            parentheticals_stream=io.StringIO(""),
            opinions_stream=None,
            citations_stream=io.StringIO(""),
            db_path=db,
            vectors_path=vec,
            embedder=emb,
            snapshot_date="2026-03-31",
        )
    except RuntimeError as exc:
        assert str(exc) == "boom"
    else:
        raise AssertionError("append_parentheticals_to_corpus should have raised")

    assert wrappers[0].closed
    assert indexers[0]._vec_fh.closed


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
