import csv
import io

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.loaders import parenthetical_loader
from icharlotte_core.legal_research.local_corpus.models import CaseRecord


def _csv(rows):
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=list(rows[0].keys()))
    writer.writeheader()
    for row in rows:
        writer.writerow(row)
    buf.seek(0)
    return buf


def _insert_case(con, case):
    row = case.to_row()
    con.execute(
        "INSERT INTO cases (%s) VALUES (%s)"
        % (",".join(row.keys()), ",".join(["?"] * len(row))),
        list(row.values()),
    )
    con.commit()


def test_load_opinion_cluster_map_caches_snapshot_rows(tmp_path):
    con = schema.connect(str(tmp_path / "c.db"))
    schema.create_schema(con)
    opinions = _csv([
        {"id": "10", "cluster_id": "100", "plain_text": "ignored"},
        {"id": "20", "cluster_id": "200", "plain_text": "ignored"},
        {"id": "", "cluster_id": "300", "plain_text": "ignored"},
    ])

    out = parenthetical_loader.load_opinion_cluster_map(
        con,
        opinions_stream=opinions,
        snapshot_date="2026-03-31",
        refresh=True,
    )

    assert out == {"10": "100", "20": "200"}
    assert isinstance(out, parenthetical_loader.OpinionClusterLookup)
    assert out.get("10") == "100"
    assert out.get("missing", "fallback") == "fallback"
    rows = con.execute(
        "SELECT opinion_id, cluster_id, snapshot_date FROM courtlistener_opinion_map "
        "ORDER BY opinion_id"
    ).fetchall()
    assert [tuple(row) for row in rows] == [
        ("10", "100", "2026-03-31"),
        ("20", "200", "2026-03-31"),
    ]


def test_load_opinion_cluster_map_parses_courtlistener_escaped_csv(tmp_path):
    con = schema.connect(str(tmp_path / "c.db"))
    schema.create_schema(con)
    opinions = io.StringIO(
        'id,plain_text,cluster_id\n'
        '10,"The court said \\"hello, world\\" before ruling.",300\n'
        '20,"Plain row.",200\n'
    )

    out = parenthetical_loader.load_opinion_cluster_map(
        con,
        opinions_stream=opinions,
        snapshot_date="2026-03-31",
        refresh=True,
    )

    assert out == {"10": "300", "20": "200"}


def test_load_opinion_cluster_map_reuses_cache_without_stream(tmp_path):
    con = schema.connect(str(tmp_path / "c.db"))
    schema.create_schema(con)
    con.execute(
        "INSERT INTO courtlistener_opinion_map VALUES (?,?,?)",
        ("10", "100", "2026-03-31"),
    )
    con.commit()

    out = parenthetical_loader.load_opinion_cluster_map(
        con,
        opinions_stream=None,
        snapshot_date="2026-03-31",
        refresh=False,
    )

    assert out == {"10": "100"}


def test_load_opinion_cluster_map_keeps_same_opinion_across_snapshots(tmp_path):
    con = schema.connect(str(tmp_path / "c.db"))
    schema.create_schema(con)

    out_a = parenthetical_loader.load_opinion_cluster_map(
        con,
        opinions_stream=_csv([
            {"id": "10", "cluster_id": "100", "plain_text": "ignored"},
        ]),
        snapshot_date="2026-03-31",
        refresh=True,
    )
    out_b = parenthetical_loader.load_opinion_cluster_map(
        con,
        opinions_stream=_csv([
            {"id": "10", "cluster_id": "101", "plain_text": "ignored"},
        ]),
        snapshot_date="2026-06-30",
        refresh=True,
    )
    cached_a = parenthetical_loader.load_opinion_cluster_map(
        con,
        opinions_stream=None,
        snapshot_date="2026-03-31",
        refresh=False,
    )

    assert out_a == {"10": "100"}
    assert out_b == {"10": "101"}
    assert cached_a == {"10": "100"}


def test_build_cluster_case_map_prefers_direct_cl_case_then_cap_citation(tmp_path):
    con = schema.connect(str(tmp_path / "c.db"))
    schema.create_schema(con)
    _insert_case(con, CaseRecord(case_uid="cl:100", source="cl", citation="15 Cal. 5th 1"))
    _insert_case(
        con,
        CaseRecord(
            case_uid="cap:aguilar",
            source="cap",
            citation="25 Cal. 4th 826",
            parallel_citations=["107 Cal. Rptr. 2d 841"],
        ),
    )
    citations = _csv([
        {"id": "1", "volume": "15", "reporter": "Cal. 5th", "page": "1", "type": "8", "cluster_id": "100"},
        {"id": "2", "volume": "25", "reporter": "Cal. 4th", "page": "826", "type": "8", "cluster_id": "300"},
        {"id": "3", "volume": "999", "reporter": "N.Y.3d", "page": "1", "type": "8", "cluster_id": "400"},
    ])

    out = parenthetical_loader.build_cluster_case_map(con, citations_stream=citations)

    assert out["100"] == "cl:100"
    assert out["300"] == "cap:aguilar"
    assert "400" not in out


def test_build_cluster_case_map_uses_deterministic_local_citation_preference(tmp_path):
    con = schema.connect(str(tmp_path / "c.db"))
    schema.create_schema(con)
    _insert_case(con, CaseRecord(case_uid="cl:999", source="cl", citation="25 Cal. 4th 826"))
    _insert_case(con, CaseRecord(case_uid="cap:aguilar", source="cap", citation="25 Cal. 4th 826"))
    citations = _csv([
        {"id": "1", "volume": "25", "reporter": "Cal. 4th", "page": "826", "type": "8", "cluster_id": "300"},
    ])

    out = parenthetical_loader.build_cluster_case_map(con, citations_stream=citations)

    assert out["300"] == "cap:aguilar"


def test_iter_parenthetical_passages_filters_scores_and_attaches_to_local_cases():
    parentheticals = _csv([
        {
            "id": "900",
            "text": "describing Aguilar as assigning the summary judgment burden",
            "score": "0.95",
            "described_opinion_id": "10",
            "describing_opinion_id": "20",
            "group_id": "5",
        },
        {
            "id": "901",
            "text": "low quality description",
            "score": "0.10",
            "described_opinion_id": "10",
            "describing_opinion_id": "21",
            "group_id": "5",
        },
        {
            "id": "902",
            "text": "describes a missing local case",
            "score": "0.99",
            "described_opinion_id": "99",
            "describing_opinion_id": "21",
            "group_id": "5",
        },
    ])

    rows = list(
        parenthetical_loader.iter_parenthetical_passages(
            parentheticals_stream=parentheticals,
            opinion_cluster_map={"10": "300", "20": "200", "21": "201", "99": "999"},
            cluster_case_map={"300": "cap:aguilar"},
            min_score=0.5,
            max_per_case=25,
        )
    )

    assert len(rows) == 1
    passage = rows[0]
    assert passage.case_uid == "cap:aguilar"
    assert passage.passage_uid == "cap:aguilar#parenthetical:900"
    assert passage.passage_type == "parenthetical"
    assert passage.source == "courtlistener_parenthetical"
    assert passage.parenthetical_id == "900"
    assert passage.parenthetical_score == 0.95
    assert passage.described_opinion_id == "10"
    assert passage.describing_opinion_id == "20"
    assert passage.describing_cluster_id == "200"
    assert "summary judgment burden" in passage.text


def test_iter_parenthetical_passages_keeps_top_scored_rows_per_case():
    parentheticals = _csv([
        {"id": "1", "text": "score one", "score": "0.1", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
        {"id": "2", "text": "score nine", "score": "0.9", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
        {"id": "3", "text": "score five", "score": "0.5", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
    ])

    rows = list(
        parenthetical_loader.iter_parenthetical_passages(
            parentheticals_stream=parentheticals,
            opinion_cluster_map={"10": "300", "20": "200"},
            cluster_case_map={"300": "cap:aguilar"},
            min_score=0.0,
            max_per_case=2,
        )
    )

    assert [p.parenthetical_id for p in rows] == ["2", "3"]
    assert [p.ordinal for p in rows] == [1_000_000, 1_000_001]


def test_iter_parenthetical_passages_tie_cap_keeps_lowest_ids():
    parentheticals = _csv([
        {"id": "1", "text": "score nine one", "score": "0.9", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
        {"id": "2", "text": "score nine two", "score": "0.9", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
        {"id": "3", "text": "score nine three", "score": "0.9", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
    ])

    rows = list(
        parenthetical_loader.iter_parenthetical_passages(
            parentheticals_stream=parentheticals,
            opinion_cluster_map={"10": "300", "20": "200"},
            cluster_case_map={"300": "cap:aguilar"},
            min_score=0.0,
            max_per_case=2,
        )
    )

    assert [p.parenthetical_id for p in rows] == ["1", "2"]


def test_iter_parenthetical_passages_tie_cap_sorts_numeric_ids_numerically():
    parentheticals = _csv([
        {"id": "2", "text": "score nine two", "score": "0.9", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
        {"id": "10", "text": "score nine ten", "score": "0.9", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
        {"id": "3", "text": "score nine three", "score": "0.9", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
    ])

    rows = list(
        parenthetical_loader.iter_parenthetical_passages(
            parentheticals_stream=parentheticals,
            opinion_cluster_map={"10": "300", "20": "200"},
            cluster_case_map={"300": "cap:aguilar"},
            min_score=0.0,
            max_per_case=2,
        )
    )

    assert [p.parenthetical_id for p in rows] == ["2", "3"]


def test_iter_parenthetical_passages_retains_only_max_per_case_rows():
    parentheticals = _csv([
        {"id": "1", "text": "score one", "score": "0.1", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
        {"id": "2", "text": "score two", "score": "0.2", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
        {"id": "3", "text": "score three", "score": "0.3", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
        {"id": "4", "text": "score four", "score": "0.4", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
        {"id": "5", "text": "score five", "score": "0.5", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
    ])

    rows = list(
        parenthetical_loader.iter_parenthetical_passages(
            parentheticals_stream=parentheticals,
            opinion_cluster_map={"10": "300", "20": "200"},
            cluster_case_map={"300": "cap:aguilar"},
            min_score=0.0,
            max_per_case=3,
        )
    )

    assert [p.parenthetical_id for p in rows] == ["5", "4", "3"]
