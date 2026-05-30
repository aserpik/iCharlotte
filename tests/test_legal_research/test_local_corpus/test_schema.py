import sqlite3
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord
from icharlotte_core.legal_research.local_corpus import schema


def test_caserecord_roundtrips_to_row():
    rec = CaseRecord(
        case_uid="cap:1", source="cap", name="People v. Snow",
        name_abbreviation="People v. Snow", citation="30 Cal. 4th 43",
        parallel_citations=["131 Cal. Rptr. 2d 1"], court="Cal.",
        decision_date="2003-04-03", year="2003", docket_number="S018033",
        url="https://x", full_text="full text body",
        citation_count=12, latest_citing_year="2019", cites_to=["536 U.S. 584"],
    )
    row = rec.to_row()
    assert row["case_uid"] == "cap:1"
    assert row["parallel_citations"] == '["131 Cal. Rptr. 2d 1"]'  # JSON-encoded
    back = CaseRecord.from_row(row)
    assert back.parallel_citations == ["131 Cal. Rptr. 2d 1"]
    assert back.cites_to == ["536 U.S. 584"]


def test_schema_creates_tables_and_fts():
    con = sqlite3.connect(":memory:")
    schema.create_schema(con)
    names = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type IN ('table','view')")}
    assert {"cases", "passages", "citation_edges"}.issubset(names)
    # FTS5 virtual table is queryable
    con.execute("INSERT INTO passages_fts(rowid, text) VALUES (1, 'duty of care negligence')")
    hits = con.execute("SELECT rowid FROM passages_fts WHERE passages_fts MATCH 'negligence'").fetchall()
    assert hits == [(1,)]
