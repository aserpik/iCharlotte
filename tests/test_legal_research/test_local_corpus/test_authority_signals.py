from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.models import CaseRecord
from icharlotte_core.legal_research.local_corpus.authority_signals import build_signals


def _insert(con, rec: CaseRecord):
    row = rec.to_row()
    con.execute(
        "INSERT INTO cases (%s) VALUES (%s)" % (
            ",".join(row.keys()), ",".join(["?"] * len(row))),
        list(row.values()),
    )


def test_build_signals_counts_inbound_and_latest_year():
    con = schema.connect(":memory:")
    schema.create_schema(con)
    # case A is cited by B (2023) and C (2019)
    _insert(con, CaseRecord(case_uid="cap:A", source="cap", citation="30 Cal. 4th 43",
                            decision_date="2003-01-01", year="2003"))
    _insert(con, CaseRecord(case_uid="cap:B", source="cap", citation="40 Cal. 4th 1",
                            decision_date="2023-01-01", year="2023", cites_to=["30 Cal. 4th 43"]))
    _insert(con, CaseRecord(case_uid="cap:C", source="cap", citation="35 Cal. 4th 9",
                            decision_date="2019-01-01", year="2019", cites_to=["30 Cal. 4th 43"]))
    con.commit()

    build_signals(con)

    row = con.execute("SELECT citation_count, latest_citing_year FROM cases WHERE case_uid='cap:A'").fetchone()
    assert row["citation_count"] == 2
    assert row["latest_citing_year"] == "2023"
