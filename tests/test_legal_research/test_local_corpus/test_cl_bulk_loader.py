import csv
import io

from icharlotte_core.legal_research.local_corpus.loaders import cl_bulk_loader


def _csv(rows: list[dict]) -> io.StringIO:
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=list(rows[0].keys()))
    w.writeheader()
    for r in rows:
        w.writerow(r)
    buf.seek(0)
    return buf


def test_identifies_ca_by_citation_and_filters_published_recent():
    # citations file: cluster_id -> reporter/volume/page. CA = reporter starts "Cal."
    citations = _csv([
        {"id": "1", "volume": "15", "reporter": "Cal. 5th", "page": "1", "type": "8", "cluster_id": "100"},
        {"id": "2", "volume": "70", "reporter": "Cal. App. 5th", "page": "9", "type": "8", "cluster_id": "101"},
        {"id": "3", "volume": "50", "reporter": "Cal. 3d", "page": "1", "type": "8", "cluster_id": "102"},  # old
        {"id": "4", "volume": "1", "reporter": "N.Y.3d", "page": "1", "type": "8", "cluster_id": "103"},   # not CA
        {"id": "5", "volume": "99", "reporter": "Cal. App. 5th", "page": "5", "type": "8", "cluster_id": "104"},  # unpub
    ])
    clusters = _csv([
        {"id": "100", "date_filed": "2023-06-01", "case_name": "Recent Pub v. CA",
         "case_name_short": "Recent", "citation_count": "7", "precedential_status": "Published", "docket_id": "9"},
        {"id": "101", "date_filed": "2024-01-15", "case_name": "Another Pub v. CA",
         "case_name_short": "Another", "citation_count": "2", "precedential_status": "Published", "docket_id": "9"},
        {"id": "102", "date_filed": "1990-01-01", "case_name": "Old v. CA",
         "case_name_short": "Old", "citation_count": "99", "precedential_status": "Published", "docket_id": "9"},
        {"id": "104", "date_filed": "2023-03-01", "case_name": "Unpub v. CA",
         "case_name_short": "Unpub", "citation_count": "0", "precedential_status": "Unpublished", "docket_id": "9"},
    ])
    opinions = _csv([
        {"cluster_id": "100", "plain_text": "Recent published opinion about arbitration.", "html": "", "type": "020lead"},
        {"cluster_id": "101", "plain_text": "Another published opinion about CEQA.", "html": "", "type": "020lead"},
        {"cluster_id": "102", "plain_text": "Old opinion.", "html": "", "type": "020lead"},
        {"cluster_id": "104", "plain_text": "Unpublished opinion.", "html": "", "type": "020lead"},
    ])

    out = list(cl_bulk_loader.iter_recent_ca_cases(
        citations_stream=citations, clusters_stream=clusters, opinions_stream=opinions,
        cutoff_date="2017-01-01", published_only=True,
    ))
    uids = {c.case_uid for c, _ in out}
    assert uids == {"cl:100", "cl:101"}   # CA + published + post-2017 only
    by_uid = {c.case_uid: c for c, _ in out}
    rec = by_uid["cl:100"]
    assert rec.source == "cl"
    assert rec.citation == "15 Cal. 5th 1"
    assert rec.year == "2023"
    assert rec.citation_count == 7
    assert "arbitration" in rec.full_text
    # passages present (keyword-indexable)
    case, passages = next(o for o in out if o[0].case_uid == "cl:100")
    assert passages and passages[0].case_uid == "cl:100"


def test_parses_courtlistener_backslash_escaped_opinion_text():
    citations = _csv([
        {"id": "1", "volume": "15", "reporter": "Cal. 5th", "page": "1", "type": "8", "cluster_id": "100"},
    ])
    clusters = _csv([
        {"id": "100", "date_filed": "2023-06-01", "case_name": "Recent Pub v. CA",
         "case_name_short": "Recent", "citation_count": "7", "precedential_status": "Published", "docket_id": "9"},
    ])
    opinions = io.StringIO(
        "cluster_id,plain_text,html,type\n"
        '100,"The court said \\"hello, world\\" before ruling.",,020lead\n'
    )

    out = list(cl_bulk_loader.iter_recent_ca_cases(
        citations_stream=citations, clusters_stream=clusters, opinions_stream=opinions,
        cutoff_date="2017-01-01", published_only=True,
    ))

    assert len(out) == 1
    case, passages = out[0]
    assert case.case_uid == "cl:100"
    assert '"hello, world"' in case.full_text
    assert '"hello, world"' in passages[0].text
