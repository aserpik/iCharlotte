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


def test_streams_only_ca_recent_cases():
    courts = _csv([
        {"id": "cal", "full_name": "Supreme Court of California"},
        {"id": "ny", "full_name": "New York Court of Appeals"},
    ])
    clusters = _csv([
        {"id": "100", "court_id": "cal", "case_name": "Recent v. CA",
         "date_filed": "2023-06-01", "citation": "15 Cal. 5th 1"},
        {"id": "101", "court_id": "cal", "case_name": "Old v. CA",
         "date_filed": "1990-01-01", "citation": "50 Cal. 3d 1"},  # before cutoff
        {"id": "102", "court_id": "ny", "case_name": "NY thing",
         "date_filed": "2024-01-01", "citation": "1 N.Y.3d 1"},     # wrong court
    ])
    opinions = _csv([
        {"cluster_id": "100", "plain_text": "Recent CA opinion about privacy.", "html": ""},
        {"cluster_id": "101", "plain_text": "Old opinion.", "html": ""},
        {"cluster_id": "102", "plain_text": "NY opinion.", "html": ""},
    ])

    out = list(cl_bulk_loader.iter_recent_ca_cases(
        courts_stream=courts, clusters_stream=clusters, opinions_stream=opinions,
        cutoff_date="2020-01-01",
    ))
    assert len(out) == 1
    case, passages = out[0]
    assert case.case_uid == "cl:100"
    assert case.source == "cl"
    assert case.citation == "15 Cal. 5th 1"
    assert case.year == "2023"
    assert "privacy" in case.full_text
    assert passages and passages[0].case_uid == "cl:100"
