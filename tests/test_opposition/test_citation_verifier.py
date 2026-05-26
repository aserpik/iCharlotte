from unittest.mock import MagicMock, patch

from icharlotte_core.legal_research.sources.courtlistener import (
    BASE_URL,
    CourtListenerClient,
    REQUEST_TIMEOUT,
    opinion_url_for_cluster,
)
from icharlotte_core.opposition.citation_verifier import (
    find_supporting_passage,
    map_lookup_record,
    verify_citations,
)


def test_lookup_citations_blank_text_does_not_call_network():
    client = CourtListenerClient(token="tok")

    with patch(
        "icharlotte_core.legal_research.sources.courtlistener.requests.post"
    ) as mock_post:
        assert client.lookup_citations("   ") == []

    mock_post.assert_not_called()


@patch("icharlotte_core.legal_research.sources.courtlistener.requests.post")
def test_lookup_citations_returns_raw_records(mock_post):
    response = MagicMock()
    response.json.return_value = [
        {
            "citation": "Smith v. Jones, 1 Cal.5th 1",
            "clusters": [{"id": 123, "case_name": "Smith v. Jones"}],
        }
    ]
    response.raise_for_status = MagicMock()
    mock_post.return_value = response

    client = CourtListenerClient(token="tok")
    records = client.lookup_citations("Smith v. Jones, 1 Cal.5th 1")

    assert records == [
        {
            "citation": "Smith v. Jones, 1 Cal.5th 1",
            "clusters": [{"id": 123, "case_name": "Smith v. Jones"}],
        }
    ]
    mock_post.assert_called_once()
    assert mock_post.call_args.kwargs["headers"]["Authorization"] == "Token tok"
    assert (
        mock_post.call_args.kwargs["headers"].get("Content-Type")
        != "application/json"
    )
    assert mock_post.call_args.kwargs["data"] == {
        "text": "Smith v. Jones, 1 Cal.5th 1"
    }
    assert mock_post.call_args.kwargs["timeout"] == REQUEST_TIMEOUT
    response.raise_for_status.assert_called_once()


@patch("icharlotte_core.legal_research.sources.courtlistener.requests.post")
def test_lookup_citations_handles_api_exception(mock_post):
    mock_post.side_effect = Exception("network down")

    client = CourtListenerClient(token="tok")

    assert client.lookup_citations("Smith v. Jones") == []


@patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
def test_get_cluster_returns_json(mock_get):
    response = MagicMock()
    response.json.return_value = {"id": 123, "absolute_url": "/opinion/123/smith/"}
    response.raise_for_status = MagicMock()
    mock_get.return_value = response

    client = CourtListenerClient(token="tok")
    cluster = client.get_cluster("123")

    assert cluster == {"id": 123, "absolute_url": "/opinion/123/smith/"}
    mock_get.assert_called_once_with(
        f"{BASE_URL}/clusters/123/",
        headers=client._headers(),
        timeout=REQUEST_TIMEOUT,
    )


@patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
def test_get_cluster_handles_api_exception(mock_get):
    mock_get.side_effect = Exception("network down")

    client = CourtListenerClient(token="tok")

    assert client.get_cluster(123) is None


@patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
def test_get_cluster_returns_none_for_non_dict_json(mock_get):
    response = MagicMock()
    response.json.return_value = [{"id": 123}]
    response.raise_for_status = MagicMock()
    mock_get.return_value = response

    client = CourtListenerClient(token="tok")

    assert client.get_cluster(123) is None
    response.raise_for_status.assert_called_once()


def test_opinion_url_for_cluster_prefixes_relative_absolute_url():
    assert (
        opinion_url_for_cluster({"absolute_url": "/opinion/123/smith/"})
        == "https://www.courtlistener.com/opinion/123/smith/"
    )
    assert (
        opinion_url_for_cluster(
            {"absolute_url": "https://www.courtlistener.com/opinion/123/smith/"}
        )
        == "https://www.courtlistener.com/opinion/123/smith/"
    )
    assert opinion_url_for_cluster({}) == ""


def test_opinion_url_for_cluster_joins_path_without_leading_slash():
    assert (
        opinion_url_for_cluster({"absolute_url": "opinion/123/smith/"})
        == "https://www.courtlistener.com/opinion/123/smith/"
    )


def test_map_lookup_record_found_record():
    record = {
        "citation": "Smith v. Jones, 1 Cal.5th 1",
        "normalized_citations": ["1 Cal.5th 1"],
        "status": 200,
        "caseName": "Smith v. Jones",
        "court": "Cal.",
        "dateFiled": "2020-01-02",
        "cluster_id": 123,
        "absolute_url": "/opinion/123/smith/",
        "warning": "parallel citation omitted",
    }

    verification = map_lookup_record(record)

    assert verification.citation_text == "Smith v. Jones, 1 Cal.5th 1"
    assert verification.normalized_citation == "1 Cal.5th 1"
    assert verification.status == "exists_support_unconfirmed"
    assert verification.case_name == "Smith v. Jones"
    assert verification.court == "Cal."
    assert verification.date == "2020-01-02"
    assert verification.cluster_id == "123"
    assert (
        verification.opinion_url
        == "https://www.courtlistener.com/opinion/123/smith/"
    )
    assert verification.warning == "parallel citation omitted"


def test_map_lookup_record_prefers_first_cluster_fields():
    record = {
        "citation": "Smith v. Jones, 1 Cal.5th 1",
        "normalized_citations": ["1 Cal.5th 1"],
        "status": 200,
        "id": 999,
        "case_name": "Wrong Top-Level Case",
        "court": "Wrong Court",
        "date_filed": "1999-01-01",
        "clusters": [
            {
                "id": 123,
                "case_name": "Smith v. Jones",
                "court": "Cal.",
                "date_filed": "2020-01-02",
                "absolute_url": "/opinion/123/smith/",
            }
        ],
    }

    verification = map_lookup_record(record)

    assert verification.cluster_id == "123"
    assert verification.case_name == "Smith v. Jones"
    assert verification.court == "Cal."
    assert verification.date == "2020-01-02"


def test_map_lookup_record_not_found_record():
    record = {
        "citation": "Missing v. Case, 999 Cal.9th 9",
        "normalized_citations": ["999 Cal.9th 9"],
        "status": 404,
        "error": "not found",
    }

    verification = map_lookup_record(record)

    assert verification.citation_text == "Missing v. Case, 999 Cal.9th 9"
    assert verification.normalized_citation == "999 Cal.9th 9"
    assert verification.status == "not_found"
    assert verification.warning == "not found"


def test_find_supporting_passage_returns_overlap_match():
    text = (
        "A landowner generally owes a duty of care to maintain property. "
        "The statute concerns venue only."
    )

    passage = find_supporting_passage(
        "landowner owes duty to maintain property with reasonable care",
        text,
    )

    assert passage == "A landowner generally owes a duty of care to maintain property."


def test_find_supporting_passage_returns_empty_for_no_match():
    text = "The case discusses arbitration clauses and forum selection."

    assert find_supporting_passage("premises liability duty of care", text) == ""


def test_verify_citations_marks_support_passage_found_without_semantic_checker():
    class FakeCourtListener:
        def lookup_citations(self, draft_text):
            assert "Smith v. Jones" in draft_text
            return [
                {
                    "citation": "Smith v. Jones, 1 Cal.5th 1",
                    "normalized_citations": ["1 Cal.5th 1"],
                    "status": 200,
                    "cluster_id": 123,
                    "absolute_url": "/opinion/123/smith/",
                },
                {
                    "citation": "Missing v. Case, 999 Cal.9th 9",
                    "normalized_citations": ["999 Cal.9th 9"],
                    "status": 404,
                },
            ]

        def get_opinion_text(self, cluster_id):
            assert cluster_id == "123"
            return "Landowners owe a duty to maintain property in safe condition."

    results = verify_citations(
        "Smith v. Jones, 1 Cal.5th 1; Missing v. Case, 999 Cal.9th 9",
        citation_propositions={
            "1 Cal.5th 1": "landowners owe duty to maintain property safely",
        },
        courtlistener=FakeCourtListener(),
    )

    assert [result.status for result in results] == [
        "support_passage_found",
        "not_found",
    ]
    assert (
        results[0].supporting_passage
        == "Landowners owe a duty to maintain property in safe condition."
    )
    assert results[1].supporting_passage == ""


def test_verify_citations_marks_verified_when_support_checker_confirms():
    class FakeCourtListener:
        def lookup_citations(self, draft_text):
            return [
                {
                    "citation": "Smith v. Jones, 1 Cal.5th 1",
                    "normalized_citations": ["1 Cal.5th 1"],
                    "status": 200,
                    "clusters": [{"id": 123}],
                },
            ]

        def get_opinion_text(self, cluster_id):
            return "Landowners owe a duty to maintain property in safe condition."

    results = verify_citations(
        "Smith v. Jones, 1 Cal.5th 1",
        citation_propositions={
            "1 Cal.5th 1": "landowners owe duty to maintain property safely",
        },
        courtlistener=FakeCourtListener(),
        support_checker=lambda verification, proposition, passage: True,
    )

    assert results[0].status == "verified"


def test_verify_citations_caches_repeated_cluster_opinion_text():
    class FakeCourtListener:
        def __init__(self):
            self.calls = 0

        def lookup_citations(self, draft_text):
            return [
                {
                    "citation": "Smith v. Jones, 1 Cal.5th 1",
                    "normalized_citations": ["1 Cal.5th 1"],
                    "status": 200,
                    "clusters": [{"id": 123}],
                },
                {
                    "citation": "Smith v. Jones, 1 Cal.5th 2",
                    "normalized_citations": ["1 Cal.5th 2"],
                    "status": 200,
                    "clusters": [{"id": 123}],
                },
            ]

        def get_opinion_text(self, cluster_id):
            self.calls += 1
            return "Landowners owe a duty to maintain property in safe condition."

    client = FakeCourtListener()
    results = verify_citations(
        "Smith v. Jones, 1 Cal.5th 1; Smith v. Jones, 1 Cal.5th 2",
        citation_propositions={
            "1 Cal.5th 1": "landowners owe duty to maintain property safely",
            "1 Cal.5th 2": "landowners owe duty to maintain property safely",
        },
        courtlistener=client,
    )

    assert [result.status for result in results] == [
        "support_passage_found",
        "support_passage_found",
    ]
    assert client.calls == 1


def test_verify_citations_sets_offsets_for_whitespace_compacted_passage():
    class FakeCourtListener:
        def lookup_citations(self, draft_text):
            return [
                {
                    "citation": "Smith v. Jones, 1 Cal.5th 1",
                    "normalized_citations": ["1 Cal.5th 1"],
                    "status": 200,
                    "clusters": [{"id": 123}],
                },
            ]

        def get_opinion_text(self, cluster_id):
            return (
                "Intro.\n"
                "Landowners owe a duty\n"
                "to maintain property in safe condition.\n"
                "End."
            )

    opinion_text = (
        "Intro.\n"
        "Landowners owe a duty\n"
        "to maintain property in safe condition.\n"
        "End."
    )
    results = verify_citations(
        "Smith v. Jones, 1 Cal.5th 1",
        citation_propositions={
            "1 Cal.5th 1": "landowners owe duty to maintain property safely",
        },
        courtlistener=FakeCourtListener(),
    )

    assert results[0].status == "support_passage_found"
    assert results[0].support_start is not None
    assert results[0].support_end is not None
    assert "Landowners owe" in opinion_text[
        results[0].support_start : results[0].support_end
    ]


def test_verify_citations_preserves_unconfirmed_when_support_missing():
    class FakeCourtListener:
        def lookup_citations(self, draft_text):
            return [
                {
                    "citation": "Smith v. Jones, 1 Cal.5th 1",
                    "normalized_citations": ["1 Cal.5th 1"],
                    "status": 200,
                    "id": 123,
                },
            ]

        def get_opinion_text(self, cluster_id):
            assert cluster_id == "123"
            return "This opinion discusses arbitration and venue."

    results = verify_citations(
        "Smith v. Jones, 1 Cal.5th 1",
        citation_propositions={
            "Smith v. Jones, 1 Cal.5th 1": "landowners owe a duty of care",
        },
        courtlistener=FakeCourtListener(),
    )

    assert len(results) == 1
    assert results[0].status == "exists_support_unconfirmed"
    assert results[0].supporting_passage == ""


def test_replacement_candidates_are_added_for_unconfirmed_support():
    class ReplacementClient:
        def lookup_citations(self, text):
            return [
                {
                    "citation": "69 Cal.2d 108",
                    "normalized_citations": ["69 Cal.2d 108"],
                    "status": 200,
                    "clusters": [{"id": 10, "case_name": "Rowland v. Christian"}],
                }
            ]

        def get_opinion_text(self, cluster_id):
            return "No matching support terms here."

        def search_opinions(self, query, max_results=5):
            from icharlotte_core.legal_research.models import CaseResult

            assert max_results == 5
            return [
                CaseResult(
                    name="Aguilar v. Atlantic Richfield Co.",
                    citation="25 Cal.4th 826",
                    date="2001-07-12",
                    court="cal",
                    snippet="summary judgment burden and triable issue rule",
                    url="https://www.courtlistener.com/opinion/aguilar/",
                    cluster_id=99,
                )
            ]

    result = verify_citations(
        "Text (Rowland v. Christian (1968) 69 Cal.2d 108.)",
        citation_propositions={"69 Cal.2d 108": "summary judgment triable issue burden"},
        courtlistener=ReplacementClient(),
    )

    assert result[0].status == "exists_support_unconfirmed"
    assert (
        result[0].replacement_candidates[0]["case_name"]
        == "Aguilar v. Atlantic Richfield Co."
    )
