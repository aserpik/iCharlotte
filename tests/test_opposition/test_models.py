from icharlotte_core.opposition import (
    CitationVerification,
    DraftDocument,
    MotionMetadata,
    OutlineNode,
    SectionPlanItem,
)


def test_motion_metadata_round_trip_and_required_missing():
    metadata = MotionMetadata(
        motion_type="Motion for Summary Judgment",
        moving_party="Defendant",
        opposing_party="Plaintiff",
        relief_requested="Judgment in defendant's favor",
        hearing_date="2026-07-10",
        opposition_due_date="2026-06-26",
        procedural_posture="Discovery complete",
        principal_arguments=["No triable issue", "No damages evidence"],
        opposition_posture="Oppose on factual disputes",
    )

    restored = MotionMetadata.from_dict(metadata.to_dict())

    assert restored == metadata
    assert restored.required_missing() == []
    assert MotionMetadata().required_missing() == [
        "Motion Type",
        "Relief Requested",
        "Principal Arguments",
    ]


def test_outline_node_round_trip_preserves_nested_selected_states():
    outline = OutlineNode(
        id="root",
        text="Argument",
        selected=False,
        children=[
            OutlineNode(id="a", text="Facts", selected=True),
            OutlineNode(
                id="b",
                text="Law",
                selected=False,
                children=[OutlineNode(id="b1", text="Standard", selected=True)],
            ),
        ],
    )

    restored = OutlineNode.from_dict(outline.to_dict())

    assert restored == outline
    assert restored.selected is False
    assert restored.children[0].selected is True
    assert restored.children[1].selected is False
    assert restored.children[1].children[0].selected is True


def test_draft_document_round_trip_preserves_citation_records():
    citation = CitationVerification(
        citation_text="Smith v. Jones",
        normalized_citation="Smith v. Jones, 1 Cal.5th 1",
        status="verified",
        case_name="Smith v. Jones",
        replacement_candidates=["Smith v. Brown"],
    )
    draft = DraftDocument(
        title="Opposition",
        body_text="Argument text",
        citations=[citation],
        preview_path="C:/tmp/opposition.docx",
    )

    restored = DraftDocument.from_dict(draft.to_dict())

    assert restored == draft
    assert restored.citations[0].citation_text == "Smith v. Jones"
    assert restored.citations[0].replacement_candidates == ["Smith v. Brown"]


def test_public_list_fields_normalize_string_values():
    metadata = MotionMetadata.from_dict({"principal_arguments": "No triable issue"})
    section = SectionPlanItem.from_dict({"path": "Argument"})
    citation = CitationVerification.from_dict(
        {"replacement_candidates": "Better Authority"}
    )

    assert metadata.principal_arguments == ["No triable issue"]
    assert section.path == ["Argument"]
    assert citation.replacement_candidates == ["Better Authority"]


def test_citation_verification_defaults_for_new_verdict_fields():
    cv = CitationVerification()
    assert cv.verdict == ""
    assert cv.kind == ""
    assert cv.proposition == ""
    assert cv.evidence == ""
    assert cv.note == ""
    assert cv.law_code == ""
    assert cv.section_num == ""
    assert cv.body_offset is None


def test_citation_verification_from_dict_roundtrip_with_verdict_fields():
    data = {
        "citation_text": "Cottini v. Enloe Medical Center (2014) 226 Cal.App.4th 401",
        "verdict": "SUPPORTED",
        "kind": "case",
        "proposition": "Trial courts retain discretion to deny untimely motions.",
        "evidence": "The trial court did not abuse its discretion.",
        "note": "Direct support; matches the brief's framing.",
        "body_offset": 1284,
    }
    cv = CitationVerification.from_dict(data)
    assert cv.verdict == "SUPPORTED"
    assert cv.kind == "case"
    assert cv.body_offset == 1284
    assert cv.proposition.startswith("Trial courts")

    out = cv.to_dict()
    for key, expected in data.items():
        assert out[key] == expected


def test_citation_verification_from_dict_statute_kind():
    data = {
        "citation_text": "Code Civ. Proc., § 2024.020",
        "kind": "statute",
        "law_code": "CCP",
        "section_num": "2024.020",
        "verdict": "PARTIAL",
    }
    cv = CitationVerification.from_dict(data)
    assert cv.kind == "statute"
    assert cv.law_code == "CCP"
    assert cv.section_num == "2024.020"
    assert cv.verdict == "PARTIAL"


def test_retrieved_authority_roundtrip():
    from icharlotte_core.opposition.models import RetrievedAuthority

    ra = RetrievedAuthority(
        argument_id="arg-1",
        argument_text="Discovery cutoff bars the motion",
        cluster_id="12345",
        case_name="Cottini v. Enloe Medical Center",
        citation="226 Cal.App.4th 401",
        supports="A trial court retains discretion over late discovery motions.",
        passage="The trial court did not abuse its discretion.",
        opinion_url="https://www.courtlistener.com/opinion/12345/cottini/",
        citation_count=37,
        latest_citing_year="2021",
    )
    data = ra.to_dict()
    restored = RetrievedAuthority.from_dict(data)
    assert restored == ra
    assert RetrievedAuthority.from_dict({}).cluster_id == ""
