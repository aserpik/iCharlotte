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
