"""Schema validation tests for depo_prep_lib.schemas."""
import pytest

from Scripts.depo_prep_lib.schemas import (
    DeponentStatement,
    FactualAnchor,
    Inconsistency,
    SourceDigest,
    Topic,
    Question,
    TopicQuestions,
    validate_source_digest_dict,
    validate_topics_dict,
)


def test_source_digest_roundtrip():
    digest = SourceDigest(
        source_id="med_records.pdf",
        source_kind="medical_records",
        deponent_statements=[
            DeponentStatement(text="I had no prior pain.", location="p.47:18-22", context="Direct exam")
        ],
        factual_anchors=[
            FactualAnchor(fact="MRI 2024-09-12 showed 4mm protrusion", location="p.12 Bates DEF-00154",
                          topic_tags=["injury", "imaging"])
        ],
        inconsistencies=[
            Inconsistency(claim_a="RFA #7: pain immediate", claim_a_source="this file, RFA #7",
                          claim_b="ER triage: no acute pain", claim_b_source="med_records p.3",
                          topic_tags=["credibility"])
        ],
        summary="Med records show chronic LBP prior to accident.",
    )
    d = digest.to_dict()
    assert d["source_id"] == "med_records.pdf"
    assert d["deponent_statements"][0]["location"] == "p.47:18-22"
    # Round-trip
    digest2 = SourceDigest.from_dict(d)
    assert digest2.summary == digest.summary


def test_validate_source_digest_rejects_missing_field():
    bad = {"source_id": "x", "source_kind": "other"}  # missing required lists
    with pytest.raises(ValueError, match="missing"):
        validate_source_digest_dict(bad)


def test_topic_default_checked_true():
    t = Topic(id="t01", title="Pain timeline", strategic_note="Establish baseline",
              relevant_digest_refs=["a.pdf#factual_anchors[0]"])
    assert t.default_checked is True
    assert t.lawyer_added is False


def test_question_optional_fields_default_none():
    q = Question(n=1, text="When did pain begin?")
    assert q.purpose is None
    assert q.source_facts is None
    assert q.impeachment_hook is None
    assert q.objection_alts is None


def test_topic_questions_roundtrip():
    tq = TopicQuestions(
        topic_id="t01",
        questions=[Question(n=1, text="Q1", purpose="P1")],
    )
    d = tq.to_dict()
    assert d["questions"][0]["purpose"] == "P1"
    tq2 = TopicQuestions.from_dict(d)
    assert tq2.questions[0].text == "Q1"


def test_validate_topics_dict_accepts_minimal_topic():
    payload = {"topics": [{"id": "t01", "title": "X", "strategic_note": "Y",
                            "relevant_digest_refs": []}]}
    validate_topics_dict(payload)  # should not raise


def test_validate_topics_dict_rejects_non_list_topics():
    with pytest.raises(ValueError):
        validate_topics_dict({"topics": "nope"})
