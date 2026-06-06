from icharlotte_core.legal_research.deep_research import (
    CitationAuditStatus,
    ResearchPacket,
    SelectedAuthority,
)
from icharlotte_core.legal_research.deep_research.verification import (
    audit_citations_against_packet,
    contains_verbatim_quote,
)


def test_contains_verbatim_quote_normalizes_whitespace():
    source_text = "The court must weigh\n the likely benefit\t against the burden."
    quote = "court must weigh the likely benefit against"

    assert contains_verbatim_quote(source_text, quote) is True


def test_contains_verbatim_quote_normalizes_case():
    source_text = "The Court Must Weigh the Likely Benefit Against the Burden."
    quote = "court must weigh the likely benefit against"

    assert contains_verbatim_quote(source_text, quote) is True


def test_contains_verbatim_quote_rejects_changed_words():
    source_text = "The court must weigh the likely benefit against the burden."
    quote = "court must weigh the possible benefit against"

    assert contains_verbatim_quote(source_text, quote) is False


def test_contains_verbatim_quote_rejects_empty_inputs():
    assert contains_verbatim_quote("", "quoted words") is False
    assert contains_verbatim_quote("source text", "") is False


def test_audit_citations_passes_known_packet_case():
    packet = ResearchPacket(
        selected_authorities=[
            SelectedAuthority(
                case_name="Smith v. Jones",
                citation="12 Cal.5th 100",
                year="2024",
                verification_status="verified",
            )
        ]
    )

    audit = audit_citations_against_packet(
        "Smith v. Jones (2024) 12 Cal.5th 100 controls the issue.",
        packet,
    )

    assert len(audit.items) == 1
    assert audit.items[0].citation_text == "Smith v. Jones (2024) 12 Cal.5th 100"
    assert audit.items[0].status == CitationAuditStatus.SUPPORTED
    assert audit.has_off_packet_citations is False


def test_audit_citations_matches_reporter_only_packet_case():
    packet = ResearchPacket(
        selected_authorities=[
            SelectedAuthority(
                case_name="Smith v. Jones",
                citation="12 Cal.5th 100",
                year="2024",
                verification_status="verified",
            )
        ]
    )

    audit = audit_citations_against_packet(
        "Smith v. Jones (2024) 12 Cal.5th 100 controls the issue.",
        packet,
    )

    assert audit.items[0].status == CitationAuditStatus.SUPPORTED


def test_audit_citations_flags_off_packet_case():
    packet = ResearchPacket()

    audit = audit_citations_against_packet(
        "Smith v. Jones (2024) 12 Cal.5th 100 controls the issue.",
        packet,
    )

    assert len(audit.items) == 1
    assert audit.items[0].citation_text == "Smith v. Jones (2024) 12 Cal.5th 100"
    assert audit.items[0].status == CitationAuditStatus.OFF_PACKET
    assert audit.has_off_packet_citations is True


def test_audit_citations_excludes_unverified_authorities():
    packet = ResearchPacket(
        selected_authorities=[
            SelectedAuthority(
                case_name="Smith v. Jones",
                citation="12 Cal.5th 100",
                year="2024",
                verification_status="unverified_firm",
            ),
            SelectedAuthority(
                case_name="Pending v. Review",
                citation="15 Cal.App.5th 200",
                year="2023",
                verification_status="pending",
            ),
        ]
    )

    audit = audit_citations_against_packet(
        "Smith v. Jones (2024) 12 Cal.5th 100 and "
        "Pending v. Review (2023) 15 Cal.App.5th 200 are cited.",
        packet,
    )

    assert [item.status for item in audit.items] == [
        CitationAuditStatus.OFF_PACKET,
        CitationAuditStatus.OFF_PACKET,
    ]


def test_audit_citations_treats_none_status_as_off_packet():
    packet = ResearchPacket(
        selected_authorities=[
            SelectedAuthority(
                case_name="Smith v. Jones",
                citation="12 Cal.5th 100",
                year="2024",
                verification_status=None,
            )
        ]
    )

    audit = audit_citations_against_packet(
        "Smith v. Jones (2024) 12 Cal.5th 100 controls the issue.",
        packet,
    )

    assert len(audit.items) == 1
    assert audit.items[0].status == CitationAuditStatus.OFF_PACKET
