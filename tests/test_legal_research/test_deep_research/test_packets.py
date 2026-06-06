from icharlotte_core.legal_research.deep_research import (
    ResearchPacket,
    SelectedAuthority,
    StatutoryMaterial,
)
from icharlotte_core.legal_research.deep_research.packets import (
    packet_to_prompt_block,
    packet_to_research_basis_markdown,
)


def _packet():
    return ResearchPacket(
        selected_authorities=[
            SelectedAuthority(
                case_name="Smith v. Jones",
                citation="12 Cal.5th 100",
                year="2024",
                court="Cal.",
                supports="Trial courts must consider proportionality.",
                verbatim_quote="The court must weigh the likely benefit against the burden.",
                selection_reason="Recent California Supreme Court case with direct opinion text.",
                parenthetical_summary="holding that discovery burden matters",
                parenthetical_source="Later v. Case (2025) 15 Cal.5th 200",
                verification_status="verified",
            ),
            SelectedAuthority(
                case_name="Unverified Firm Case",
                citation="99 Cal.App.5th 1",
                year="2023",
                source="firm",
                supports="Firm-only cite.",
                verbatim_quote="",
                verification_status="unverified_firm",
            ),
            SelectedAuthority(
                case_name="Pending Review Case",
                citation="88 Cal.App.5th 2",
                year="2022",
                supports="This authority has not completed verification.",
                verification_status="pending",
            ),
        ],
        statutory_materials=[
            StatutoryMaterial(
                code="Code Civ. Proc.",
                section="2017.020",
                title="Discovery limitation on burden or expense",
                text="The court shall limit discovery if the burden outweighs the likely benefit.",
            )
        ],
        searches_run=["local semantic: proportional discovery burden"],
        warnings=["CourtListener was not called."],
    )


def test_prompt_block_includes_verified_authority_quote():
    block = packet_to_prompt_block(_packet())

    assert "[DEEP RESEARCH AUTHORITY]" in block
    assert "Smith v. Jones (2024) 12 Cal.5th 100" in block
    assert "Quote: The court must weigh the likely benefit against the burden." in block


def test_prompt_block_labels_parentheticals_as_research_notes():
    block = packet_to_prompt_block(_packet())

    assert "Parenthetical research note:" in block
    assert "not a quote from the cited opinion" in block


def test_prompt_block_excludes_unverified_firm_authority_from_verified_section():
    block = packet_to_prompt_block(_packet())

    assert "Unverified Firm Case" not in block


def test_prompt_block_excludes_non_firm_non_verified_statuses():
    block = packet_to_prompt_block(_packet())

    assert "Pending Review Case" not in block


def test_prompt_block_includes_statutory_material():
    block = packet_to_prompt_block(_packet())

    assert "Statutory material:" in block
    assert "Code Civ. Proc. section 2017.020: Discovery limitation on burden or expense" in block
    assert "Text: The court shall limit discovery if the burden outweighs the likely benefit." in block


def test_research_basis_mentions_searches_and_warnings():
    basis = packet_to_research_basis_markdown(_packet())

    assert "Research Basis" in basis
    assert "local semantic: proportional discovery burden" in basis
    assert "CourtListener was not called." in basis


def test_research_basis_lists_excluded_authorities_under_non_prompt_safe_label():
    basis = packet_to_research_basis_markdown(_packet())

    assert "Excluded or unverified authorities (not prompt-safe):" in basis
    assert "Unverified Firm Case (2023) 99 Cal.App.5th 1" in basis
    assert "verification_status=unverified_firm" in basis
    assert "Pending Review Case (2022) 88 Cal.App.5th 2" in basis
    assert "verification_status=pending" in basis


def test_packet_instance_methods_delegate_to_formatters():
    packet = _packet()

    assert packet.to_prompt_block() == packet_to_prompt_block(packet)
    assert packet.to_research_basis_markdown() == packet_to_research_basis_markdown(packet)
