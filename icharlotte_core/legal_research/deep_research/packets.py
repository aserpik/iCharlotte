"""Prompt and display formatting for deep research packets."""
from __future__ import annotations

from .models import ResearchPacket, SelectedAuthority


def _normalized_verification_status(authority: SelectedAuthority) -> str:
    return (authority.verification_status or "").strip().lower()


def _verified_authorities(packet: ResearchPacket) -> list[SelectedAuthority]:
    return [
        authority
        for authority in packet.selected_authorities
        if _normalized_verification_status(authority) == "verified"
    ]


def _excluded_authorities(packet: ResearchPacket) -> list[SelectedAuthority]:
    return [
        authority
        for authority in packet.selected_authorities
        if _normalized_verification_status(authority) != "verified"
    ]


def packet_to_prompt_block(packet: ResearchPacket) -> str:
    lines = [
        "[DEEP RESEARCH AUTHORITY]",
        "Use only the verified authorities in this block for legal citations.",
        "Do not cite cases from memory or from unverified research notes.",
        "",
    ]
    authorities = _verified_authorities(packet)
    if authorities:
        lines.append("Verified case law:")
        for authority in authorities:
            lines.append(f"- {authority.formatted_citation}")
            if authority.supports:
                lines.append(f"  Supports: {authority.supports}")
            if authority.verbatim_quote:
                lines.append(f"  Quote: {authority.verbatim_quote}")
            if authority.parenthetical_summary:
                lines.append(
                    "  Parenthetical research note: "
                    f"{authority.parenthetical_summary} "
                    f"(source: {authority.parenthetical_source or 'CourtListener bulk'}; "
                    "not a quote from the cited opinion)."
                )
            if authority.limitations:
                lines.append(f"  Limitation: {authority.limitations}")
    else:
        lines.append("Verified case law: none.")
    if packet.statutory_materials:
        lines.append("")
        lines.append("Statutory material:")
        for statute in packet.statutory_materials:
            lines.append(f"- {statute.code} section {statute.section}: {statute.title}")
            if statute.text:
                lines.append(f"  Text: {statute.text}")
    if packet.warnings:
        lines.append("")
        lines.append("Warnings:")
        for warning in packet.warnings:
            lines.append(f"- {warning}")
    lines.append("")
    lines.append("[/DEEP RESEARCH AUTHORITY]")
    return "\n".join(lines)


def packet_to_research_basis_markdown(packet: ResearchPacket) -> str:
    lines = ["### Research Basis"]
    if packet.searches_run:
        lines.append("")
        lines.append("Searches run:")
        for search in packet.searches_run:
            lines.append(f"- {search}")
    verified = _verified_authorities(packet)
    if verified:
        lines.append("")
        lines.append("Authorities selected:")
        for authority in verified:
            reason = authority.selection_reason or authority.supports
            lines.append(f"- {authority.formatted_citation}: {reason}")
    excluded = _excluded_authorities(packet)
    if excluded:
        lines.append("")
        lines.append("Excluded or unverified authorities (not prompt-safe):")
        for authority in excluded:
            status = _normalized_verification_status(authority) or "unknown"
            reason = authority.selection_reason or authority.supports or "No prompt-safe verification."
            lines.append(
                f"- {authority.formatted_citation}: "
                f"verification_status={status}; {reason}"
            )
    if packet.warnings:
        lines.append("")
        lines.append("Warnings:")
        for warning in packet.warnings:
            lines.append(f"- {warning}")
    return "\n".join(lines)
