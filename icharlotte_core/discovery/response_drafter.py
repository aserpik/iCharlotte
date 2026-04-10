"""
Phase 3: Response Drafter.

Assembles the complete plain-text response document by combining parsed
requests, selected objections, and substantive responses.  Also provides
LLM prompt builders so the UI layer can draft substantive responses for
each discovery type.
"""
from __future__ import annotations

from typing import Dict, List, Optional

from icharlotte_core.discovery.response_parser import ParsedDiscovery
from icharlotte_core.discovery.response_rules import ResponseRules

# ---------------------------------------------------------------------------
# Header templates
# ---------------------------------------------------------------------------

_HEADERS: Dict[str, tuple] = {
    "FI": (
        "FORM INTERROGATORY NO. {n}:",
        "RESPONSE TO FORM INTERROGATORY NO. {n}:",
    ),
    "SI": (
        "SPECIAL INTERROGATORY NO. {n}:",
        "RESPONSE TO SPECIAL INTERROGATORY NO. {n}:",
    ),
    "RFA": (
        "REQUEST FOR ADMISSION NO. {n}:",
        "RESPONSE TO REQUEST FOR ADMISSION NO. {n}:",
    ),
    "RPD": (
        "REQUEST FOR PRODUCTION NO. {n}:",
        "RESPONSE TO REQUEST NO. {n}:",
    ),
}


# ---------------------------------------------------------------------------
# format_single_response
# ---------------------------------------------------------------------------

def format_single_response(
    disc_type: str,
    request_number: str,
    request_text: str,
    objections: str,
    substantive: str,
    waiver: str,
    reservation: str,
) -> str:
    """
    Format a single request/response block as plain text.

    The layout is:

        REQUEST HEADER
        <request_text>

        RESPONSE HEADER
        <objections>
        <waiver>
        <substantive>
        <reservation>
    """
    req_tmpl, resp_tmpl = _HEADERS.get(disc_type, ("{n}:", "RESPONSE {n}:"))
    req_header = req_tmpl.format(n=request_number)
    resp_header = resp_tmpl.format(n=request_number)

    parts = [req_header, request_text, "", resp_header]

    if objections:
        parts.append(objections)
    if waiver:
        parts.append(waiver)
    if substantive:
        parts.append(substantive)
    if reservation:
        parts.append(reservation)

    return "\n".join(parts)


# ---------------------------------------------------------------------------
# FI series helpers
# ---------------------------------------------------------------------------

def is_fi_series(number_str: str, series_list: List[int]) -> bool:
    """
    Return True if the integer prefix of *number_str* (e.g. "16" from "16.1")
    is contained in *series_list*.
    """
    try:
        prefix = int(number_str.split(".")[0])
    except (ValueError, IndexError):
        return False
    return prefix in series_list


def get_fi_fixed_response(number_str: str, rules: ResponseRules) -> Optional[str]:
    """
    Return the fixed (pre-written) response text for certain FI numbers,
    or None if the question needs an LLM-drafted response.

    Fixed FIs:
      - 1.1  → fi_1_1_response
      - 15.1 → fi_15_1_response
      - 16.x → fi_16_response (section 16.0 series)
      - 17.1 → placeholder "[PENDING …]"
    """
    n = number_str.strip()

    if n == "1.1":
        return rules.fi_1_1_response

    if n == "15.1":
        return rules.fi_15_1_response

    if is_fi_series(n, [16]):
        return rules.fi_16_response

    if n == "17.1":
        return "[PENDING — complete after RFA responses are finalized]"

    return None


# ---------------------------------------------------------------------------
# detect_inapplicable_fi
# ---------------------------------------------------------------------------

# FI series that are applicable only to specific case types
_WRONGFUL_DEATH_SERIES = [10]          # FI 10.x series
_EMPLOYMENT_SERIES = list(range(200, 300))  # 200+ series


def detect_inapplicable_fi(fi_number: str, case_type: str = "negligence") -> bool:
    """
    Return True if the given FI number is inapplicable for *case_type*.

    Wrongful-death FIs (10.x series) are inapplicable for non-WD cases.
    Employment FIs (200+ series) are inapplicable for non-employment cases.
    """
    if is_fi_series(fi_number, _WRONGFUL_DEATH_SERIES):
        return case_type != "wrongful_death"

    if is_fi_series(fi_number, _EMPLOYMENT_SERIES):
        return case_type != "employment"

    return False


# ---------------------------------------------------------------------------
# LLM prompt builders
# ---------------------------------------------------------------------------

_SI_STYLE_INSTRUCTIONS = {
    "minimal": (
        "Draft as narrowly as possible. Use as few words as possible. "
        "Only provide the information specifically requested. "
        "If the wording allows a vague response, prefer the vague response. "
        "Avoid providing information damaging to the client."
    ),
    "moderate": "Draft standard responsive answers.",
    "detailed": "Provide full, thorough responsive answers.",
}

_RFA_POSTURE_INSTRUCTIONS = {
    "cautious": (
        "Lean toward Deny or the insufficient information response. "
        "Only use Admit when the fact is definitively not in dispute."
    ),
    "balanced": "Use your best judgment.",
    "cooperative": "Lean toward admitting undisputed facts.",
}

_RPD_POSTURE_INSTRUCTIONS = {
    "context_dependent": (
        "Evaluate each request individually. "
        "Where documents are available and non-privileged, respond 'Will comply.' "
        "Where documents are unavailable or the request is overbroad, respond "
        "'Unable to comply' with a brief explanation."
    ),
    "produce_all": (
        "Default to 'Will comply' for all requests unless there is a clear privilege issue."
    ),
    "withhold_pending_protective": (
        "Default to withholding documents pending entry of a protective order."
    ),
}


def build_si_prompt(
    request_text: str,
    context_text: str,
    rules: ResponseRules,
) -> str:
    """Build an LLM prompt for drafting a Special Interrogatory substantive response."""
    style_instr = _SI_STYLE_INSTRUCTIONS.get(
        rules.si_response_style,
        _SI_STYLE_INSTRUCTIONS["moderate"],
    )

    parts = [
        "You are a California civil litigation attorney drafting discovery responses.",
        "",
        f"DRAFTING INSTRUCTION: {style_instr}",
        "",
        "Draft a substantive response to the following Special Interrogatory.",
        "Do not include objections — those are handled separately.",
        "Write only the factual, responsive answer.",
        "",
        f"INTERROGATORY:\n{request_text}",
        "",
        f"CASE CONTEXT:\n{context_text}",
    ]

    if rules.custom_instructions:
        parts += ["", f"ADDITIONAL INSTRUCTIONS:\n{rules.custom_instructions}"]

    return "\n".join(parts)


def build_rfa_prompt(
    request_text: str,
    context_text: str,
    rules: ResponseRules,
) -> str:
    """Build an LLM prompt for drafting a Request for Admission response."""
    posture_instr = _RFA_POSTURE_INSTRUCTIONS.get(
        rules.rfa_default_posture,
        _RFA_POSTURE_INSTRUCTIONS["cautious"],
    )

    parts = [
        "You are a California civil litigation attorney drafting responses to Requests for Admission.",
        "",
        "You must choose exactly one of the three California-approved response forms:",
        "  1. Admit",
        "  2. Deny",
        "  3. After a reasonable inquiry concerning the matter in this request, the information "
        "known or readily obtainable is insufficient to enable this party to admit the matter.",
        "",
        f"POSTURE INSTRUCTION: {posture_instr}",
        "",
        f"REQUEST FOR ADMISSION:\n{request_text}",
        "",
        f"CASE CONTEXT:\n{context_text}",
        "",
        "Respond with only the chosen response form, optionally followed by a brief explanation.",
    ]

    if rules.custom_instructions:
        parts += ["", f"ADDITIONAL INSTRUCTIONS:\n{rules.custom_instructions}"]

    return "\n".join(parts)


def build_rpd_prompt(
    request_text: str,
    context_text: str,
    rules: ResponseRules,
) -> str:
    """Build an LLM prompt for drafting a Request for Production response."""
    posture_instr = _RPD_POSTURE_INSTRUCTIONS.get(
        rules.rpd_default_posture,
        _RPD_POSTURE_INSTRUCTIONS["context_dependent"],
    )

    parts = [
        "You are a California civil litigation attorney drafting responses to Requests for Production.",
        "",
        "You must choose one of these response forms (or combine them):",
        "  1. Responding Party will comply with this request and produce all responsive, "
        "non-privileged documents in its possession, custody, or control.",
        "  2. Responding Party is unable to comply with this request because no responsive "
        "documents exist within Responding Party's possession, custody, or control.",
        "",
        f"POSTURE INSTRUCTION: {posture_instr}",
        "",
        f"REQUEST FOR PRODUCTION:\n{request_text}",
        "",
        f"CASE CONTEXT:\n{context_text}",
        "",
        "Write only the substantive response. Do not include objections.",
    ]

    if rules.custom_instructions:
        parts += ["", f"ADDITIONAL INSTRUCTIONS:\n{rules.custom_instructions}"]

    return "\n".join(parts)


def build_fi_17_1_prompt(rfa_responses_text: str, rules: ResponseRules) -> str:
    """
    Build an LLM prompt for drafting the FI 17.1 response.

    FI 17.1 asks: for each RFA that was not admitted, identify all facts,
    witnesses, and documents supporting the denial.
    """
    parts = [
        "You are a California civil litigation attorney drafting the response to Form Interrogatory 17.1.",
        "",
        "Form Interrogatory 17.1 requires the responding party to:",
        "  For each Request for Admission in the preceding set that is not admitted in full, state:",
        "  (a) the number of the request;",
        "  (b) all facts upon which the denial or qualified response is based;",
        "  (c) the name, ADDRESS, and telephone number of all PERSONS who have knowledge of those facts;",
        "  (d) the name, ADDRESS, and telephone number of each PERSON who has each DOCUMENT.",
        "",
        "Based on the RFA responses below, draft the response to FI 17.1.",
        "For each denied or qualified RFA, provide the required sub-parts (a) through (d).",
        "Use '[Unknown]' where information is not yet available.",
        "",
        f"RFA RESPONSES:\n{rfa_responses_text}",
    ]

    if rules.custom_instructions:
        parts += ["", f"ADDITIONAL INSTRUCTIONS:\n{rules.custom_instructions}"]

    return "\n".join(parts)


# ---------------------------------------------------------------------------
# assemble_plain_text
# ---------------------------------------------------------------------------

_INTRO_TEMPLATES = {
    "FI": "intro_template_fi",
    "SI": "intro_template_si",
    "RFA": "intro_template_rfa",
    "RPD": "intro_template_rpd",
}

_PRELIMINARY_ATTRS = {
    "FI": "preliminary_statement_fi",
    "SI": "preliminary_statement_si",
    "RFA": "preliminary_statement_rfa",
    "RPD": "preliminary_statement_rpd",
}

_TYPES_WITH_GENERAL_OBJECTIONS = {"RFA", "RPD"}


def _short_name(full_name: str) -> str:
    """Extract a short party name (last word / last segment after whitespace)."""
    parts = full_name.strip().split()
    return parts[-1] if parts else full_name


def assemble_plain_text(
    parsed: ParsedDiscovery,
    rules: ResponseRules,
    objections_map: Dict[str, str],
    responses_map: Dict[str, str],
) -> str:
    """
    Assemble the complete plain-text discovery response document.

    Parameters
    ----------
    parsed:
        The parsed incoming discovery document.
    rules:
        Strategy and boilerplate rules.
    objections_map:
        Mapping of request number → selected objection text.
    responses_map:
        Mapping of request number → substantive response text.

    Returns
    -------
    str
        The full response document as plain text.
    """
    dtype = parsed.discovery_type
    prop = parsed.propounding_party
    resp = parsed.responding_party
    set_num = parsed.set_number
    set_word = parsed.set_word  # e.g. "ONE"

    prop_short = _short_name(prop)
    resp_short = _short_name(resp)
    set_word_title = set_word.title()  # "One"

    # ------------------------------------------------------------------
    # 1. Party identification block
    # ------------------------------------------------------------------
    lines: List[str] = []
    lines.append(f"PROPOUNDING PARTY: {prop}")
    lines.append(f"RESPONDING PARTY:  {resp}")
    lines.append(f"SET NO.:           {set_word} ({set_num})")
    lines.append("")

    # ------------------------------------------------------------------
    # 2. "TO" line
    # ------------------------------------------------------------------
    lines.append(f"TO {prop.upper()} AND TO ITS ATTORNEYS OF RECORD:")
    lines.append("")

    # ------------------------------------------------------------------
    # 3. Introduction paragraph
    # ------------------------------------------------------------------
    intro_attr = _INTRO_TEMPLATES.get(dtype, "intro_template_si")
    intro_tmpl: str = getattr(rules, intro_attr, "")
    intro = intro_tmpl.format(
        responding_party=resp,
        responding_short=resp_short,
        propounding_party=prop,
        propounding_short=prop_short,
        set_word_title=set_word_title,
    )
    lines.append(intro)
    lines.append("")

    # ------------------------------------------------------------------
    # 4. Preliminary statement
    # ------------------------------------------------------------------
    prelim_attr = _PRELIMINARY_ATTRS.get(dtype, "preliminary_statement_si")
    prelim_tmpl: str = getattr(rules, prelim_attr, "")
    # Some preliminary statements have {responding_party} placeholder (RPD)
    try:
        prelim = prelim_tmpl.format(responding_party=resp)
    except KeyError:
        prelim = prelim_tmpl

    lines.append("PRELIMINARY STATEMENT")
    lines.append("")
    lines.append(prelim)
    lines.append("")

    # ------------------------------------------------------------------
    # 5. General objections (RFA and RPD only)
    # ------------------------------------------------------------------
    if dtype in _TYPES_WITH_GENERAL_OBJECTIONS:
        if dtype == "RFA":
            gen_obj = rules.general_objections_rfa
        else:
            gen_obj = rules.general_objections_rpd

        lines.append("GENERAL OBJECTIONS")
        lines.append("")
        lines.append(gen_obj)
        lines.append("")

    # ------------------------------------------------------------------
    # 6. Individual responses
    # ------------------------------------------------------------------
    for req in parsed.requests:
        num = req.number
        obj_text = objections_map.get(num, "")
        sub_text = responses_map.get(num, "")

        # Waiver language only makes sense when there are objections AND a
        # substantive response follows them.
        waiver = rules.waiver_language if (obj_text and sub_text) else ""
        reservation = rules.reservation_clause if sub_text else ""

        block = format_single_response(
            disc_type=dtype,
            request_number=num,
            request_text=req.text,
            objections=obj_text,
            substantive=sub_text,
            waiver=waiver,
            reservation=reservation,
        )
        lines.append(block)
        lines.append("")

    return "\n".join(lines)
