"""Proposal generation for the Respond to Discovery wizard."""
from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Callable, Iterable

from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
from icharlotte_core.discovery.response_review_state import RequestReview, ReviewState
from icharlotte_core.discovery.response_rule_library import (
    RuleCategory,
    RuleMode,
    ResponseRule,
)
from icharlotte_core.discovery.response_rules import ResponseRules


# ---------------------------------------------------------------------------
# Canonical Request-for-Production substantive responses
# ---------------------------------------------------------------------------
# Firm policy: an auto-generated RPD substantive response must be EXACTLY one of
# these two statements (or empty, for an objection-only response). These are the
# single source of truth -- the wizard UI imports them so the quick-insert
# buttons and the auto-drafting path stay identical, character-for-character.

RPD_WILL_COMPLY_RESPONSE = (
    "Responding Party will comply with this request and produce all "
    "non-privileged documents in Responding Party's possession, custody and "
    "control that Responding Party understands to be responsive to this Request. "
    "Responding Party identifies and refers to the documents produced "
    "concurrently herewith."
)

RPD_UNABLE_TO_COMPLY_RESPONSE = (
    "Upon a diligent search and reasonable inquiry made in an effort to locate "
    "the item(s) requested, Responding Party is unable to comply with this "
    "request at this time because the documents responsive to this request, if "
    "they exist, are not in the possession, custody or control of Responding Party."
)

# Distinctive markers used to classify a paraphrased draft. Inability markers are
# checked first, so a hedged draft mentioning both intents resolves to "unable".
_RPD_UNABLE_MARKERS = (
    "unable to comply",
    "cannot comply",
    "can not comply",
    "unable to produce",
    "not in the possession",
    "no responsive document",
    "no documents",
    "do not exist",
    "does not exist",
    "diligent search",
    "unable to locate",
    "no such document",
)

_RPD_COMPLY_MARKERS = (
    "will comply",
    "will produce",
    "agrees to produce",
    "agree to produce",
    "produce all",
    "produce responsive",
    "produce the responsive",
    "identifies and refers",
)


def normalize_rpd_substantive_response(text: str) -> str:
    """Snap an RPD substantive response to one of the two canonical statements.

    - Empty/whitespace input returns ``""`` (an objection-only response is
      permitted; we do not fabricate a substantive statement).
    - Text already equal to a canonical statement is returned unchanged.
    - Otherwise the draft is classified by intent and replaced with the verbatim
      canonical statement: inability markers win, then compliance markers, and a
      genuinely ambiguous non-empty draft defaults to the will-comply statement.
    """
    stripped = (text or "").strip()
    if not stripped:
        return ""
    if stripped == RPD_WILL_COMPLY_RESPONSE:
        return RPD_WILL_COMPLY_RESPONSE
    if stripped == RPD_UNABLE_TO_COMPLY_RESPONSE:
        return RPD_UNABLE_TO_COMPLY_RESPONSE
    lowered = stripped.lower()
    if any(marker in lowered for marker in _RPD_UNABLE_MARKERS):
        return RPD_UNABLE_TO_COMPLY_RESPONSE
    if any(marker in lowered for marker in _RPD_COMPLY_MARKERS):
        return RPD_WILL_COMPLY_RESPONSE
    return RPD_WILL_COMPLY_RESPONSE


@dataclass(frozen=True)
class ConditionalRuleDecision:
    applies: bool
    term: str = ""


@dataclass(frozen=True)
class StructuredProposal:
    request_number: str
    conditional_objection_rule_ids: list[str] | None = None
    applied_custom_rule_ids: list[str] | None = None
    applied_instruction_rule_ids: list[str] | None = None
    ambiguous_term: str = ""
    proposed_objections: str = ""
    proposed_substantive_response: str = ""
    needs_review: bool = False
    review_reason: str = ""


RuleDecisionCallback = Callable[
    [ResponseRule, ParsedRequest, ParsedDiscovery, str],
    ConditionalRuleDecision | bool | tuple[bool, str],
]
DraftSubstantiveCallback = Callable[
    [ParsedRequest, ParsedDiscovery, str, list[ResponseRule]],
    str,
]
StructuredProposalCallback = Callable[
    [ParsedRequest, ParsedDiscovery, str, list[ResponseRule], ResponseRules],
    StructuredProposal,
]


@dataclass
class DraftCallbacks:
    should_apply_rule: RuleDecisionCallback | None = None
    draft_substantive: DraftSubstantiveCallback | None = None
    structured_proposal: StructuredProposalCallback | None = None


def generate_review_state(
    parsed: ParsedDiscovery,
    selected_rules: Iterable[ResponseRule],
    context_text: str = "",
    response_rules: ResponseRules | None = None,
    callbacks: DraftCallbacks | None = None,
    fi_mode: str = "custom",
) -> ReviewState:
    """Generate proposed per-request review state from rules and context."""
    response_rules = response_rules or ResponseRules()
    callbacks = callbacks or DraftCallbacks()
    rules = list(selected_rules or [])
    dtype = (parsed.discovery_type or "").upper()

    if dtype == "FI" and fi_mode == "fixed":
        return _generate_fixed_fi_state(parsed, response_rules, callbacks, context_text)

    reviews: list[RequestReview] = []
    for req in parsed.requests:
        objections: list[str] = []
        applied_rule_ids: list[str] = []

        if callbacks.structured_proposal:
            proposal = callbacks.structured_proposal(
                req,
                parsed,
                context_text,
                rules,
                response_rules,
            )
            reviews.append(apply_structured_proposal(req, parsed, rules, proposal))
            continue

        for rule in rules:
            if rule.category != RuleCategory.OBJECTION:
                continue
            decision = _decide_rule(rule, req, parsed, context_text, callbacks)
            if not decision.applies:
                continue
            objections.append(_format_rule_text(rule, req, decision.term))
            applied_rule_ids.append(rule.id)

        instruction_rules = [r for r in rules if r.category == RuleCategory.SUBSTANTIVE]
        applied_rule_ids.extend(r.id for r in instruction_rules)
        substantive = _draft_substantive(
            req,
            parsed,
            context_text,
            rules,
            response_rules,
            callbacks,
        )

        reviews.append(
            RequestReview(
                number=req.number,
                request_text=req.text,
                proposed_objections=_join_objections(objections),
                proposed_substantive_response=substantive,
                selected_rule_ids=applied_rule_ids,
                approved=False,
            )
        )
    return ReviewState(reviews)


def parse_structured_proposal_response(llm_text: str) -> StructuredProposal:
    text = (llm_text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValueError("Structured proposal response must be a JSON object")
    return StructuredProposal(
        request_number=str(data.get("request_number", "")),
        conditional_objection_rule_ids=_string_list(
            data.get("conditional_objection_rule_ids")
        ),
        applied_custom_rule_ids=_string_list(data.get("applied_custom_rule_ids")),
        applied_instruction_rule_ids=_string_list(
            data.get("applied_instruction_rule_ids")
        ),
        ambiguous_term=str(data.get("ambiguous_term", "")),
        proposed_objections=str(data.get("proposed_objections", "")),
        proposed_substantive_response=str(
            data.get("proposed_substantive_response", "")
        ),
        needs_review=bool(data.get("needs_review", False)),
        review_reason=str(data.get("review_reason", "")),
    )


def apply_structured_proposal(
    request: ParsedRequest,
    parsed: ParsedDiscovery,
    selected_rules: list[ResponseRule],
    proposal: StructuredProposal,
) -> RequestReview:
    rules_by_id = {rule.id: rule for rule in selected_rules}
    ambiguous_term = _grounded_ambiguous_term(request, proposal.ambiguous_term)
    proposal_rule_ids = (
        list(proposal.conditional_objection_rule_ids or [])
        + list(proposal.applied_custom_rule_ids or [])
        + list(proposal.applied_instruction_rule_ids or [])
    )
    unknown_ids = [rid for rid in proposal_rule_ids if rid not in rules_by_id]
    mandatory_ids = [
        rule.id
        for rule in selected_rules
        if rule.category == RuleCategory.OBJECTION and rule.mode == RuleMode.MANDATORY
    ]
    requested_ids = (
        mandatory_ids
        + list(proposal.conditional_objection_rule_ids or [])
        + [
            rid
            for rid in (proposal.applied_custom_rule_ids or [])
            if rules_by_id.get(rid) and rules_by_id[rid].category == RuleCategory.OBJECTION
        ]
    )

    objections = [
        _format_rule_text(rules_by_id[rid], request, ambiguous_term)
        for rid in requested_ids
        if rid in rules_by_id and rules_by_id[rid].category == RuleCategory.OBJECTION
    ]
    instruction_ids = [
        rid
        for rid in (
            list(proposal.applied_instruction_rule_ids or [])
            + [
                rid
                for rid in (proposal.applied_custom_rule_ids or [])
                if rules_by_id.get(rid)
                and rules_by_id[rid].category == RuleCategory.SUBSTANTIVE
            ]
        )
        if rid in rules_by_id
    ]
    selected_rule_ids = list(
        dict.fromkeys([rid for rid in requested_ids + instruction_ids if rid in rules_by_id])
    )
    proposal_request_number = (proposal.request_number or "").strip()
    request_number_mismatch = (
        bool(proposal_request_number) and proposal_request_number != request.number
    )
    needs_review = proposal.needs_review or bool(unknown_ids) or request_number_mismatch
    review_reason = proposal.review_reason.strip()
    if unknown_ids:
        unknown_message = "Unknown rule ID returned by model: " + ", ".join(
            sorted(set(unknown_ids))
        )
        review_reason = f"{review_reason} {unknown_message}".strip()
    if request_number_mismatch:
        mismatch_message = (
            "Proposal request number mismatch: "
            f"expected {request.number}, got {proposal_request_number}."
        )
        review_reason = f"{review_reason} {mismatch_message}".strip()

    from icharlotte_core.discovery.response_type_detector import (
        normalize_discovery_type,
    )

    substantive_response = proposal.proposed_substantive_response.strip()
    if normalize_discovery_type(parsed.discovery_type) == "RPD":
        substantive_response = normalize_rpd_substantive_response(substantive_response)

    return RequestReview(
        number=request.number,
        request_text=request.text,
        proposed_objections=_join_objections(objections),
        proposed_substantive_response=substantive_response,
        selected_rule_ids=selected_rule_ids,
        approved=False,
        needs_review=needs_review,
        review_reason=review_reason,
    )


def _grounded_ambiguous_term(request: ParsedRequest, proposed_term: str) -> str:
    term = (proposed_term or "").strip()
    if not term:
        return ""
    for defined_term in request.defined_terms_used or []:
        if term.lower() == defined_term.lower():
            return defined_term
    if term.lower() in (request.text or "").lower():
        return term
    return ""


def build_structured_proposal_prompt(
    request: ParsedRequest,
    parsed: ParsedDiscovery,
    context_packet: str,
    selected_rules: list[ResponseRule],
    response_rules: ResponseRules,
) -> str:
    conditional_rules = [
        rule for rule in selected_rules
        if rule.category == RuleCategory.OBJECTION and rule.mode == RuleMode.CONDITIONAL
    ]
    custom_rules = [rule for rule in selected_rules if rule.id.startswith("custom_")]
    instruction_rules = [
        rule for rule in selected_rules
        if rule.category == RuleCategory.SUBSTANTIVE
    ]
    responding_party = (parsed.responding_party or "").strip() or "[responding party not identified]"
    propounding_party = (parsed.propounding_party or "").strip() or "[propounding party not identified]"
    return "\n".join(
        [
            "You are a California civil litigation attorney drafting proposed discovery responses.",
            "Return ONLY a JSON object matching the requested schema.",
            "",
            "PARTIES:",
            f"- RESPONDING PARTY (your client; draft every response on behalf of the responding party): {responding_party}",
            f"- PROPOUNDING PARTY (the opposing party who served these requests; never answer on their behalf): {propounding_party}",
            "",
            f"DISCOVERY TYPE: {(parsed.discovery_type or '').upper()}",
            f"REQUEST NUMBER: {request.number}",
            f"REQUEST TEXT:\n{request.text}",
            "",
            f"RESPONSE POSTURE:\n{_response_posture(parsed, response_rules)}",
            "",
            "CONDITIONAL OBJECTION RULES:",
            _format_rules_for_prompt(conditional_rules),
            "",
            "CUSTOM RULES:",
            _format_rules_for_prompt(custom_rules),
            "",
            "SUBSTANTIVE INSTRUCTION RULES:",
            _format_rules_for_prompt(instruction_rules),
            "",
            f"REQUEST-SPECIFIC CONTEXT PACKET:\n{context_packet or '[NO SPECIFIC CONTEXT FOUND]'}",
            "",
            "Hard requirements:",
            "- Draft every response on behalf of the responding party identified above. Never answer from the propounding party's perspective.",
            '- In interrogatories, the words "you" and "your" refer to the responding party. Supply the responding party\'s own information, not the propounding party\'s.',
            "- The context packet may contain documents authored by, or describing, the propounding party (e.g. their complaint or claims). Treat those as the opposing party's material, not the responding party's own statements.",
            "- Do not include objections in proposed_substantive_response.",
            "- Do not include waiver or reservation language.",
            "- Do not invent facts not supported by the context packet.",
            "- If context is weak, use cautious default language and set needs_review true.",
            "- Select only objection rule IDs that apply.",
            "- Do not draft objection text. The application formats objections from the rule IDs you select.",
            "",
            "JSON schema:",
            "{",
            '  "request_number": "string",',
            '  "conditional_objection_rule_ids": ["rule_id"],',
            '  "applied_custom_rule_ids": ["rule_id"],',
            '  "applied_instruction_rule_ids": ["rule_id"],',
            '  "ambiguous_term": "string",',
            '  "proposed_substantive_response": "string",',
            '  "needs_review": false,',
            '  "review_reason": "string"',
            "}",
        ]
    )


def build_fallback_structured_proposal(
    request: ParsedRequest,
    parsed: ParsedDiscovery,
    context_packet: str,
) -> StructuredProposal:
    dtype = (parsed.discovery_type or "").upper()
    weak_context = not bool((context_packet or "").strip())
    review_reason = "No specific context found." if weak_context else "Model response could not be parsed."
    if dtype == "RFA":
        substantive = (
            "After a reasonable inquiry concerning the matter in this request, "
            "the information known or readily obtainable to Responding Party is "
            "insufficient to enable Responding Party to admit the matter."
        )
    elif dtype == "RPD":
        substantive = RPD_UNABLE_TO_COMPLY_RESPONSE
    elif dtype in {"FI", "SI"}:
        substantive = "Responding Party lacks sufficient information to provide a further substantive response at this time."
    else:
        substantive = ""
    return StructuredProposal(
        request_number=request.number,
        conditional_objection_rule_ids=[],
        applied_custom_rule_ids=[],
        applied_instruction_rule_ids=[],
        proposed_substantive_response=substantive,
        needs_review=True,
        review_reason=review_reason,
    )


def _format_rules_for_prompt(rules: list[ResponseRule]) -> str:
    if not rules:
        return "[none]"
    return "\n".join(
        f"- {rule.id}: {rule.name}\n  Description: {rule.description}\n  Text: {rule.output_text}"
        for rule in rules
    )


def _response_posture(parsed: ParsedDiscovery, response_rules: ResponseRules) -> str:
    dtype = (parsed.discovery_type or "").upper()
    if dtype == "RFA":
        return (
            "Use Admit, Deny, or the insufficient-information response. "
            "Admit only when the context clearly supports admission."
        )
    if dtype == "RPD":
        return (
            "Set proposed_substantive_response to EXACTLY one of these two "
            "verbatim statements (copy character-for-character; do not paraphrase, "
            "shorten, or add words):\n"
            f'(A) "{RPD_WILL_COMPLY_RESPONSE}"\n'
            f'(B) "{RPD_UNABLE_TO_COMPLY_RESPONSE}"\n'
            "Use (A) when responsive, non-privileged documents exist or will be "
            "produced. Use (B) when no responsive documents are in the responding "
            "party's possession, custody, or control. If you are only objecting "
            "and will neither produce nor withhold documents, leave "
            "proposed_substantive_response empty."
        )
    if dtype == "FI":
        return "Draft only the substantive factual response. Answer narrowly."
    return (
        "Answer only the question being asked using as few words as possible. "
        "Do not volunteer extra facts."
    )


def _string_list(value) -> list[str]:
    if not isinstance(value, list):
        return []
    return [str(item) for item in value if str(item).strip()]


def _generate_fixed_fi_state(
    parsed: ParsedDiscovery,
    response_rules: ResponseRules,
    callbacks: DraftCallbacks,
    context_text: str,
) -> ReviewState:
    from icharlotte_core.discovery.response_drafter import (
        detect_inapplicable_fi,
        get_fi_fixed_objections,
        get_fi_fixed_response,
        strip_fi_objections_from_fixed_response,
    )
    reviews: list[RequestReview] = []
    for req in parsed.requests:
        objections = get_fi_fixed_objections(req.number, response_rules)
        if detect_inapplicable_fi(req.number):
            substantive = "This interrogatory is not applicable to the present action."
        else:
            substantive = get_fi_fixed_response(req.number, response_rules)
            if substantive is not None:
                substantive = strip_fi_objections_from_fixed_response(
                    req.number,
                    substantive,
                    response_rules,
                )
        if substantive is None:
            substantive = _draft_substantive(
                req,
                parsed,
                context_text,
                [],
                response_rules,
                callbacks,
            )
        reviews.append(
            RequestReview(
                number=req.number,
                request_text=req.text,
                proposed_objections=objections,
                proposed_substantive_response=substantive or "",
                selected_rule_ids=["fi_fixed_objections_responses"],
                approved=False,
            )
        )
    return ReviewState(reviews)


def _decide_rule(
    rule: ResponseRule,
    request: ParsedRequest,
    parsed: ParsedDiscovery,
    context_text: str,
    callbacks: DraftCallbacks,
) -> ConditionalRuleDecision:
    if rule.mode == RuleMode.MANDATORY:
        return ConditionalRuleDecision(True)
    if rule.mode != RuleMode.CONDITIONAL:
        return ConditionalRuleDecision(False)
    if callbacks.should_apply_rule:
        return _coerce_decision(
            callbacks.should_apply_rule(rule, request, parsed, context_text)
        )
    return _heuristic_rule_decision(rule, request)


def _coerce_decision(value) -> ConditionalRuleDecision:
    if isinstance(value, ConditionalRuleDecision):
        return value
    if isinstance(value, tuple):
        applies = bool(value[0]) if value else False
        term = str(value[1]) if len(value) > 1 and value[1] else ""
        return ConditionalRuleDecision(applies, term)
    return ConditionalRuleDecision(bool(value))


def _heuristic_rule_decision(
    rule: ResponseRule,
    request: ParsedRequest,
) -> ConditionalRuleDecision:
    text = request.text or ""
    lower = text.lower()

    if rule.id == "ambiguous_term_when_unclear":
        if request.defined_terms_used:
            return ConditionalRuleDecision(True, request.defined_terms_used[0])
        match = re.search(r"\b[A-Z][A-Z0-9_ -]{3,}\b", text)
        if match:
            return ConditionalRuleDecision(True, match.group(0).strip())
        return ConditionalRuleDecision(False)

    if rule.id == "relevance_privacy_when_unrelated":
        privacy_terms = (
            "social security", "tax", "bank", "financial", "medical",
            "health", "family", "address", "phone", "private",
        )
        return ConditionalRuleDecision(any(term in lower for term in privacy_terms))

    if rule.id == "burdensome_when_lots_of_information":
        broad_terms = (
            "all documents", "all communications", "each and every",
            "any and all", "every person", "all facts", "all evidence",
        )
        return ConditionalRuleDecision(any(term in lower for term in broad_terms))

    if rule.id == "privilege_when_potentially_privileged":
        privilege_terms = (
            "attorney", "counsel", "lawyer", "legal advice", "work product",
            "investigation", "claim file", "communication",
        )
        return ConditionalRuleDecision(any(term in lower for term in privilege_terms))

    if rule.id == "expert_legal_conclusion_when_called_for":
        expert_terms = (
            "expert", "opinion", "legal conclusion", "conclude",
            "standard of care", "reasonable", "negligent",
        )
        return ConditionalRuleDecision(any(term in lower for term in expert_terms))

    return ConditionalRuleDecision(False)


def _format_rule_text(
    rule: ResponseRule,
    request: ParsedRequest,
    term: str = "",
) -> str:
    text = rule.output_text.strip()
    if "{term}" in text:
        replacement = term or (request.defined_terms_used[0] if request.defined_terms_used else "term")
        text = text.replace("{term}", replacement)
    return text


def _join_objections(objections: Iterable[str]) -> str:
    seen: set[str] = set()
    parts: list[str] = []
    for objection in objections:
        text = (objection or "").strip()
        if text and text not in seen:
            seen.add(text)
            parts.append(text)
    return "\n\n".join(parts)


def _draft_substantive(
    request: ParsedRequest,
    parsed: ParsedDiscovery,
    context_text: str,
    selected_rules: list[ResponseRule],
    response_rules: ResponseRules,
    callbacks: DraftCallbacks,
) -> str:
    if callbacks.draft_substantive:
        return callbacks.draft_substantive(
            request,
            parsed,
            context_text,
            selected_rules,
        ) or ""

    dtype = (parsed.discovery_type or "").upper()
    if dtype == "FI":
        from icharlotte_core.discovery.response_drafter import get_fi_fixed_response
        return get_fi_fixed_response(request.number, response_rules) or ""
    return ""


def _build_pending_review_state(
    parsed: ParsedDiscovery,
    response_rules: ResponseRules,
    fi_mode: str,
) -> ReviewState:
    """Build a ReviewState where every non-skipped request is pending.

    Skipped requests (FI fixed responses, FI inapplicable) get their fixed
    text immediately and ``is_pending=False``. All other requests get empty
    draft text, ``is_pending=True``, and ``review_reason='Generating...'``.
    """
    from icharlotte_core.discovery.response_drafter import (
        detect_inapplicable_fi,
        get_fi_fixed_objections,
        get_fi_fixed_response,
        strip_fi_objections_from_fixed_response,
    )
    from icharlotte_core.discovery.response_type_detector import normalize_discovery_type

    dtype = normalize_discovery_type(parsed.discovery_type)
    is_fi_fixed = dtype == "FI" and fi_mode == "fixed"

    reviews: list[RequestReview] = []
    for req in parsed.requests:
        if is_fi_fixed and detect_inapplicable_fi(req.number):
            reviews.append(
                RequestReview(
                    number=req.number,
                    request_text=req.text,
                    proposed_objections=get_fi_fixed_objections(
                        req.number, response_rules,
                    ),
                    proposed_substantive_response=(
                        "This interrogatory is not applicable to the present action."
                    ),
                    selected_rule_ids=["fi_fixed_objections_responses"],
                    is_pending=False,
                )
            )
            continue
        if is_fi_fixed:
            fixed = get_fi_fixed_response(req.number, response_rules)
            if fixed is not None:
                cleaned = strip_fi_objections_from_fixed_response(
                    req.number, fixed, response_rules,
                )
                reviews.append(
                    RequestReview(
                        number=req.number,
                        request_text=req.text,
                        proposed_objections=get_fi_fixed_objections(
                            req.number, response_rules,
                        ),
                        proposed_substantive_response=cleaned,
                        selected_rule_ids=["fi_fixed_objections_responses"],
                        is_pending=False,
                    )
                )
                continue
        reviews.append(
            RequestReview(
                number=req.number,
                request_text=req.text,
                proposed_objections="",
                proposed_substantive_response="",
                selected_rule_ids=[],
                is_pending=True,
                needs_review=False,
                review_reason="Generating...",
            )
        )
    return ReviewState(reviews)


def _apply_proposal_to_review_state(
    review_state: ReviewState,
    req_number: str,
    proposal: StructuredProposal,
    parsed: ParsedDiscovery,
    selected_rules: list[ResponseRule],
    response_rules: ResponseRules,
    fi_mode: str,
) -> RequestReview | None:
    """Merge ``proposal`` into the matching row of ``review_state``.

    Returns the updated RequestReview, or None if no row with that number
    exists. The row's ``is_pending`` is cleared.
    """
    from icharlotte_core.discovery.response_type_detector import normalize_discovery_type

    target_index = next(
        (i for i, r in enumerate(review_state.requests) if r.number == req_number),
        -1,
    )
    if target_index < 0:
        return None
    parsed_request = next(
        (req for req in parsed.requests if req.number == req_number), None,
    )
    if parsed_request is None:
        return None

    new_review = apply_structured_proposal(
        request=parsed_request,
        parsed=parsed,
        selected_rules=selected_rules,
        proposal=proposal,
    )
    # Preserve user-driven selections (quick-objection IDs).
    previous = review_state.requests[target_index]
    new_review.approved = False
    new_review.selected_quick_objection_ids = list(previous.selected_quick_objection_ids)
    new_review.is_pending = False
    review_state.requests[target_index] = new_review

    # Run the fixed-FI proposal-warning pass for this single row if needed.
    if (
        normalize_discovery_type(parsed.discovery_type) == "FI"
        and fi_mode == "fixed"
    ):
        from icharlotte_core.discovery.response_drafter import (
            detect_inapplicable_fi,
            get_fi_fixed_objections,
            get_fi_fixed_response,
        )
        # In fixed mode the rule engine has no selected rules, so objections
        # come from the firm's canned per-number FI objections. Apply them for
        # rows that took the pending/LLM path — numbers with a canned objection
        # but no fixed response (e.g. 12.x/13.x/14.x/20.x) would otherwise show
        # a substantive response with no objections.
        fixed_objections = get_fi_fixed_objections(parsed_request.number, response_rules)
        if fixed_objections:
            new_review.proposed_objections = fixed_objections
        if not detect_inapplicable_fi(parsed_request.number) and (
            get_fi_fixed_response(parsed_request.number, response_rules) is None
        ):
            if proposal.needs_review:
                new_review.needs_review = True
            if proposal.review_reason.strip():
                new_review.review_reason = proposal.review_reason.strip()

    return new_review
