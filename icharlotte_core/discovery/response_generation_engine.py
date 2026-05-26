"""Proposal generation for the Respond to Discovery wizard."""
from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Callable, Iterable

from icharlotte_core.discovery.response_drafter import (
    detect_inapplicable_fi,
    get_fi_fixed_objections,
    get_fi_fixed_response,
    strip_fi_objections_from_fixed_response,
)
from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
from icharlotte_core.discovery.response_review_state import RequestReview, ReviewState
from icharlotte_core.discovery.response_rule_library import (
    RuleCategory,
    RuleMode,
    ResponseRule,
)
from icharlotte_core.discovery.response_rules import ResponseRules


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
    proposal_rule_ids = (
        list(proposal.conditional_objection_rule_ids or [])
        + list(proposal.applied_custom_rule_ids or [])
        + list(proposal.applied_instruction_rule_ids or [])
    )
    unknown_ids = [rid for rid in proposal_rule_ids if rid not in rules_by_id]
    mandatory_ids = [
        rule.id
        for rule in selected_rules
        if not unknown_ids
        and rule.category == RuleCategory.OBJECTION
        and rule.mode == RuleMode.MANDATORY
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
        _format_rule_text(rules_by_id[rid], request, proposal.ambiguous_term)
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
    needs_review = proposal.needs_review or bool(unknown_ids)
    review_reason = proposal.review_reason.strip()
    if unknown_ids:
        unknown_message = "Unknown rule ID returned by model: " + ", ".join(
            sorted(set(unknown_ids))
        )
        review_reason = f"{review_reason} {unknown_message}".strip()

    return RequestReview(
        number=request.number,
        request_text=request.text,
        proposed_objections=_join_objections(objections),
        proposed_substantive_response=proposal.proposed_substantive_response.strip(),
        selected_rule_ids=selected_rule_ids,
        approved=False,
        needs_review=needs_review,
        review_reason=review_reason,
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
        return get_fi_fixed_response(request.number, response_rules) or ""
    return ""
