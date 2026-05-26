"""Review-state helpers for proposed discovery responses."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable

from icharlotte_core.discovery.response_rule_library import ResponseRule


@dataclass
class RequestReview:
    number: str
    request_text: str
    proposed_objections: str = ""
    proposed_substantive_response: str = ""
    selected_rule_ids: list[str] = field(default_factory=list)
    selected_quick_objection_ids: list[str] = field(default_factory=list)
    approved: bool = False
    needs_review: bool = False
    review_reason: str = ""

    def to_dict(self) -> dict:
        return {
            "number": self.number,
            "request_text": self.request_text,
            "proposed_objections": self.proposed_objections,
            "proposed_substantive_response": self.proposed_substantive_response,
            "selected_rule_ids": list(self.selected_rule_ids),
            "selected_quick_objection_ids": list(self.selected_quick_objection_ids),
            "approved": bool(self.approved),
            "needs_review": bool(self.needs_review),
            "review_reason": self.review_reason,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "RequestReview":
        return cls(
            number=str(data.get("number", "")),
            request_text=str(data.get("request_text", "")),
            proposed_objections=str(data.get("proposed_objections", "")),
            proposed_substantive_response=str(
                data.get("proposed_substantive_response", "")
            ),
            selected_rule_ids=list(data.get("selected_rule_ids", [])),
            selected_quick_objection_ids=list(
                data.get("selected_quick_objection_ids", [])
            ),
            approved=bool(data.get("approved", False)),
            needs_review=bool(data.get("needs_review", False)),
            review_reason=str(data.get("review_reason", "")),
        )


@dataclass
class ReviewState:
    requests: list[RequestReview] = field(default_factory=list)

    def all_approved(self) -> bool:
        return all(item.approved for item in self.requests)

    def to_dict(self) -> dict:
        return {"requests": [item.to_dict() for item in self.requests]}

    @classmethod
    def from_dict(cls, data: dict) -> "ReviewState":
        raw = (data or {}).get("requests", [])
        return cls(
            requests=[
                RequestReview.from_dict(item)
                for item in raw
                if isinstance(item, dict)
            ]
        )

    def to_response_maps(self) -> tuple[dict[str, str], dict[str, str]]:
        objections_map = {
            item.number: item.proposed_objections.strip()
            for item in self.requests
        }
        responses_map = {
            item.number: item.proposed_substantive_response.strip()
            for item in self.requests
        }
        return objections_map, responses_map

    def to_plain_text(self, parsed, rules) -> str:
        from icharlotte_core.discovery.response_drafter import assemble_plain_text

        objections_map, responses_map = self.to_response_maps()
        return assemble_plain_text(parsed, rules, objections_map, responses_map)


def _format_objection(rule: ResponseRule, term: str | None = None) -> str:
    text = rule.output_text.strip()
    if "{term}" in text and term:
        text = text.replace("{term}", term)
    return text


def _split_objection_blocks(text: str) -> list[str]:
    return [part.strip() for part in (text or "").split("\n\n") if part.strip()]


def _join_objection_blocks(parts: Iterable[str]) -> str:
    return "\n\n".join(part.strip() for part in parts if part and part.strip())


def insert_quick_objection(
    review: RequestReview,
    rule: ResponseRule,
    term: str | None = None,
) -> None:
    """Insert a quick objection into a review item without duplicating it."""
    formatted = _format_objection(rule, term)
    parts = _split_objection_blocks(review.proposed_objections)
    if formatted not in parts:
        parts.append(formatted)
        review.proposed_objections = _join_objection_blocks(parts)
    if rule.id not in review.selected_quick_objection_ids:
        review.selected_quick_objection_ids.append(rule.id)


def remove_quick_objection(
    review: RequestReview,
    rule: ResponseRule,
    term: str | None = None,
) -> None:
    """Remove a quick objection from a review item."""
    formatted = _format_objection(rule, term)
    parts = [p for p in _split_objection_blocks(review.proposed_objections) if p != formatted]
    review.proposed_objections = _join_objection_blocks(parts)
    review.selected_quick_objection_ids = [
        rid for rid in review.selected_quick_objection_ids if rid != rule.id
    ]
