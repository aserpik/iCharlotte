"""Structured rules for the Respond to Discovery wizard."""
from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass
from enum import Enum
from typing import Iterable, List


class RuleMode(str, Enum):
    MANDATORY = "mandatory"
    CONDITIONAL = "conditional"
    INSTRUCTION = "instruction"


class RuleCategory(str, Enum):
    OBJECTION = "objection"
    SUBSTANTIVE = "substantive"


@dataclass(frozen=True)
class ResponseRule:
    id: str
    name: str
    category: RuleCategory
    mode: RuleMode
    applies_to: tuple[str, ...]
    description: str
    output_text: str
    enabled_by_default: bool = True
    is_global: bool = False

    def to_dict(self) -> dict:
        data = asdict(self)
        data["category"] = self.category.value
        data["mode"] = self.mode.value
        data["applies_to"] = list(self.applies_to)
        return data

    @classmethod
    def from_dict(cls, data: dict) -> "ResponseRule":
        return cls(
            id=str(data["id"]),
            name=str(data["name"]),
            category=RuleCategory(data["category"]),
            mode=RuleMode(data["mode"]),
            applies_to=tuple(data.get("applies_to", [])),
            description=str(data.get("description", "")),
            output_text=str(data.get("output_text", "")),
            enabled_by_default=bool(data.get("enabled_by_default", True)),
            is_global=bool(data.get("is_global", False)),
        )


ALWAYS_VAGUE_TEXT = (
    "Responding Party objects to this Request on the grounds that it calls for "
    "speculation and is vague, ambiguous, uncertain and overbroad."
)

AMBIGUOUS_TERM_RULE_TEXT = (
    'Responding Party specifically objects to this Interrogatory on the grounds '
    'that the term "{term}" is undefined and therefore vague, ambiguous, '
    "uncertain, confusing, unintelligible and overbroad. (Code Civ. Proc., "
    "\u00a7 2030.060, subd. (e).)"
)

RELEVANCE_PRIVACY_RULE_TEXT = (
    "Responding Party objects to this request on the grounds that it is "
    "irrelevant and not reasonably calculated to lead to the discovery of "
    "admissible evidence and seeks to invade Responding Party's privacy."
)

RELEVANCE_PRIVACY_QUICK_TEXT = (
    "Responding Party objects to this Request on the grounds that it is not "
    "relevant and not reasonably calculated to lead to the discovery of "
    "admissible evidence and seeks to invade Responding Party's privacy."
)

BURDENSOME_TEXT = (
    "Responding Party objects to this Request on the grounds that it is unduly "
    "burdensome and so overly broad and unlimited as to time and scope as to be "
    "an unwarranted annoyance, embarrassment, and is oppressive; to comply with "
    "the Request would be an undue burden and expense on Responding Party and "
    "is calculated to annoy and harass Responding Party. (See Code of Civ. "
    "Proc., \u00a7 2030.090, subd. (b); and Columbia Broadcasting System, Inc. "
    "v. Super. Ct. (1968) 263 Cal.App.2d 12, 19.)."
)

PRIVILEGE_TEXT = (
    "Responding Party objects to this request to the extent that it seeks to "
    "invade attorney client privilege and/or attorney work product privilege."
)

EXPERT_LEGAL_TEXT = (
    "Responding Party further objects to this request on the grounds that it "
    "calls for an expert opinion and a legal conclusion."
)

ARGUMENTATIVE_TEXT = (
    "Responding Party further objects to this Request on the grounds that it, "
    "as phrased, is argumentative and requires the adoption of an assumption, "
    "which is improper; the question assumes facts which may or may not be true, "
    "but the form of the question requires that the answer adopt the assumption."
)

COMPOUND_TEXT = (
    "Responding Party objects to this Interrogatory on the grounds that it is "
    "compound in form."
)

UNDEFINED_TERM_QUICK_TEXT = (
    'Responding Party specifically objects to this Request on the grounds that '
    'the term "{term}" is undefined and therefore vague, ambiguous, uncertain, '
    "confusing, unintelligible and overbroad."
)

MINIMAL_DIRECT_ANSWER_TEXT = (
    "Answer only the question being asked using as few words as possible."
)


_STANDARD_OBJECTION_RULES = (
    ResponseRule(
        id="always_vague_ambiguous_overbroad",
        name='ALWAYS include "vague, ambiguous, overbroad" objections',
        category=RuleCategory.OBJECTION,
        mode=RuleMode.MANDATORY,
        applies_to=("SI", "RFA", "RPD", "FI_CUSTOM"),
        description=(
            "If selected, this objection is inserted for every request, "
            "regardless of the request wording."
        ),
        output_text=ALWAYS_VAGUE_TEXT,
    ),
    ResponseRule(
        id="ambiguous_term_when_unclear",
        name=(
            "Include ambiguous term objection when the discovery request "
            "contains word(s) that are confusing or not obviously clear"
        ),
        category=RuleCategory.OBJECTION,
        mode=RuleMode.CONDITIONAL,
        applies_to=("SI", "RFA", "RPD", "FI_CUSTOM"),
        description=(
            "Include when a request contains an undefined, confusing, or unclear "
            "term. The term is substituted into the objection text."
        ),
        output_text=AMBIGUOUS_TERM_RULE_TEXT,
    ),
    ResponseRule(
        id="relevance_privacy_when_unrelated",
        name=(
            "Include relevance and privacy objections when Discovery asks for "
            "information not related to the case"
        ),
        category=RuleCategory.OBJECTION,
        mode=RuleMode.CONDITIONAL,
        applies_to=("SI", "RFA", "RPD", "FI_CUSTOM"),
        description=(
            "Include when the request appears unrelated to the case or invades "
            "Responding Party's privacy."
        ),
        output_text=RELEVANCE_PRIVACY_RULE_TEXT,
    ),
    ResponseRule(
        id="burdensome_when_lots_of_information",
        name=(
            "Include burdensome objections when the discovery asks for a lot of "
            "information"
        ),
        category=RuleCategory.OBJECTION,
        mode=RuleMode.CONDITIONAL,
        applies_to=("SI", "RFA", "RPD", "FI_CUSTOM"),
        description=(
            "Include when the request is broad, unlimited in time or scope, or "
            "would require excessive effort or expense."
        ),
        output_text=BURDENSOME_TEXT,
    ),
    ResponseRule(
        id="privilege_when_potentially_privileged",
        name=(
            "Include privilege objections when the discovery asks for "
            "potentially privileged information"
        ),
        category=RuleCategory.OBJECTION,
        mode=RuleMode.CONDITIONAL,
        applies_to=("SI", "RFA", "RPD", "FI_CUSTOM"),
        description=(
            "Include when the request could seek attorney-client or attorney "
            "work-product information."
        ),
        output_text=PRIVILEGE_TEXT,
    ),
    ResponseRule(
        id="expert_legal_conclusion_when_called_for",
        name=(
            "Include Expert Opinion or Legal Conclusion objections when Request "
            "calls for potential legal conclusion or expert opinion"
        ),
        category=RuleCategory.OBJECTION,
        mode=RuleMode.CONDITIONAL,
        applies_to=("SI", "RFA", "RPD", "FI_CUSTOM"),
        description=(
            "Include when the request calls for an expert opinion, expert "
            "analysis, or a legal conclusion."
        ),
        output_text=EXPERT_LEGAL_TEXT,
    ),
)

_SUBSTANTIVE_RULES = (
    ResponseRule(
        id="minimal_direct_answer",
        name=MINIMAL_DIRECT_ANSWER_TEXT,
        category=RuleCategory.SUBSTANTIVE,
        mode=RuleMode.INSTRUCTION,
        applies_to=("SI", "FI_CUSTOM"),
        description="Use narrow, direct answers and do not volunteer extra facts.",
        output_text=MINIMAL_DIRECT_ANSWER_TEXT,
    ),
)

_QUICK_RULES = (
    ResponseRule(
        "quick_vague",
        "Vague / Ambiguous / Overbroad",
        RuleCategory.OBJECTION,
        RuleMode.MANDATORY,
        ("ALL",),
        "",
        ALWAYS_VAGUE_TEXT,
    ),
    ResponseRule(
        "quick_relevance_privacy",
        "Relevance / Privacy",
        RuleCategory.OBJECTION,
        RuleMode.MANDATORY,
        ("ALL",),
        "",
        RELEVANCE_PRIVACY_QUICK_TEXT,
    ),
    ResponseRule(
        "quick_expert_legal",
        "Expert Opinion / Legal Conclusion",
        RuleCategory.OBJECTION,
        RuleMode.MANDATORY,
        ("ALL",),
        "",
        EXPERT_LEGAL_TEXT,
    ),
    ResponseRule(
        "quick_privilege",
        "Privilege",
        RuleCategory.OBJECTION,
        RuleMode.MANDATORY,
        ("ALL",),
        "",
        PRIVILEGE_TEXT,
    ),
    ResponseRule(
        "quick_burdensome",
        "Burdensome",
        RuleCategory.OBJECTION,
        RuleMode.MANDATORY,
        ("ALL",),
        "",
        BURDENSOME_TEXT,
    ),
    ResponseRule(
        "quick_argumentative",
        "Argumentative",
        RuleCategory.OBJECTION,
        RuleMode.MANDATORY,
        ("ALL",),
        "",
        ARGUMENTATIVE_TEXT,
    ),
    ResponseRule(
        "quick_compound",
        "Compound",
        RuleCategory.OBJECTION,
        RuleMode.MANDATORY,
        ("ALL",),
        "",
        COMPOUND_TEXT,
    ),
    ResponseRule(
        "quick_undefined_term",
        "Undefined Term",
        RuleCategory.OBJECTION,
        RuleMode.MANDATORY,
        ("ALL",),
        "",
        UNDEFINED_TERM_QUICK_TEXT,
    ),
)


def built_in_rules_for(discovery_type: str, fi_mode: str = "fixed") -> List[ResponseRule]:
    """Return built-in rules for the detected discovery type."""
    dtype = (discovery_type or "").upper()
    if dtype == "FI" and fi_mode == "fixed":
        return []

    applies_key = "FI_CUSTOM" if dtype == "FI" else dtype
    return [
        rule
        for rule in _STANDARD_OBJECTION_RULES + _SUBSTANTIVE_RULES
        if applies_key in rule.applies_to
    ]


def get_quick_objection_rules() -> List[ResponseRule]:
    """Return review-screen quick objection rules."""
    return list(_QUICK_RULES)


def load_global_rules(path: str) -> List[ResponseRule]:
    """Load globally saved custom rules from JSON."""
    if not path or not os.path.exists(path):
        return []
    with open(path, "r", encoding="utf-8") as fh:
        raw = json.load(fh)
    if not isinstance(raw, list):
        return []
    return [ResponseRule.from_dict(item) for item in raw if isinstance(item, dict)]


def save_global_rules(path: str, rules: Iterable[ResponseRule]) -> None:
    """Save global custom rules to JSON."""
    if not path:
        raise ValueError("A path is required to save global rules")
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump([rule.to_dict() for rule in rules], fh, indent=2)
