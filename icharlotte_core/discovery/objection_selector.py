"""
Phase 2: Objection Selector.

Rule-based + LLM hybrid objection selection.
- Form interrogatories use fixed objections from ResponseRules.
- SI/RFA/RPD use rule-based pre-selection merged with LLM selection.

The actual LLM call happens in the UI layer; this module builds prompts
and parses responses.
"""
import json
import re
from typing import Dict, Optional, Set

from icharlotte_core.discovery.response_parser import ParsedRequest
from icharlotte_core.discovery.response_drafter import get_fi_fixed_objections
from icharlotte_core.discovery.response_rules import ResponseRules


# ---------------------------------------------------------------------------
# Default objection library
# ---------------------------------------------------------------------------

_DEFAULT_OBJECTIONS: Dict[int, str] = {
    1: (
        "Responding Party objects to this Request on the grounds that it calls for speculation "
        "and is vague, ambiguous, uncertain and overbroad."
    ),
    2: (
        "Responding Party objects to this Request on the grounds that it is not relevant and not "
        "reasonably calculated to lead to the discovery of admissible evidence and seeks to invade "
        "Responding Party's privacy."
    ),
    3: (
        "Responding Party further objects to this Request on the grounds that it seeks premature "
        "disclosure of expert opinion and/or a legal conclusion."
    ),
    4: (
        "Responding Party further objects to this Request on the grounds that it seeks to invade "
        "attorney client privilege and/or violates the attorney work-product privilege."
    ),
    5: (
        "Responding Party further objects to this Request on the grounds that it, as phrased, is "
        "argumentative and requires the adoption of an assumption, which is improper; the question "
        "assumes facts which may or may not be true, but the form of the question requires that the "
        "answer adopt the assumption."
    ),
    6: (
        "Responding Party objects to this Request on the grounds that it is unduly burdensome and so "
        "overly broad and unlimited as to time and scope as to be an unwarranted annoyance, "
        "embarrassment, and is oppressive; to comply with the Request would be an undue burden and "
        "expense on Responding Party and is calculated to annoy and harass Responding Party."
    ),
    7: (
        "Responding Party objects to this Interrogatory on the grounds that it is burdensome and "
        "harassing in that it calls for production of a list or summary where no such list or summary "
        "presently exists."
    ),
    8: (
        "Responding Party objects to this Interrogatory on the grounds that it has, in substance, been "
        "previously propounded. Continuous discovery into the same matter is oppressive, harassing, "
        "burdensome, and contrary to the legislative intent of the Discovery Act."
    ),
    9: (
        "Responding Party objects to this Interrogatory on the grounds that it is compound in form."
    ),
    10: (
        'Responding Party specifically objects to this Request on the grounds that the term "{term}" '
        "is undefined and therefore vague, ambiguous, uncertain, confusing, unintelligible and overbroad."
    ),
    11: (
        "Responding Party specifically objects to Propounding Party's definition of \"{term}\" on the "
        "grounds that it is vague, ambiguous, uncertain, confusing, unintelligible and overbroad, and "
        "is argumentative and requires the adoption of an assumption, which is improper."
    ),
    12: (
        "Responding Party objects to this Interrogatory on the grounds that the information and/or "
        "documents sought are equally and/or more readily available to Propounding Party."
    ),
}


# ---------------------------------------------------------------------------
# ObjectionMenu
# ---------------------------------------------------------------------------

class ObjectionMenu:
    """Container for the numbered objection library."""

    def __init__(self, objections: Dict[int, str]) -> None:
        self._objections: Dict[int, str] = dict(objections)

    @classmethod
    def load_defaults(cls) -> "ObjectionMenu":
        """Return an ObjectionMenu populated with the 12 standard objections."""
        return cls(_DEFAULT_OBJECTIONS)

    def get(self, objection_id: int) -> str:
        """Return the objection text for the given ID."""
        return self._objections[objection_id]

    def all_text(self) -> str:
        """Return all objections formatted with IDs, suitable for an LLM prompt."""
        lines = []
        for obj_id in sorted(self._objections):
            lines.append(f"{obj_id}. {self._objections[obj_id]}")
        return "\n".join(lines)

    @property
    def objections(self) -> Dict[int, str]:
        """The underlying dict of ID → objection text."""
        return self._objections


# ---------------------------------------------------------------------------
# Form Interrogatory fixed objections
# ---------------------------------------------------------------------------

def select_fi_objections(rules: ResponseRules, number_str: str | None = None) -> str:
    """Return fixed Form Interrogatory objections from ResponseRules."""
    if number_str:
        return get_fi_fixed_objections(number_str, rules)
    return rules.fi_objections


# ---------------------------------------------------------------------------
# Rule-based pre-selection
# ---------------------------------------------------------------------------

_EXPERT_PATTERN = re.compile(r"\b(expert|opinion)\b", re.IGNORECASE)


def rule_based_preselect(
    request: ParsedRequest,
    rules: ResponseRules,
    menu: ObjectionMenu,
) -> Set[int]:
    """
    Apply rule-based logic to pre-select objection IDs.

    Rules applied (in order):
    - Compound flag (auto_flag_compound)      → ID 9
    - always_include_privacy_objection        → ID 2
    - always_include_privilege_objection      → ID 4
    - always_include_burden_objection         → ID 6
    - Defined terms + auto_flag_broad_definitions → IDs 10 and/or 11
    - "expert" or "opinion" keyword           → ID 3
    """
    ids: Set[int] = set()

    if request.is_compound and rules.auto_flag_compound:
        ids.add(9)

    if rules.always_include_privacy_objection:
        ids.add(2)

    if rules.always_include_privilege_objection:
        ids.add(4)

    if rules.always_include_burden_objection:
        ids.add(6)

    if request.defined_terms_used and rules.auto_flag_broad_definitions:
        ids.add(10)
        ids.add(11)

    if _EXPERT_PATTERN.search(request.text):
        ids.add(3)

    return ids


# ---------------------------------------------------------------------------
# LLM prompt builder
# ---------------------------------------------------------------------------

_AGGRESSIVENESS_INSTRUCTIONS = {
    "aggressive": (
        "Select ALL objections that might apply. Err on the side of over-inclusion — "
        "a failure to include an objection waives it."
    ),
    "moderate": "Select objections that are reasonably applicable.",
    "conservative": "Select only objections that clearly and directly apply.",
}


def build_objection_prompt(
    request_text: str,
    menu: ObjectionMenu,
    aggressiveness: str = "aggressive",
) -> str:
    """
    Build an LLM prompt for objection selection.

    Args:
        request_text: The full text of the discovery request.
        menu: The ObjectionMenu containing available objections.
        aggressiveness: One of 'aggressive', 'moderate', 'conservative'.

    Returns:
        A prompt string ready to send to an LLM.
    """
    instruction = _AGGRESSIVENESS_INSTRUCTIONS.get(
        aggressiveness, _AGGRESSIVENESS_INSTRUCTIONS["aggressive"]
    )

    return (
        f"You are a California civil litigation defense attorney reviewing a discovery request.\n\n"
        f"DISCOVERY REQUEST:\n{request_text}\n\n"
        f"AVAILABLE OBJECTIONS:\n{menu.all_text()}\n\n"
        f"INSTRUCTION: {instruction}\n\n"
        f"Return ONLY the objection IDs that apply as a comma-separated list of numbers "
        f"(e.g., 1, 3, 6). Do not include any other text."
    )


# ---------------------------------------------------------------------------
# LLM response parser
# ---------------------------------------------------------------------------

def parse_objection_response(llm_text: str) -> Set[int]:
    """
    Parse an LLM response into a set of objection IDs.

    Handles:
    - JSON array: "[1, 3, 5]"
    - Comma-separated: "1, 2, 4, 6"
    - Natural language with embedded numbers: "Objections 1, 3, and 7 apply."
    """
    text = llm_text.strip()

    # Try JSON array first
    try:
        parsed = json.loads(text)
        if isinstance(parsed, list):
            return {int(x) for x in parsed if str(x).strip().isdigit() or isinstance(x, int)}
    except (json.JSONDecodeError, ValueError):
        pass

    # Fall back to extracting all integers via regex
    return {int(m) for m in re.findall(r"\d+", text)}


# ---------------------------------------------------------------------------
# Merge
# ---------------------------------------------------------------------------

def merge_objections(rule_ids: Set[int], llm_ids: Set[int]) -> Set[int]:
    """Return the union of rule-based and LLM-selected objection IDs."""
    return rule_ids | llm_ids


# ---------------------------------------------------------------------------
# Formatter
# ---------------------------------------------------------------------------

def format_objections(
    objection_ids: Set[int],
    menu: ObjectionMenu,
    term: Optional[str] = None,
) -> str:
    """
    Format selected objections into a single string.

    Args:
        objection_ids: Set of objection IDs to include.
        menu: The ObjectionMenu used for text lookup.
        term: If provided, substituted for ``{term}`` placeholders in
              objections 10 and 11.

    Returns:
        Space-separated objection sentences in ascending ID order.
    """
    parts = []
    for obj_id in sorted(objection_ids):
        text = menu.get(obj_id)
        if term is not None and "{term}" in text:
            text = text.replace("{term}", term)
        parts.append(text)
    return " ".join(parts)
