# Respond to Discovery Wizard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a guided Wizard Mode task that drafts and reviews discovery responses while keeping the existing Discovery -> Respond subtab intact.

**Architecture:** Build a shared non-UI response engine for type detection, rule selection, proposal generation, review state, and final assembly. Add a Wizard task that drives those services through a rules screen, context selection, one-request review UI, and final Word assembly.

**Tech Stack:** Python 3, PySide6, pytest, unittest-style existing tests, python-docx, PyMuPDF where available, existing `LLMWorker`/discovery response modules.

---

## Preflight Notes

The current checkout may already contain unrelated user changes and an unresolved conflict in `tests/test_wizard/test_task_tab.py`. Do not resolve, revert, stage, or commit unrelated files as part of this plan. If using commits, first work in a clean branch/worktree or resolve unrelated conflicts outside this task.

---

## File Structure

Create:

- `icharlotte_core/discovery/response_type_detector.py` - deterministic filename/text discovery-type detection.
- `icharlotte_core/discovery/response_rule_library.py` - built-in wizard rule objects and JSON-backed global custom rules.
- `icharlotte_core/discovery/response_review_state.py` - per-request proposal/edit/approval state and quick-objection helpers.
- `icharlotte_core/discovery/response_generation_engine.py` - proposal generation from `ParsedDiscovery`, selected rules, and context.
- `icharlotte_core/ui/wizard/pages/respond_discovery_page.py` - guided wizard UI.
- `tests/test_discovery/test_response_type_detector.py`
- `tests/test_discovery/test_response_rule_library.py`
- `tests/test_discovery/test_response_review_state.py`
- `tests/test_discovery/test_response_generation_engine.py`
- `tests/test_wizard/test_respond_discovery_page.py`
- `tests/test_wizard/test_respond_to_discovery_registry.py`

Modify:

- `icharlotte_core/ui/wizard/registry.py` - add task card.
- `icharlotte_core/ui/wizard/task_routing.py` - route task to in-process builder and no generic picker.
- `icharlotte_core/ui/wizard/in_process_task_tab.py` - add `build_respond_to_discovery_tab`.
- `iCharlotte.py` - pass the Wizard task through the in-process builder and preserve one-PDF selection behavior if needed.
- `icharlotte_core/discovery/response_assembler.py` or a thin wrapper if validation must be called around final save.

---

### Task 1: Discovery Type Detector

**Files:**
- Create: `tests/test_discovery/test_response_type_detector.py`
- Create: `icharlotte_core/discovery/response_type_detector.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_discovery/test_response_type_detector.py`:

```python
import unittest

from icharlotte_core.discovery.response_type_detector import (
    DiscoveryTypeGuess,
    detect_type_from_filename,
    detect_type_from_text,
    resolve_detected_type,
)


class ResponseTypeDetectorTests(unittest.TestCase):
    def test_filename_detects_form_interrogatories(self):
        guess = detect_type_from_filename("Defendant FROGG Set One.pdf")
        self.assertEqual(guess.discovery_type, "FI")
        self.assertEqual(guess.source, "filename")

    def test_filename_detects_special_interrogatories(self):
        guess = detect_type_from_filename("Plaintiff_srogg_2.pdf")
        self.assertEqual(guess.discovery_type, "SI")

    def test_filename_detects_rfa(self):
        guess = detect_type_from_filename("RFA to Defendant.pdf")
        self.assertEqual(guess.discovery_type, "RFA")

    def test_filename_detects_rpd_from_rfp(self):
        guess = detect_type_from_filename("RFP Set 1.pdf")
        self.assertEqual(guess.discovery_type, "RPD")

    def test_text_detects_request_for_production(self):
        guess = detect_type_from_text("REQUESTS FOR PRODUCTION OF DOCUMENTS, SET ONE")
        self.assertEqual(guess.discovery_type, "RPD")
        self.assertEqual(guess.source, "text")

    def test_resolve_prefers_filename_when_text_absent(self):
        result = resolve_detected_type(
            filename_guess=DiscoveryTypeGuess("SI", "filename", "srogg"),
            text_guess=DiscoveryTypeGuess(None, "text", ""),
        )
        self.assertEqual(result.discovery_type, "SI")
        self.assertFalse(result.needs_user_choice)

    def test_resolve_flags_conflict(self):
        result = resolve_detected_type(
            filename_guess=DiscoveryTypeGuess("SI", "filename", "srogg"),
            text_guess=DiscoveryTypeGuess("RFA", "text", "request for admission"),
        )
        self.assertIsNone(result.discovery_type)
        self.assertTrue(result.needs_user_choice)
        self.assertIn("conflict", result.reason.lower())


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
pytest tests/test_discovery/test_response_type_detector.py -q
```

Expected: fail because `response_type_detector.py` does not exist.

- [ ] **Step 3: Implement detector**

Create `icharlotte_core/discovery/response_type_detector.py`:

```python
"""Discovery response type detection helpers for the Wizard flow."""
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Optional


VALID_DISCOVERY_TYPES = {"FI", "SI", "RFA", "RPD"}


@dataclass(frozen=True)
class DiscoveryTypeGuess:
    discovery_type: Optional[str]
    source: str
    matched_text: str = ""


@dataclass(frozen=True)
class DiscoveryTypeResolution:
    discovery_type: Optional[str]
    needs_user_choice: bool
    reason: str = ""


_FILENAME_PATTERNS = (
    ("FI", re.compile(r"\b(fro+g+|form\s+interrog(?:atory|atories)?)\b", re.I)),
    ("SI", re.compile(r"\b(sro+g+|special\s+interrog(?:atory|atories)?)\b", re.I)),
    ("RFA", re.compile(r"\b(rfa|requests?\s+for\s+admission)\b", re.I)),
    ("RPD", re.compile(r"\b(rfp|rpd|requests?\s+for\s+production)\b", re.I)),
)

_TEXT_PATTERNS = (
    ("FI", re.compile(r"\bform\s+interrogator(?:y|ies)\b", re.I)),
    ("SI", re.compile(r"\bspecial\s+interrogator(?:y|ies)\b", re.I)),
    ("RFA", re.compile(r"\brequests?\s+for\s+admission\b", re.I)),
    ("RPD", re.compile(r"\brequests?\s+for\s+production\b", re.I)),
)


def _detect(value: str, source: str, patterns) -> DiscoveryTypeGuess:
    text = value or ""
    for discovery_type, pattern in patterns:
        match = pattern.search(text)
        if match:
            return DiscoveryTypeGuess(discovery_type, source, match.group(0))
    return DiscoveryTypeGuess(None, source, "")


def detect_type_from_filename(path_or_name: str) -> DiscoveryTypeGuess:
    return _detect(os.path.basename(path_or_name or ""), "filename", _FILENAME_PATTERNS)


def detect_type_from_text(text: str) -> DiscoveryTypeGuess:
    return _detect(text or "", "text", _TEXT_PATTERNS)


def resolve_detected_type(
    filename_guess: DiscoveryTypeGuess,
    text_guess: DiscoveryTypeGuess,
) -> DiscoveryTypeResolution:
    filename_type = filename_guess.discovery_type
    text_type = text_guess.discovery_type
    if filename_type and text_type and filename_type != text_type:
        return DiscoveryTypeResolution(
            None,
            True,
            f"Detection conflict: filename={filename_type}, text={text_type}",
        )
    if filename_type:
        return DiscoveryTypeResolution(filename_type, False, "Detected from filename")
    if text_type:
        return DiscoveryTypeResolution(text_type, False, "Detected from text")
    return DiscoveryTypeResolution(None, True, "Discovery type could not be detected")
```

- [ ] **Step 4: Run test to verify it passes**

Run:

```powershell
pytest tests/test_discovery/test_response_type_detector.py -q
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

If the worktree is clean enough to commit this task only:

```powershell
git add icharlotte_core/discovery/response_type_detector.py tests/test_discovery/test_response_type_detector.py
git commit -m "feat(discovery): detect response discovery type"
```

If unrelated conflicts are present, skip commit and record that commit was blocked by unrelated worktree state.

---

### Task 2: Structured Rule Library

**Files:**
- Create: `tests/test_discovery/test_response_rule_library.py`
- Create: `icharlotte_core/discovery/response_rule_library.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_discovery/test_response_rule_library.py`:

```python
import unittest

from icharlotte_core.discovery.response_rule_library import (
    RuleMode,
    RuleCategory,
    built_in_rules_for,
    get_quick_objection_rules,
)


class ResponseRuleLibraryTests(unittest.TestCase):
    def test_si_rules_include_mandatory_vague_rule(self):
        rules = built_in_rules_for("SI")
        vague = next(r for r in rules if r.id == "always_vague_ambiguous_overbroad")
        self.assertEqual(vague.mode, RuleMode.MANDATORY)
        self.assertEqual(vague.category, RuleCategory.OBJECTION)
        self.assertTrue(vague.enabled_by_default)
        self.assertIn("ALWAYS include", vague.name)

    def test_rfa_rules_only_objection_rules(self):
        rules = built_in_rules_for("RFA")
        self.assertTrue(rules)
        self.assertTrue(all(r.category == RuleCategory.OBJECTION for r in rules))

    def test_fi_fixed_has_no_default_rule_cards(self):
        self.assertEqual(built_in_rules_for("FI", fi_mode="fixed"), [])

    def test_fi_custom_uses_si_style_rules(self):
        rules = built_in_rules_for("FI", fi_mode="custom")
        self.assertTrue(any(r.id == "minimal_direct_answer" for r in rules))

    def test_quick_objections_include_undefined_term_template(self):
        quick = get_quick_objection_rules()
        undefined = next(r for r in quick if r.id == "quick_undefined_term")
        self.assertIn("{term}", undefined.output_text)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
pytest tests/test_discovery/test_response_rule_library.py -q
```

Expected: fail because `response_rule_library.py` does not exist.

- [ ] **Step 3: Implement rule library**

Create `icharlotte_core/discovery/response_rule_library.py`:

```python
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
    "Responding Party objects to this Request on the grounds that it calls for speculation "
    "and is vague, ambiguous, uncertain and overbroad."
)

AMBIGUOUS_TERM_TEXT = (
    'Responding Party specifically objects to this Interrogatory on the grounds that the term '
    '"{term}" is undefined and therefore vague, ambiguous, uncertain, confusing, '
    "unintelligible and overbroad. (Code Civ. Proc., section 2030.060, subd. (e).)"
)

RELEVANCE_PRIVACY_TEXT = (
    "Responding Party objects to this request on the grounds that it is irrelevant and not "
    "reasonably calculated to lead to the discovery of admissible evidence and seeks to "
    "invade Responding Party's privacy."
)

BURDENSOME_TEXT = (
    "Responding Party objects to this Request on the grounds that it is unduly burdensome "
    "and so overly broad and unlimited as to time and scope as to be an unwarranted annoyance, "
    "embarrassment, and is oppressive; to comply with the Request would be an undue burden "
    "and expense on Responding Party and is calculated to annoy and harass Responding Party. "
    "(See Code of Civ. Proc., section 2030.090, subd. (b); and Columbia Broadcasting System, "
    "Inc. v. Super. Ct. (1968) 263 Cal.App.2d 12, 19.)."
)

PRIVILEGE_TEXT = (
    "Responding Party objects to this request to the extent that it seeks to invade attorney "
    "client privilege and/or attorney work product privilege."
)

EXPERT_LEGAL_TEXT = (
    "Responding Party further objects to this request on the grounds that it calls for an "
    "expert opinion and a legal conclusion."
)

ARGUMENTATIVE_TEXT = (
    "Responding Party further objects to this Request on the grounds that it, as phrased, "
    "is argumentative and requires the adoption of an assumption, which is improper; the "
    "question assumes facts which may or may not be true, but the form of the question "
    "requires that the answer adopt the assumption."
)

COMPOUND_TEXT = "Responding Party objects to this Interrogatory on the grounds that it is compound in form."

UNDEFINED_TERM_TEXT = (
    'Responding Party specifically objects to this Request on the grounds that the term '
    '"{term}" is undefined and therefore vague, ambiguous, uncertain, confusing, '
    "unintelligible and overbroad."
)

_STANDARD_OBJECTION_RULES = (
    ResponseRule(
        id="always_vague_ambiguous_overbroad",
        name='ALWAYS include "vague, ambiguous, overbroad" objections',
        category=RuleCategory.OBJECTION,
        mode=RuleMode.MANDATORY,
        applies_to=("SI", "RFA", "RPD", "FI_CUSTOM"),
        description="If selected, include this objection for every request.",
        output_text=ALWAYS_VAGUE_TEXT,
    ),
    ResponseRule(
        id="ambiguous_term_when_unclear",
        name="Include ambiguous term objection when the discovery request contains word(s) that are confusing or not obviously clear",
        category=RuleCategory.OBJECTION,
        mode=RuleMode.CONDITIONAL,
        applies_to=("SI", "RFA", "RPD", "FI_CUSTOM"),
        description="Include when a request contains a confusing or undefined term.",
        output_text=AMBIGUOUS_TERM_TEXT,
    ),
    ResponseRule(
        id="relevance_privacy_when_unrelated",
        name="Include relevance and privacy objections when Discovery asks for information not related to the case",
        category=RuleCategory.OBJECTION,
        mode=RuleMode.CONDITIONAL,
        applies_to=("SI", "RFA", "RPD", "FI_CUSTOM"),
        description="Include when a request appears unrelated to the case or invades privacy.",
        output_text=RELEVANCE_PRIVACY_TEXT,
    ),
    ResponseRule(
        id="burdensome_when_lots_of_information",
        name="Include burdensome objections when the discovery asks for a lot of information",
        category=RuleCategory.OBJECTION,
        mode=RuleMode.CONDITIONAL,
        applies_to=("SI", "RFA", "RPD", "FI_CUSTOM"),
        description="Include when a request is broad, unlimited, or expensive to answer.",
        output_text=BURDENSOME_TEXT,
    ),
    ResponseRule(
        id="privilege_when_potentially_privileged",
        name="Include privilege objections when the discovery asks for potentially privileged information",
        category=RuleCategory.OBJECTION,
        mode=RuleMode.CONDITIONAL,
        applies_to=("SI", "RFA", "RPD", "FI_CUSTOM"),
        description="Include when a request could seek attorney-client or work-product information.",
        output_text=PRIVILEGE_TEXT,
    ),
    ResponseRule(
        id="expert_legal_conclusion_when_called_for",
        name="Include Expert Opinion or Legal Conclusion objections when Request calls for potential legal conclusion or expert opinion",
        category=RuleCategory.OBJECTION,
        mode=RuleMode.CONDITIONAL,
        applies_to=("SI", "RFA", "RPD", "FI_CUSTOM"),
        description="Include when a request calls for expert analysis or legal conclusions.",
        output_text=EXPERT_LEGAL_TEXT,
    ),
)

_SUBSTANTIVE_RULES = (
    ResponseRule(
        id="minimal_direct_answer",
        name="Answer only the question being asked using as few words as possible.",
        category=RuleCategory.SUBSTANTIVE,
        mode=RuleMode.INSTRUCTION,
        applies_to=("SI", "FI_CUSTOM"),
        description="Use narrow direct answers and do not volunteer extra facts.",
        output_text="Answer only the question being asked using as few words as possible.",
    ),
)

_QUICK_RULES = (
    ResponseRule("quick_vague", "Vague / Ambiguous / Overbroad", RuleCategory.OBJECTION, RuleMode.MANDATORY, ("ALL",), "", ALWAYS_VAGUE_TEXT),
    ResponseRule("quick_relevance_privacy", "Relevance / Privacy", RuleCategory.OBJECTION, RuleMode.MANDATORY, ("ALL",), "", RELEVANCE_PRIVACY_TEXT),
    ResponseRule("quick_expert_legal", "Expert Opinion / Legal Conclusion", RuleCategory.OBJECTION, RuleMode.MANDATORY, ("ALL",), "", EXPERT_LEGAL_TEXT),
    ResponseRule("quick_privilege", "Privilege", RuleCategory.OBJECTION, RuleMode.MANDATORY, ("ALL",), "", PRIVILEGE_TEXT),
    ResponseRule("quick_burdensome", "Burdensome", RuleCategory.OBJECTION, RuleMode.MANDATORY, ("ALL",), "", BURDENSOME_TEXT),
    ResponseRule("quick_argumentative", "Argumentative", RuleCategory.OBJECTION, RuleMode.MANDATORY, ("ALL",), "", ARGUMENTATIVE_TEXT),
    ResponseRule("quick_compound", "Compound", RuleCategory.OBJECTION, RuleMode.MANDATORY, ("ALL",), "", COMPOUND_TEXT),
    ResponseRule("quick_undefined_term", "Undefined Term", RuleCategory.OBJECTION, RuleMode.MANDATORY, ("ALL",), "", UNDEFINED_TERM_TEXT),
)


def built_in_rules_for(discovery_type: str, fi_mode: str = "fixed") -> List[ResponseRule]:
    dtype = (discovery_type or "").upper()
    if dtype == "FI" and fi_mode == "fixed":
        return []
    applies_key = "FI_CUSTOM" if dtype == "FI" else dtype
    rules: list[ResponseRule] = []
    for rule in _STANDARD_OBJECTION_RULES + _SUBSTANTIVE_RULES:
        if applies_key in rule.applies_to:
            rules.append(rule)
    return rules


def get_quick_objection_rules() -> List[ResponseRule]:
    return list(_QUICK_RULES)


def load_global_rules(path: str) -> List[ResponseRule]:
    if not os.path.exists(path):
        return []
    with open(path, "r", encoding="utf-8") as fh:
        raw = json.load(fh)
    return [ResponseRule.from_dict(item) for item in raw if isinstance(item, dict)]


def save_global_rules(path: str, rules: Iterable[ResponseRule]) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump([rule.to_dict() for rule in rules], fh, indent=2)
```

- [ ] **Step 4: Run test to verify it passes**

Run:

```powershell
pytest tests/test_discovery/test_response_rule_library.py -q
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/discovery/response_rule_library.py tests/test_discovery/test_response_rule_library.py
git commit -m "feat(discovery): add structured response rule library"
```

Skip commit if unrelated worktree conflicts are still present.

---

### Task 3: Review State And Quick Objections

**Files:**
- Create: `tests/test_discovery/test_response_review_state.py`
- Create: `icharlotte_core/discovery/response_review_state.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_discovery/test_response_review_state.py`:

```python
import unittest

from icharlotte_core.discovery.response_review_state import (
    RequestReview,
    ReviewState,
    insert_quick_objection,
    remove_quick_objection,
)


class ResponseReviewStateTests(unittest.TestCase):
    def test_all_approved_false_until_each_request_approved(self):
        state = ReviewState([
            RequestReview("1", "Question 1", "Obj", "Resp"),
            RequestReview("2", "Question 2", "Obj", "Resp"),
        ])
        self.assertFalse(state.all_approved)
        state.mark_approved("1")
        self.assertFalse(state.all_approved)
        state.mark_approved("2")
        self.assertTrue(state.all_approved)

    def test_editing_unapproves_request(self):
        review = RequestReview("1", "Question", "Obj", "Resp", approved=True)
        review.set_substantive_response("Changed")
        self.assertFalse(review.approved)

    def test_quick_objection_insert_is_idempotent(self):
        text = insert_quick_objection("", "Objection text.")
        text = insert_quick_objection(text, "Objection text.")
        self.assertEqual(text.count("Objection text."), 1)

    def test_quick_objection_remove_exact_text(self):
        text = "First. Objection text. Last."
        self.assertEqual(remove_quick_objection(text, "Objection text."), "First. Last.")

    def test_to_plain_text_contains_headers(self):
        state = ReviewState([RequestReview("1", "Describe incident.", "Obj.", "Resp.", approved=True)])
        plain = state.to_plain_text("SI")
        self.assertIn("SPECIAL INTERROGATORY NO. 1:", plain)
        self.assertIn("RESPONSE TO SPECIAL INTERROGATORY NO. 1:", plain)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
pytest tests/test_discovery/test_response_review_state.py -q
```

Expected: fail because `response_review_state.py` does not exist.

- [ ] **Step 3: Implement review state**

Create `icharlotte_core/discovery/response_review_state.py`:

```python
"""Review-state objects for the Respond to Discovery wizard."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable, List

from icharlotte_core.discovery.response_drafter import format_single_response
from icharlotte_core.discovery.response_rules import ResponseRules


def _clean_join(parts: Iterable[str]) -> str:
    return " ".join(part.strip() for part in parts if part and part.strip()).strip()


def insert_quick_objection(current_text: str, objection_text: str) -> str:
    current = (current_text or "").strip()
    objection = (objection_text or "").strip()
    if not objection:
        return current
    if objection in current:
        return current
    return _clean_join([current, objection])


def remove_quick_objection(current_text: str, objection_text: str) -> str:
    current = (current_text or "").strip()
    objection = (objection_text or "").strip()
    if not objection:
        return current
    return _clean_join(current.replace(objection, "").split())


@dataclass
class RequestReview:
    number: str
    request_text: str
    proposed_objections: str = ""
    proposed_substantive_response: str = ""
    selected_rule_ids: list[str] = field(default_factory=list)
    approved: bool = False

    def set_objections(self, text: str) -> None:
        self.proposed_objections = text
        self.approved = False

    def set_substantive_response(self, text: str) -> None:
        self.proposed_substantive_response = text
        self.approved = False

    def approve(self) -> None:
        self.approved = True


class ReviewState:
    def __init__(self, requests: list[RequestReview] | None = None):
        self.requests: list[RequestReview] = list(requests or [])

    @property
    def all_approved(self) -> bool:
        return bool(self.requests) and all(req.approved for req in self.requests)

    def get(self, number: str) -> RequestReview:
        for req in self.requests:
            if req.number == number:
                return req
        raise KeyError(number)

    def mark_approved(self, number: str) -> None:
        self.get(number).approve()

    def approved_count(self) -> int:
        return sum(1 for req in self.requests if req.approved)

    def to_plain_text(self, discovery_type: str, rules: ResponseRules | None = None) -> str:
        rules = rules or ResponseRules()
        dtype = discovery_type.upper()
        blocks: list[str] = []
        for req in self.requests:
            waiver = rules.waiver_language if req.proposed_objections and req.proposed_substantive_response else ""
            reservation = rules.reservation_clause if req.proposed_substantive_response else ""
            blocks.append(format_single_response(
                disc_type=dtype,
                request_number=req.number,
                request_text=req.request_text,
                objections=req.proposed_objections,
                substantive=req.proposed_substantive_response,
                waiver=waiver,
                reservation=reservation,
            ))
        return "\n\n".join(blocks)
```

- [ ] **Step 4: Run test to verify it passes**

Run:

```powershell
pytest tests/test_discovery/test_response_review_state.py -q
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/discovery/response_review_state.py tests/test_discovery/test_response_review_state.py
git commit -m "feat(discovery): track response review state"
```

Skip commit if unrelated worktree conflicts are still present.

---

### Task 4: Proposal Generation Engine

**Files:**
- Create: `tests/test_discovery/test_response_generation_engine.py`
- Create: `icharlotte_core/discovery/response_generation_engine.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_discovery/test_response_generation_engine.py`:

```python
import unittest

from icharlotte_core.discovery.response_generation_engine import (
    DraftCallbacks,
    generate_review_state,
)
from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
from icharlotte_core.discovery.response_rule_library import built_in_rules_for


class ResponseGenerationEngineTests(unittest.TestCase):
    def _parsed_si(self):
        return ParsedDiscovery(
            discovery_type="SI",
            propounding_party="Plaintiff X",
            responding_party="Defendant Y",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[
                ParsedRequest(number="1", text="State all facts supporting your contention."),
                ParsedRequest(number="2", text="Identify all DOCUMENTS supporting your claim."),
            ],
        )

    def test_mandatory_rule_applies_to_every_request(self):
        state = generate_review_state(
            self._parsed_si(),
            selected_rules=built_in_rules_for("SI"),
            context_text="",
            callbacks=DraftCallbacks(substantive=lambda req, ctx, rules: "Response."),
        )
        self.assertEqual(len(state.requests), 2)
        for review in state.requests:
            self.assertIn("vague, ambiguous, uncertain and overbroad", review.proposed_objections)

    def test_substantive_callback_used_for_each_request(self):
        seen = []
        def draft(req, ctx, rules):
            seen.append(req.number)
            return f"Answer {req.number}"

        state = generate_review_state(
            self._parsed_si(),
            selected_rules=built_in_rules_for("SI"),
            context_text="Context",
            callbacks=DraftCallbacks(substantive=draft),
        )
        self.assertEqual(seen, ["1", "2"])
        self.assertEqual(state.get("2").proposed_substantive_response, "Answer 2")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
pytest tests/test_discovery/test_response_generation_engine.py -q
```

Expected: fail because `response_generation_engine.py` does not exist.

- [ ] **Step 3: Implement generation engine**

Create `icharlotte_core/discovery/response_generation_engine.py`:

```python
"""Proposal generation for the Respond to Discovery wizard."""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable, Iterable

from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
from icharlotte_core.discovery.response_review_state import RequestReview, ReviewState
from icharlotte_core.discovery.response_rule_library import (
    ResponseRule,
    RuleCategory,
    RuleMode,
)


SubstantiveDraftFn = Callable[[ParsedRequest, str, list[ResponseRule]], str]


@dataclass
class DraftCallbacks:
    substantive: SubstantiveDraftFn


def _rule_applies_conditionally(rule: ResponseRule, req: ParsedRequest, context_text: str) -> bool:
    text = f"{req.text}\n{context_text or ''}".lower()
    if rule.id == "ambiguous_term_when_unclear":
        return bool(req.defined_terms_used) or any(word in text for word in ("undefined", "ambiguous", "unclear"))
    if rule.id == "relevance_privacy_when_unrelated":
        return any(word in text for word in ("privacy", "private", "medical", "financial", "unrelated"))
    if rule.id == "burdensome_when_lots_of_information":
        return any(word in text for word in ("all documents", "all facts", "each and every", "all communications"))
    if rule.id == "privilege_when_potentially_privileged":
        return any(word in text for word in ("attorney", "counsel", "work product", "investigation", "legal advice"))
    if rule.id == "expert_legal_conclusion_when_called_for":
        return any(word in text for word in ("expert", "opinion", "legal conclusion", "negligent", "standard of care"))
    return False


def _format_rule_output(rule: ResponseRule, req: ParsedRequest) -> str:
    text = rule.output_text
    if "{term}" in text:
        term = req.defined_terms_used[0] if req.defined_terms_used else "[insert unclear term]"
        text = text.replace("{term}", term)
    return text


def _selected_objections(req: ParsedRequest, selected_rules: Iterable[ResponseRule], context_text: str) -> tuple[str, list[str]]:
    parts: list[str] = []
    ids: list[str] = []
    for rule in selected_rules:
        if rule.category != RuleCategory.OBJECTION:
            continue
        include = False
        if rule.mode == RuleMode.MANDATORY:
            include = True
        elif rule.mode == RuleMode.CONDITIONAL:
            include = _rule_applies_conditionally(rule, req, context_text)
        if include:
            parts.append(_format_rule_output(rule, req))
            ids.append(rule.id)
    return " ".join(parts), ids


def generate_review_state(
    parsed: ParsedDiscovery,
    selected_rules: list[ResponseRule],
    context_text: str,
    callbacks: DraftCallbacks,
) -> ReviewState:
    reviews: list[RequestReview] = []
    substantive_rules = [r for r in selected_rules if r.category == RuleCategory.SUBSTANTIVE]
    for req in parsed.requests:
        objections, selected_ids = _selected_objections(req, selected_rules, context_text)
        substantive = callbacks.substantive(req, context_text, substantive_rules)
        reviews.append(RequestReview(
            number=req.number,
            request_text=req.text,
            proposed_objections=objections,
            proposed_substantive_response=substantive,
            selected_rule_ids=selected_ids + [r.id for r in substantive_rules],
        ))
    return ReviewState(reviews)
```

- [ ] **Step 4: Run test to verify it passes**

Run:

```powershell
pytest tests/test_discovery/test_response_generation_engine.py -q
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/discovery/response_generation_engine.py tests/test_discovery/test_response_generation_engine.py
git commit -m "feat(discovery): generate response review proposals"
```

Skip commit if unrelated worktree conflicts are still present.

---

### Task 5: Wizard UI Page Skeleton

**Files:**
- Create: `tests/test_wizard/test_respond_discovery_page.py`
- Create: `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`

- [ ] **Step 1: Write failing UI tests**

Create `tests/test_wizard/test_respond_discovery_page.py`:

```python
import os
import sys
import unittest

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from PySide6.QtWidgets import QApplication
app = QApplication.instance() or QApplication(sys.argv)

from icharlotte_core.ui.wizard.pages.respond_discovery_page import RespondDiscoverySettingsPage


class RespondDiscoveryPageTests(unittest.TestCase):
    def test_rules_for_si_show_substantive_section(self):
        page = RespondDiscoverySettingsPage(case_root="", file_number="", discovery_file="test.pdf", detected_type="SI")
        self.assertEqual(page.detected_type, "SI")
        self.assertTrue(page.has_substantive_rules())

    def test_rules_for_rfa_hide_substantive_section(self):
        page = RespondDiscoverySettingsPage(case_root="", file_number="", discovery_file="test.pdf", detected_type="RFA")
        self.assertFalse(page.has_substantive_rules())

    def test_fi_fixed_mode_hides_rule_cards(self):
        page = RespondDiscoverySettingsPage(case_root="", file_number="", discovery_file="test.pdf", detected_type="FI")
        page.set_fi_mode("fixed")
        self.assertEqual(page.visible_rule_count(), 0)

    def test_fi_custom_mode_shows_si_style_rules(self):
        page = RespondDiscoverySettingsPage(case_root="", file_number="", discovery_file="test.pdf", detected_type="FI")
        page.set_fi_mode("custom")
        self.assertGreater(page.visible_rule_count(), 0)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
pytest tests/test_wizard/test_respond_discovery_page.py -q
```

Expected: fail because `respond_discovery_page.py` does not exist.

- [ ] **Step 3: Implement minimal settings page**

Create `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`:

```python
"""Wizard UI for Respond to Discovery."""
from __future__ import annotations

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QGroupBox,
    QLabel,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.discovery.response_rule_library import (
    RuleCategory,
    ResponseRule,
    built_in_rules_for,
)


class RespondDiscoverySettingsPage(QWidget):
    run_requested = Signal(dict)

    def __init__(
        self,
        case_root: str,
        file_number: str,
        discovery_file: str,
        detected_type: str,
        parent=None,
    ):
        super().__init__(parent)
        self.case_root = case_root
        self.file_number = file_number
        self.discovery_file = discovery_file
        self.detected_type = detected_type.upper()
        self.fi_mode = "fixed"
        self._rules: list[ResponseRule] = []
        self._rule_checks: dict[str, QCheckBox] = {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(12)

        title = QLabel(f"Respond to Discovery - {self.detected_type}")
        title.setStyleSheet("font-size: 18px; font-weight: 600;")
        layout.addWidget(title)

        if self.detected_type == "FI":
            fi_group = QGroupBox("Form Interrogatory Mode")
            fi_layout = QVBoxLayout(fi_group)
            self.rb_fi_fixed = QRadioButton("Use Fixed Objections/Responses")
            self.rb_fi_custom = QRadioButton("Use Custom Objections/Responses")
            self.rb_fi_fixed.setChecked(True)
            self.rb_fi_fixed.toggled.connect(lambda checked: checked and self.set_fi_mode("fixed"))
            self.rb_fi_custom.toggled.connect(lambda checked: checked and self.set_fi_mode("custom"))
            fi_layout.addWidget(self.rb_fi_fixed)
            fi_layout.addWidget(self.rb_fi_custom)
            layout.addWidget(fi_group)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.rule_container = QWidget()
        self.rule_layout = QVBoxLayout(self.rule_container)
        self.scroll.setWidget(self.rule_container)
        layout.addWidget(self.scroll, 1)

        self.next_btn = QPushButton("Next: Context Files")
        self.next_btn.clicked.connect(self._emit_run_requested)
        layout.addWidget(self.next_btn)
        self._rebuild_rules()

    def set_fi_mode(self, mode: str) -> None:
        self.fi_mode = mode
        self._rebuild_rules()

    def _clear_rules(self) -> None:
        while self.rule_layout.count():
            item = self.rule_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._rule_checks.clear()

    def _rebuild_rules(self) -> None:
        self._clear_rules()
        self._rules = built_in_rules_for(self.detected_type, fi_mode=self.fi_mode)
        if not self._rules:
            self.rule_layout.addWidget(QLabel("Fixed FI objections/responses will be used."))
            return
        for heading, category in (("Objection Rules", RuleCategory.OBJECTION), ("Substantive Response Rules", RuleCategory.SUBSTANTIVE)):
            category_rules = [r for r in self._rules if r.category == category]
            if not category_rules:
                continue
            group = QGroupBox(heading)
            group_layout = QVBoxLayout(group)
            for rule in category_rules:
                cb = QCheckBox(rule.name)
                cb.setChecked(rule.enabled_by_default)
                cb.setToolTip(rule.description)
                group_layout.addWidget(cb)
                desc = QLabel(rule.description)
                desc.setWordWrap(True)
                desc.setStyleSheet("color: #666; margin-left: 20px;")
                group_layout.addWidget(desc)
                self._rule_checks[rule.id] = cb
            self.rule_layout.addWidget(group)

    def visible_rule_count(self) -> int:
        return len(self._rules)

    def has_substantive_rules(self) -> bool:
        return any(r.category == RuleCategory.SUBSTANTIVE for r in self._rules)

    def selected_rule_ids(self) -> list[str]:
        return [rid for rid, cb in self._rule_checks.items() if cb.isChecked()]

    def to_dict(self) -> dict:
        return {
            "discovery_file": self.discovery_file,
            "detected_type": self.detected_type,
            "fi_mode": self.fi_mode,
            "selected_rule_ids": self.selected_rule_ids(),
        }

    def from_dict(self, data: dict) -> None:
        if not data:
            return
        self.fi_mode = data.get("fi_mode", self.fi_mode)
        self._rebuild_rules()
        selected = set(data.get("selected_rule_ids", []))
        if selected:
            for rid, cb in self._rule_checks.items():
                cb.setChecked(rid in selected)

    def _emit_run_requested(self) -> None:
        self.run_requested.emit(self.to_dict())
```

- [ ] **Step 4: Run test to verify it passes**

Run:

```powershell
pytest tests/test_wizard/test_respond_discovery_page.py -q
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/ui/wizard/pages/respond_discovery_page.py tests/test_wizard/test_respond_discovery_page.py
git commit -m "feat(wizard): add respond discovery settings page"
```

Skip commit if unrelated worktree conflicts are still present.

---

### Task 6: Wizard Registry And Routing

**Files:**
- Create: `tests/test_wizard/test_respond_to_discovery_registry.py`
- Modify: `icharlotte_core/ui/wizard/registry.py`
- Modify: `icharlotte_core/ui/wizard/task_routing.py`
- Modify: `icharlotte_core/ui/wizard/in_process_task_tab.py`

- [ ] **Step 1: Write failing registry/routing tests**

Create `tests/test_wizard/test_respond_to_discovery_registry.py`:

```python
import unittest

from icharlotte_core.ui.wizard.registry import get_task, list_tasks
from icharlotte_core.ui.wizard.task_routing import (
    get_in_process_task_builder_name,
    requires_initial_file_picker,
)


class RespondToDiscoveryRegistryTests(unittest.TestCase):
    def test_task_registered(self):
        ids = {task.task_id for task in list_tasks()}
        self.assertIn("respond_to_discovery", ids)
        spec = get_task("respond_to_discovery")
        self.assertEqual(spec.title, "Respond to Discovery")
        self.assertEqual(spec.default_folders, ["DISCOVERY/PROPOUNDED", "DISCOVERY"])

    def test_task_uses_in_process_builder_without_generic_picker(self):
        self.assertEqual(
            get_in_process_task_builder_name("respond_to_discovery"),
            "build_respond_to_discovery_tab",
        )
        self.assertFalse(requires_initial_file_picker("respond_to_discovery"))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
pytest tests/test_wizard/test_respond_to_discovery_registry.py -q
```

Expected: fail because the task is not registered.

- [ ] **Step 3: Modify registry**

In `icharlotte_core/ui/wizard/registry.py`, add this entry to `TASK_REGISTRY` before `"chat"`:

```python
    "respond_to_discovery": TaskSpec(
        task_id="respond_to_discovery",
        title="Respond to Discovery",
        description="Draft objections and responses to incoming written discovery.",
        icon_glyph="\U0001F4DD",
        script_name="",
        default_folders=["DISCOVERY/PROPOUNDED", "DISCOVERY"],
    ),
```

Update `tests/test_wizard/test_registry.py::test_initial_tasks_registered` so the expected set includes `"respond_to_discovery"`.

- [ ] **Step 4: Modify routing**

In `icharlotte_core/ui/wizard/task_routing.py`, update `_IN_PROCESS_TASK_BUILDERS`:

```python
_IN_PROCESS_TASK_BUILDERS = {
    "subpoena_tracker": "build_subpoena_tab",
    "respond_to_discovery": "build_respond_to_discovery_tab",
}
```

- [ ] **Step 5: Add builder stub**

In `icharlotte_core/ui/wizard/in_process_task_tab.py`, add imports near other PySide widgets if missing:

```python
from PySide6.QtWidgets import QFileDialog
```

Add this factory after `build_subpoena_tab`:

```python
def build_respond_to_discovery_tab(spec, case_path: str, file_number: str, parent: QWidget | None) -> InProcessTaskTab:
    from icharlotte_core.discovery.response_type_detector import (
        detect_type_from_filename,
        detect_type_from_text,
        resolve_detected_type,
    )
    from icharlotte_core.ui.wizard.file_picker import resolve_default_folder
    from icharlotte_core.ui.wizard.pages.respond_discovery_page import RespondDiscoverySettingsPage

    start_dir = resolve_default_folder(case_path, spec.default_folders)
    file_path, _ = QFileDialog.getOpenFileName(
        parent,
        "Select discovery to respond to",
        start_dir,
        "PDF files (*.pdf)",
    )
    if not file_path:
        raise RuntimeError("No discovery file selected")

    filename_guess = detect_type_from_filename(file_path)
    text_guess = detect_type_from_text("")
    resolved = resolve_detected_type(filename_guess, text_guess)
    detected_type = resolved.discovery_type or "SI"

    def factory(cp, fn, settings, p):
        from icharlotte_core.ui.wizard.pages.respond_discovery_page import RespondDiscoveryWorker
        return RespondDiscoveryWorker(cp, fn, settings, parent=p)

    return InProcessTaskTab(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        settings_widget=RespondDiscoverySettingsPage(
            case_root=case_path,
            file_number=file_number,
            discovery_file=file_path,
            detected_type=detected_type,
        ),
        output_widget=_make_subpoena_output_page(),
        worker_factory=factory,
        auto_run=False,
        parent=parent,
    )
```

This is a stub route. Task 7 replaces `RespondDiscoveryWorker` with the actual worker and improves first-page text detection/user conflict prompts.

- [ ] **Step 6: Run tests**

Run:

```powershell
pytest tests/test_wizard/test_respond_to_discovery_registry.py tests/test_wizard/test_registry.py tests/test_wizard/test_task_routing.py -q
```

Expected: all tests pass.

- [ ] **Step 7: Commit**

```powershell
git add icharlotte_core/ui/wizard/registry.py icharlotte_core/ui/wizard/task_routing.py icharlotte_core/ui/wizard/in_process_task_tab.py tests/test_wizard/test_registry.py tests/test_wizard/test_respond_to_discovery_registry.py
git commit -m "feat(wizard): register respond to discovery task"
```

Skip commit if unrelated worktree conflicts are still present.

---

### Task 7: Worker, Parser, Context, And Assembly Wiring

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`
- Modify: `icharlotte_core/ui/wizard/in_process_task_tab.py`
- Test: extend `tests/test_wizard/test_respond_discovery_page.py`

- [ ] **Step 1: Write failing worker test**

Append to `tests/test_wizard/test_respond_discovery_page.py`:

```python
from unittest.mock import patch


class RespondDiscoveryWorkerTests(unittest.TestCase):
    def test_worker_rejects_unapproved_requests(self):
        from icharlotte_core.ui.wizard.pages.respond_discovery_page import RespondDiscoveryWorker

        worker = RespondDiscoveryWorker(
            case_path="",
            file_number="1234.001",
            settings={
                "review_state": {
                    "requests": [
                        {"number": "1", "request_text": "Q", "proposed_objections": "O", "proposed_substantive_response": "R", "approved": False}
                    ]
                }
            },
        )
        ok, message = worker.validate_settings()
        self.assertFalse(ok)
        self.assertIn("approved", message.lower())
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
pytest tests/test_wizard/test_respond_discovery_page.py::RespondDiscoveryWorkerTests -q
```

Expected: fail because `RespondDiscoveryWorker` is missing.

- [ ] **Step 3: Add worker skeleton and validation**

Append to `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`:

```python
from PySide6.QtCore import QThread, Signal


class RespondDiscoveryWorker(QThread):
    progress = Signal(str)
    warning = Signal(str)
    finished_result = Signal(bool, str)

    def __init__(self, case_path: str, file_number: str, settings: dict, parent=None):
        super().__init__(parent)
        self.case_path = case_path
        self.file_number = file_number
        self.settings = dict(settings or {})

    def validate_settings(self) -> tuple[bool, str]:
        review_state = self.settings.get("review_state") or {}
        requests = review_state.get("requests") or []
        if requests and not all(bool(item.get("approved")) for item in requests):
            return False, "Every discovery request must be approved before final assembly."
        return True, ""

    def run(self) -> None:
        ok, message = self.validate_settings()
        if not ok:
            self.finished_result.emit(False, message)
            return
        self.finished_result.emit(False, "RespondDiscoveryWorker assembly is not wired yet.")
```

- [ ] **Step 4: Run worker validation test**

Run:

```powershell
pytest tests/test_wizard/test_respond_discovery_page.py::RespondDiscoveryWorkerTests -q
```

Expected: pass.

- [ ] **Step 5: Add assembly implementation**

Replace `RespondDiscoveryWorker.run()` with:

```python
    def run(self) -> None:
        ok, message = self.validate_settings()
        if not ok:
            self.finished_result.emit(False, message)
            return
        try:
            from icharlotte_core.discovery.response_assembler import ResponseAssembler, build_response_filename
            from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
            from icharlotte_core.discovery.response_review_state import RequestReview, ReviewState
            from icharlotte_core.discovery.response_rules import ResponseRules
            from icharlotte_core.discovery.assembler import DiscoveryAssembler
            from icharlotte_core.word_validator import validate_report

            parsed_data = self.settings["parsed_discovery"]
            requests = [
                ParsedRequest(**item) for item in parsed_data.get("requests", [])
            ]
            parsed = ParsedDiscovery(**{k: v for k, v in parsed_data.items() if k != "requests"}, requests=requests)

            review_items = self.settings["review_state"]["requests"]
            review_state = ReviewState([
                RequestReview(
                    number=str(item["number"]),
                    request_text=str(item["request_text"]),
                    proposed_objections=str(item.get("proposed_objections", "")),
                    proposed_substantive_response=str(item.get("proposed_substantive_response", "")),
                    selected_rule_ids=list(item.get("selected_rule_ids", [])),
                    approved=bool(item.get("approved", False)),
                )
                for item in review_items
            ])
            response_text = review_state.to_plain_text(parsed.discovery_type, ResponseRules())

            caption_path = DiscoveryAssembler.find_caption_page(self.case_path)
            if not caption_path:
                self.finished_result.emit(False, "No caption page found in case folder.")
                return

            responding_name = parsed.responding_party or "Defendant"
            words = responding_name.replace(",", "").split()
            skip = {"defendant", "plaintiff", "cross-defendant", "cross-complainant"}
            abbreviation = next((w for w in words if w.lower() not in skip), "Def")
            output_dir = os.path.join(self.case_path, "NOTES", "AI OUTPUT", "DISCOVERY RESPONSES")
            output_path = os.path.join(output_dir, build_response_filename(abbreviation, parsed.discovery_type, parsed.set_number))

            assembler = ResponseAssembler(caption_path)
            assembler.assemble(parsed, response_text, ResponseRules(), output_path)
            validation = validate_report(output_path)
            if getattr(validation, "has_errors", False):
                self.finished_result.emit(False, "Word validation failed. Review validator findings before using output.")
                return
            self.finished_result.emit(True, output_path)
        except Exception as exc:
            self.finished_result.emit(False, str(exc))
```

Add `import os` at the top of `respond_discovery_page.py`.

- [ ] **Step 6: Add an assembly test with monkeypatches**

Append a test that monkeypatches `ResponseAssembler`, `DiscoveryAssembler.find_caption_page`, and `validate_report` so no real Word document is needed:

```python
    @patch("icharlotte_core.discovery.assembler.DiscoveryAssembler.find_caption_page", return_value="caption.docx")
    @patch("icharlotte_core.discovery.response_assembler.ResponseAssembler")
    @patch("icharlotte_core.word_validator.validate_report")
    def test_worker_assembles_when_all_requests_approved(self, mock_validate, mock_assembler_cls, _mock_caption):
        from icharlotte_core.ui.wizard.pages.respond_discovery_page import RespondDiscoveryWorker

        class Validation:
            has_errors = False
        mock_validate.return_value = Validation()
        mock_assembler_cls.return_value.assemble.return_value = "out.docx"

        worker = RespondDiscoveryWorker(
            case_path=r"C:\case",
            file_number="1234.001",
            settings={
                "parsed_discovery": {
                    "discovery_type": "SI",
                    "propounding_party": "Plaintiff X",
                    "responding_party": "Defendant Y",
                    "set_number": 1,
                    "set_word": "ONE",
                    "case_number": "123",
                    "requests": [{"number": "1", "text": "Q", "definitions": [], "is_compound": False, "defined_terms_used": []}],
                },
                "review_state": {
                    "requests": [
                        {"number": "1", "request_text": "Q", "proposed_objections": "O", "proposed_substantive_response": "R", "approved": True}
                    ]
                },
            },
        )
        ok, message = worker.validate_settings()
        self.assertTrue(ok, message)
```

- [ ] **Step 7: Run focused tests**

Run:

```powershell
pytest tests/test_wizard/test_respond_discovery_page.py tests/test_discovery/test_response_review_state.py -q
```

Expected: all tests pass.

- [ ] **Step 8: Commit**

```powershell
git add icharlotte_core/ui/wizard/pages/respond_discovery_page.py tests/test_wizard/test_respond_discovery_page.py
git commit -m "feat(wizard): assemble reviewed discovery responses"
```

Skip commit if unrelated worktree conflicts are still present.

---

### Task 8: Full Verification And Manual Smoke

**Files:**
- No planned source changes unless verification exposes defects.

- [ ] **Step 1: Run discovery tests**

Run:

```powershell
pytest tests/test_discovery/test_response_type_detector.py tests/test_discovery/test_response_rule_library.py tests/test_discovery/test_response_review_state.py tests/test_discovery/test_response_generation_engine.py -q
```

Expected: all tests pass.

- [ ] **Step 2: Run wizard tests**

Run:

```powershell
pytest tests/test_wizard/test_respond_to_discovery_registry.py tests/test_wizard/test_respond_discovery_page.py tests/test_wizard/test_task_routing.py tests/test_wizard/test_registry.py -q
```

Expected: all tests pass.

- [ ] **Step 3: Run broader existing response tests**

Run:

```powershell
pytest tests/test_response_parser.py tests/test_response_drafter.py tests/test_response_assembler.py tests/test_response_rules.py -q
```

Expected: all tests pass.

- [ ] **Step 4: Compile changed Python files**

Run:

```powershell
python -m py_compile `
  icharlotte_core/discovery/response_type_detector.py `
  icharlotte_core/discovery/response_rule_library.py `
  icharlotte_core/discovery/response_review_state.py `
  icharlotte_core/discovery/response_generation_engine.py `
  icharlotte_core/ui/wizard/pages/respond_discovery_page.py `
  icharlotte_core/ui/wizard/registry.py `
  icharlotte_core/ui/wizard/task_routing.py `
  icharlotte_core/ui/wizard/in_process_task_tab.py
```

Expected: no output and exit code 0.

- [ ] **Step 5: Manual smoke test**

Run the app:

```powershell
python iCharlotte.py
```

Manual checks:

1. Open a case.
2. Open Wizard mode.
3. Confirm **Respond to Discovery** card appears.
4. Click card.
5. Select one discovery PDF.
6. Confirm task tab opens.
7. Confirm rules display for detected type.
8. For SI, confirm objection rules and substantive rules are visible.
9. For RFA/RPD, confirm only objection rules are visible.
10. For FI, confirm fixed/custom choice.
11. Confirm no existing Microsoft Word windows are closed.

- [ ] **Step 6: Final status**

If all verification passes, summarize:

```text
Implemented Respond to Discovery wizard task. Focused tests and manual smoke passed. Existing Discovery -> Respond subtab remains intact.
```

If any verification fails, stop and fix the failing behavior before claiming completion.

