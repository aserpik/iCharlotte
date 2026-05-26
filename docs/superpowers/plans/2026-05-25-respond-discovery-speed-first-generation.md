# Respond to Discovery Speed-First Generation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Improve Wizard Mode Respond to Discovery proposal generation so proposed objections and substantive responses are request-specific, context-aware, and still fast to review.

**Architecture:** Add a small deterministic context-index helper, extend the existing generation engine with a structured request-proposal contract, and map those proposals into the existing review state. The wizard UI stays mostly unchanged and only gains a lightweight warning label for requests that need closer attorney review.

**Tech Stack:** Python 3, PySide6, unittest/pytest, existing iCharlotte discovery modules, mocked LLM calls through `icharlotte_core.llm_config.call_llm`.

---

## File Structure

- Create: `icharlotte_core/discovery/response_context_index.py`
  - Owns context chunking, scoring, and request-specific packet formatting.
  - No Qt imports and no LLM calls.
- Modify: `icharlotte_core/discovery/response_review_state.py`
  - Adds backward-compatible `needs_review` and `review_reason` fields to `RequestReview`.
- Modify: `icharlotte_core/discovery/response_generation_engine.py`
  - Adds `StructuredProposal`, structured JSON prompt/parsing helpers, fallback handling, and proposal-to-review mapping.
  - Keeps existing `generate_review_state()` API compatible for current tests and callers.
- Modify: `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`
  - Replaces the broad combined substantive-response pass in the wizard worker with request-scoped structured proposals.
  - Shows warning text only when the current review item is flagged.
- Test: `tests/test_discovery/test_response_context_index.py`
  - Covers chunking/scoring/empty-packet behavior.
- Test: `tests/test_discovery/test_response_review_state.py`
  - Covers warning field persistence and backward-compatible loads.
- Test: `tests/test_discovery/test_response_generation_engine.py`
  - Covers structured proposal parsing, mandatory rules, conditional rule IDs, unknown IDs, fallback, RFA/RPD defaults, and mapping to review rows.
- Test: `tests/test_wizard/test_respond_discovery_page.py`
  - Covers review warning display and worker use of structured proposal generation with mocked LLM output.

---

### Task 1: Add Review Warning State

**Files:**
- Modify: `icharlotte_core/discovery/response_review_state.py`
- Test: `tests/test_discovery/test_response_review_state.py`

- [ ] **Step 1: Write failing persistence tests**

Add these tests to `ResponseReviewStateTests` in `tests/test_discovery/test_response_review_state.py`:

```python
    def test_review_warning_fields_round_trip(self):
        state = ReviewState(
            requests=[
                RequestReview(
                    number="1",
                    request_text="Identify witnesses.",
                    proposed_substantive_response="Unknown at this time.",
                    needs_review=True,
                    review_reason="No specific context found.",
                )
            ]
        )

        loaded = ReviewState.from_dict(state.to_dict())

        self.assertTrue(loaded.requests[0].needs_review)
        self.assertEqual(
            loaded.requests[0].review_reason,
            "No specific context found.",
        )

    def test_review_warning_fields_are_optional_for_old_state(self):
        loaded = RequestReview.from_dict(
            {
                "number": "1",
                "request_text": "Identify witnesses.",
                "proposed_objections": "",
                "proposed_substantive_response": "",
                "selected_rule_ids": [],
                "selected_quick_objection_ids": [],
                "approved": False,
            }
        )

        self.assertFalse(loaded.needs_review)
        self.assertEqual(loaded.review_reason, "")
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
python -m pytest tests/test_discovery/test_response_review_state.py -q
```

Expected: fail with `TypeError: RequestReview.__init__() got an unexpected keyword argument 'needs_review'` or `AttributeError` for missing fields.

- [ ] **Step 3: Add warning fields to `RequestReview`**

In `icharlotte_core/discovery/response_review_state.py`, update the dataclass:

```python
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
```

Update `to_dict()`:

```python
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
```

Update `from_dict()`:

```python
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
```

- [ ] **Step 4: Run review-state tests**

Run:

```powershell
python -m pytest tests/test_discovery/test_response_review_state.py -q
```

Expected: all tests in the file pass.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/discovery/response_review_state.py tests/test_discovery/test_response_review_state.py
git commit -m "feat(discovery): persist response review warnings"
```

---

### Task 2: Add Request-Scoped Context Indexing

**Files:**
- Create: `icharlotte_core/discovery/response_context_index.py`
- Test: `tests/test_discovery/test_response_context_index.py`

- [ ] **Step 1: Write failing context-index tests**

Create `tests/test_discovery/test_response_context_index.py`:

```python
import unittest

from icharlotte_core.discovery.response_context_index import (
    ContextChunk,
    build_context_chunks,
    format_context_packet,
    select_context_packet,
)
from icharlotte_core.discovery.response_parser import ParsedRequest


class ResponseContextIndexTests(unittest.TestCase):
    def test_build_context_chunks_splits_headings_and_paragraphs(self):
        chunks = build_context_chunks(
            {
                r"C:\case\status.txt": (
                    "Witnesses\n"
                    "John Smith saw the impact.\n\n"
                    "Damages\n"
                    "Plaintiff claims neck pain."
                )
            }
        )

        self.assertGreaterEqual(len(chunks), 2)
        self.assertEqual(chunks[0].source_path, r"C:\case\status.txt")
        self.assertEqual(chunks[0].sequence, 1)
        self.assertIn("John Smith", "\n".join(chunk.text for chunk in chunks))

    def test_select_context_packet_prefers_request_terms(self):
        chunks = [
            ContextChunk("status.txt", 1, "Witnesses\nJohn Smith saw the impact.", "Witnesses"),
            ContextChunk("status.txt", 2, "Damages\nPlaintiff claims neck pain.", "Damages"),
        ]
        request = ParsedRequest(number="1", text="Identify all witnesses to the INCIDENT.")

        selected = select_context_packet(request, chunks, max_chunks=1)

        self.assertEqual(len(selected), 1)
        self.assertIn("John Smith", selected[0].text)

    def test_select_context_packet_returns_empty_for_no_signal(self):
        chunks = [
            ContextChunk("status.txt", 1, "Billing notes only.", ""),
        ]
        request = ParsedRequest(number="1", text="Identify all witnesses.")

        selected = select_context_packet(request, chunks, max_chunks=3)

        self.assertEqual(selected, [])

    def test_format_context_packet_includes_source_labels(self):
        text = format_context_packet(
            [ContextChunk("status.txt", 2, "John Smith saw the impact.", "Witnesses")]
        )

        self.assertIn("[status.txt #2]", text)
        self.assertIn("John Smith saw the impact.", text)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
python -m pytest tests/test_discovery/test_response_context_index.py -q
```

Expected: fail with `ModuleNotFoundError: No module named 'icharlotte_core.discovery.response_context_index'`.

- [ ] **Step 3: Implement context indexing helper**

Create `icharlotte_core/discovery/response_context_index.py`:

```python
"""Request-scoped context selection for discovery response drafting."""
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Iterable, Mapping

from icharlotte_core.discovery.response_parser import ParsedRequest


@dataclass(frozen=True)
class ContextChunk:
    source_path: str
    sequence: int
    text: str
    heading: str = ""


_DISCOVERY_CUES = {
    "admit", "admission", "communications", "documents", "identify",
    "incident", "produce", "witness", "witnesses",
}

_LOW_SIGNAL_TERMS = {
    "all", "and", "any", "are", "each", "for", "from", "identify",
    "request", "response", "state", "that", "the", "this", "with", "you",
    "your",
}


def build_context_chunks(text_by_path: Mapping[str, str]) -> list[ContextChunk]:
    chunks: list[ContextChunk] = []
    for source_path, text in text_by_path.items():
        for sequence, chunk_text in enumerate(_split_text(text), start=1):
            heading = _detect_heading(chunk_text)
            chunks.append(
                ContextChunk(
                    source_path=source_path,
                    sequence=sequence,
                    text=chunk_text,
                    heading=heading,
                )
            )
    return chunks


def select_context_packet(
    request: ParsedRequest,
    chunks: Iterable[ContextChunk],
    max_chunks: int = 5,
    min_score: int = 2,
) -> list[ContextChunk]:
    scored = [
        (_score_chunk(request, chunk), chunk)
        for chunk in chunks
    ]
    scored = [(score, chunk) for score, chunk in scored if score >= min_score]
    scored.sort(key=lambda item: (-item[0], item[1].source_path, item[1].sequence))
    return [chunk for _score, chunk in scored[:max_chunks]]


def format_context_packet(chunks: Iterable[ContextChunk]) -> str:
    parts: list[str] = []
    for chunk in chunks:
        label = f"[{os.path.basename(chunk.source_path)} #{chunk.sequence}]"
        parts.append(f"{label}\n{chunk.text.strip()}")
    return "\n\n".join(parts)


def _split_text(text: str) -> list[str]:
    normalized = re.sub(r"\r\n?", "\n", text or "")
    raw_parts = re.split(r"\n\s*\n+", normalized)
    chunks = [re.sub(r"\s+", " ", part).strip() for part in raw_parts]
    return [chunk for chunk in chunks if len(chunk) >= 20]


def _detect_heading(text: str) -> str:
    first_line = (text or "").splitlines()[0].strip() if text else ""
    if 0 < len(first_line) <= 80 and not first_line.endswith("."):
        return first_line
    match = re.match(r"^([A-Z][A-Za-z /-]{2,60})\s+", text or "")
    return match.group(1).strip() if match else ""


def _score_chunk(request: ParsedRequest, chunk: ContextChunk) -> int:
    request_terms = _terms(request.text)
    chunk_text = f"{chunk.heading} {chunk.text}".lower()
    score = 0
    for term in request_terms:
        if term in chunk_text:
            score += 3 if term in _DISCOVERY_CUES else 2
    for term in getattr(request, "defined_terms_used", []) or []:
        if term.lower() in chunk_text:
            score += 4
    if re.search(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", request.text or "") and re.search(
        r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", chunk.text
    ):
        score += 2
    return score


def _terms(text: str) -> set[str]:
    words = {word.lower() for word in re.findall(r"[A-Za-z][A-Za-z0-9'-]{2,}", text or "")}
    return {word for word in words if word not in _LOW_SIGNAL_TERMS}
```

- [ ] **Step 4: Run context-index tests**

Run:

```powershell
python -m pytest tests/test_discovery/test_response_context_index.py -q
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/discovery/response_context_index.py tests/test_discovery/test_response_context_index.py
git commit -m "feat(discovery): add request scoped context index"
```

---

### Task 3: Add Structured Proposal Parsing And Mapping

**Files:**
- Modify: `icharlotte_core/discovery/response_generation_engine.py`
- Test: `tests/test_discovery/test_response_generation_engine.py`

- [ ] **Step 1: Write failing structured proposal tests**

Append these tests to `ResponseGenerationEngineTests` in `tests/test_discovery/test_response_generation_engine.py`:

```python
    def test_structured_proposal_ignores_model_objection_text(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
            apply_structured_proposal,
        )

        parsed = _parsed()
        rules = built_in_rules_for("SI")
        proposal = StructuredProposal(
            request_number="2",
            conditional_objection_rule_ids=["ambiguous_term_when_unclear"],
            applied_custom_rule_ids=[],
            applied_instruction_rule_ids=["minimal_direct_answer"],
            ambiguous_term="INCIDENT",
            proposed_objections="MODEL SHOULD NOT CONTROL OBJECTION TEXT",
            proposed_substantive_response="No additional facts known.",
            needs_review=True,
            review_reason="No specific context found.",
        )

        review = apply_structured_proposal(parsed.requests[1], parsed, rules, proposal)

        self.assertIn('"INCIDENT"', review.proposed_objections)
        self.assertNotIn("MODEL SHOULD NOT CONTROL", review.proposed_objections)
        self.assertEqual(review.proposed_substantive_response, "No additional facts known.")
        self.assertTrue(review.needs_review)
        self.assertEqual(review.review_reason, "No specific context found.")

    def test_structured_proposal_ignores_unknown_rule_ids(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
            apply_structured_proposal,
        )

        parsed = _parsed()
        rules = built_in_rules_for("SI")
        proposal = StructuredProposal(
            request_number="1",
            conditional_objection_rule_ids=["missing_rule"],
            applied_custom_rule_ids=[],
            applied_instruction_rule_ids=[],
            proposed_substantive_response="Unknown.",
        )

        review = apply_structured_proposal(parsed.requests[0], parsed, rules, proposal)

        self.assertEqual(review.proposed_objections, "")
        self.assertTrue(review.needs_review)
        self.assertIn("Unknown rule ID", review.review_reason)

    def test_parse_structured_proposal_json_extracts_schema(self):
        from icharlotte_core.discovery.response_generation_engine import (
            parse_structured_proposal_response,
        )

        proposal = parse_structured_proposal_response(
            """
            ```json
            {
              "request_number": "1",
              "conditional_objection_rule_ids": ["privilege_when_potentially_privileged"],
              "applied_custom_rule_ids": [],
              "applied_instruction_rule_ids": ["minimal_direct_answer"],
              "ambiguous_term": "",
              "proposed_objections": "ignored",
              "proposed_substantive_response": "No privileged documents will be produced.",
              "needs_review": true,
              "review_reason": "Possible privilege issue."
            }
            ```
            """
        )

        self.assertEqual(proposal.request_number, "1")
        self.assertEqual(
            proposal.conditional_objection_rule_ids,
            ["privilege_when_potentially_privileged"],
        )
        self.assertTrue(proposal.needs_review)
        self.assertEqual(proposal.review_reason, "Possible privilege issue.")

    def test_generate_review_state_can_use_structured_proposal_callback(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
            generate_review_state,
        )

        rules = built_in_rules_for("SI")

        def propose(request, parsed, context_text, selected_rules, response_rules):
            return StructuredProposal(
                request_number=request.number,
                conditional_objection_rule_ids=["privilege_when_potentially_privileged"]
                if request.number == "1"
                else [],
                applied_instruction_rule_ids=["minimal_direct_answer"],
                proposed_substantive_response=f"Response for {request.number}.",
                needs_review=request.number == "2",
                review_reason="No specific context found." if request.number == "2" else "",
            )

        state = generate_review_state(
            _parsed(),
            rules,
            context_text="case facts",
            callbacks=DraftCallbacks(structured_proposal=propose),
        )

        self.assertIn("privilege", state.requests[0].proposed_objections.lower())
        self.assertEqual(state.requests[0].proposed_substantive_response, "Response for 1.")
        self.assertTrue(state.requests[1].needs_review)
        self.assertEqual(state.requests[1].review_reason, "No specific context found.")
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
python -m pytest tests/test_discovery/test_response_generation_engine.py -q
```

Expected: fail with import errors for `StructuredProposal`, `apply_structured_proposal`, or `parse_structured_proposal_response`.

- [ ] **Step 3: Add structured proposal model and callback type**

In `icharlotte_core/discovery/response_generation_engine.py`, add imports:

```python
import json
```

Add after `ConditionalRuleDecision`:

```python
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


StructuredProposalCallback = Callable[
    [ParsedRequest, ParsedDiscovery, str, list[ResponseRule], ResponseRules],
    StructuredProposal,
]
```

Update `DraftCallbacks`:

```python
@dataclass
class DraftCallbacks:
    should_apply_rule: RuleDecisionCallback | None = None
    draft_substantive: DraftSubstantiveCallback | None = None
    structured_proposal: StructuredProposalCallback | None = None
```

- [ ] **Step 4: Route `generate_review_state()` through structured proposals when provided**

In the loop inside `generate_review_state()`, before the current objection loop, add:

```python
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
```

Leave the existing heuristic/callback behavior untouched after this branch.

- [ ] **Step 5: Add parser and proposal mapping helpers**

Append these helpers before `_decide_rule()`:

```python
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
        conditional_objection_rule_ids=_string_list(data.get("conditional_objection_rule_ids")),
        applied_custom_rule_ids=_string_list(data.get("applied_custom_rule_ids")),
        applied_instruction_rule_ids=_string_list(data.get("applied_instruction_rule_ids")),
        ambiguous_term=str(data.get("ambiguous_term", "")),
        proposed_objections=str(data.get("proposed_objections", "")),
        proposed_substantive_response=str(data.get("proposed_substantive_response", "")),
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

    unknown_ids = [rid for rid in requested_ids if rid not in rules_by_id]
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
                if rules_by_id.get(rid) and rules_by_id[rid].category == RuleCategory.SUBSTANTIVE
            ]
        )
        if rid in rules_by_id
    ]
    selected_rule_ids = list(dict.fromkeys([rid for rid in requested_ids + instruction_ids if rid in rules_by_id]))
    needs_review = proposal.needs_review or bool(unknown_ids)
    review_reason = proposal.review_reason.strip()
    if unknown_ids:
        unknown_message = "Unknown rule ID returned by model: " + ", ".join(sorted(set(unknown_ids)))
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
```

- [ ] **Step 6: Run generation-engine tests**

Run:

```powershell
python -m pytest tests/test_discovery/test_response_generation_engine.py -q
```

Expected: all tests pass.

- [ ] **Step 7: Commit**

```powershell
git add icharlotte_core/discovery/response_generation_engine.py tests/test_discovery/test_response_generation_engine.py
git commit -m "feat(discovery): map structured response proposals"
```

---

### Task 4: Add Structured Proposal Prompting And Fallbacks

**Files:**
- Modify: `icharlotte_core/discovery/response_generation_engine.py`
- Test: `tests/test_discovery/test_response_generation_engine.py`

- [ ] **Step 1: Write failing prompt and fallback tests**

Append these tests to `ResponseGenerationEngineTests`:

```python
    def test_build_structured_proposal_prompt_contains_schema_and_packet(self):
        from icharlotte_core.discovery.response_generation_engine import (
            build_structured_proposal_prompt,
        )

        parsed = _parsed("RFA")
        prompt = build_structured_proposal_prompt(
            parsed.requests[0],
            parsed,
            context_packet="[status.txt #1]\nNo witnesses identified.",
            selected_rules=built_in_rules_for("RFA"),
            response_rules=ResponseRules(),
        )

        self.assertIn("Return ONLY a JSON object", prompt)
        self.assertIn("conditional_objection_rule_ids", prompt)
        self.assertIn("[status.txt #1]", prompt)
        self.assertIn("admit only when", prompt.lower())

    def test_fallback_proposal_marks_empty_context_for_review(self):
        from icharlotte_core.discovery.response_generation_engine import (
            build_fallback_structured_proposal,
        )

        parsed = _parsed("SI")
        proposal = build_fallback_structured_proposal(
            parsed.requests[0],
            parsed,
            context_packet="",
        )

        self.assertEqual(proposal.request_number, "1")
        self.assertTrue(proposal.needs_review)
        self.assertIn("No specific context found", proposal.review_reason)

    def test_rfa_fallback_uses_insufficient_information(self):
        from icharlotte_core.discovery.response_generation_engine import (
            build_fallback_structured_proposal,
        )

        parsed = _parsed("RFA")
        proposal = build_fallback_structured_proposal(
            parsed.requests[0],
            parsed,
            context_packet="",
        )

        self.assertIn("insufficient to enable Responding Party to admit", proposal.proposed_substantive_response)
        self.assertTrue(proposal.needs_review)

    def test_rpd_fallback_uses_unable_to_comply_when_context_empty(self):
        from icharlotte_core.discovery.response_generation_engine import (
            build_fallback_structured_proposal,
        )

        parsed = _parsed("RPD")
        proposal = build_fallback_structured_proposal(
            parsed.requests[0],
            parsed,
            context_packet="",
        )

        self.assertIn("unable to comply", proposal.proposed_substantive_response.lower())
        self.assertTrue(proposal.needs_review)
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
python -m pytest tests/test_discovery/test_response_generation_engine.py -q
```

Expected: fail with import errors for `build_structured_proposal_prompt` and `build_fallback_structured_proposal`.

- [ ] **Step 3: Add prompt builder and posture text**

In `icharlotte_core/discovery/response_generation_engine.py`, add:

```python
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
    return "\n".join(
        [
            "You are a California civil litigation attorney drafting proposed discovery responses.",
            "Return ONLY a JSON object matching the requested schema.",
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
            "- Do not include objections in proposed_substantive_response.",
            "- Do not include waiver or reservation language.",
            "- Do not invent facts not supported by the context packet.",
            "- If context is weak, use cautious default language and set needs_review true.",
            "- Select only objection rule IDs that apply.",
            "- Do not draft new objection text.",
            "",
            "JSON schema:",
            "{",
            '  "request_number": "string",',
            '  "conditional_objection_rule_ids": ["rule_id"],',
            '  "applied_custom_rule_ids": ["rule_id"],',
            '  "applied_instruction_rule_ids": ["rule_id"],',
            '  "ambiguous_term": "string",',
            '  "proposed_objections": "ignored by application",',
            '  "proposed_substantive_response": "string",',
            '  "needs_review": false,',
            '  "review_reason": "string"',
            "}",
        ]
    )
```

Add helpers:

```python
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
            "Say will comply only when context indicates responsive non-privileged "
            "documents exist or will be produced. Otherwise use unable-to-comply "
            "or cautious default language."
        )
    if dtype == "FI":
        return "Draft only the substantive factual response. Answer narrowly."
    return (
        "Answer only the question being asked using as few words as possible. "
        "Do not volunteer extra facts."
    )
```

- [ ] **Step 4: Add fallback proposal builder**

Add:

```python
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
        substantive = (
            "Upon a diligent search and reasonable inquiry, Responding Party is "
            "unable to comply with this request at this time because responsive "
            "documents, if they exist, have not been identified in the available context."
        )
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
```

- [ ] **Step 5: Run generation-engine tests**

Run:

```powershell
python -m pytest tests/test_discovery/test_response_generation_engine.py -q
```

Expected: all tests pass.

- [ ] **Step 6: Commit**

```powershell
git add icharlotte_core/discovery/response_generation_engine.py tests/test_discovery/test_response_generation_engine.py
git commit -m "feat(discovery): build structured proposal prompts"
```

---

### Task 5: Wire Structured Proposals Into The Wizard Worker

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`
- Test: `tests/test_wizard/test_respond_discovery_page.py`

- [ ] **Step 1: Write failing worker helper tests**

Update imports in `tests/test_wizard/test_respond_discovery_page.py`:

```python
from icharlotte_core.discovery.response_rules import ResponseRules
from icharlotte_core.ui.wizard.pages.respond_discovery_page import (
    RespondDiscoverySettingsPage,
    RespondDiscoveryWorker,
    _build_structured_proposal_map,
    _draft_substantive_response_map,
    _normalize_and_filter_parsed_discovery,
    load_respond_response_rules,
)
```

Add this test to `RespondDiscoverySettingsPageTests`:

```python
    @patch("icharlotte_core.llm_config.call_llm")
    def test_structured_proposal_map_uses_request_specific_context(self, mock_call):
        parsed = ParsedDiscovery(
            discovery_type="SI",
            propounding_party="Plaintiff Smith",
            responding_party="Defendant Jones",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[
                ParsedRequest(number="1", text="Identify all witnesses."),
                ParsedRequest(number="2", text="Identify all documents."),
            ],
        )
        mock_call.side_effect = [
            '{"request_number":"1","conditional_objection_rule_ids":[],"applied_custom_rule_ids":[],"applied_instruction_rule_ids":["minimal_direct_answer"],"ambiguous_term":"","proposed_objections":"","proposed_substantive_response":"John Smith.","needs_review":false,"review_reason":""}',
            '{"request_number":"2","conditional_objection_rule_ids":[],"applied_custom_rule_ids":[],"applied_instruction_rule_ids":["minimal_direct_answer"],"ambiguous_term":"","proposed_objections":"","proposed_substantive_response":"Photos and repair invoices.","needs_review":false,"review_reason":""}',
        ]

        proposals = _build_structured_proposal_map(
            parsed=parsed,
            selected_rules=[],
            context_text_by_path={
                "status.txt": "Witnesses\nJohn Smith saw the collision.\n\nDocuments\nPhotos and repair invoices exist."
            },
            response_rules=ResponseRules(),
        )

        self.assertEqual(proposals["1"].proposed_substantive_response, "John Smith.")
        self.assertEqual(proposals["2"].proposed_substantive_response, "Photos and repair invoices.")
        first_prompt = mock_call.call_args_list[0].args[0]
        second_prompt = mock_call.call_args_list[1].args[0]
        self.assertIn("John Smith saw the collision", first_prompt)
        self.assertIn("Photos and repair invoices exist", second_prompt)
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
python -m pytest tests/test_wizard/test_respond_discovery_page.py::RespondDiscoverySettingsPageTests::test_structured_proposal_map_uses_request_specific_context -q
```

Expected: fail with import error for `_build_structured_proposal_map`.

- [ ] **Step 3: Add imports to wizard page**

In `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`, update the generation-engine imports:

```python
from icharlotte_core.discovery.response_generation_engine import (
    DraftCallbacks,
    StructuredProposal,
    build_fallback_structured_proposal,
    build_structured_proposal_prompt,
    generate_review_state,
    parse_structured_proposal_response,
)
```

Add context-index imports:

```python
from icharlotte_core.discovery.response_context_index import (
    build_context_chunks,
    format_context_packet,
    select_context_packet,
)
```

- [ ] **Step 4: Add structured proposal map helper**

Replace `_draft_substantive_response_map()` usage in new code, but keep the old function in place for existing tests. Add this helper above `_draft_substantive_response_map()`:

```python
def _build_structured_proposal_map(
    parsed: ParsedDiscovery,
    selected_rules: list[ResponseRule],
    context_text_by_path: dict[str, str],
    response_rules: ResponseRules | None = None,
) -> dict[str, StructuredProposal]:
    from icharlotte_core.llm_config import call_llm

    response_rules = response_rules or ResponseRules()
    chunks = build_context_chunks(context_text_by_path)
    proposals: dict[str, StructuredProposal] = {}
    for req in parsed.requests:
        packet_chunks = select_context_packet(req, chunks)
        context_packet = format_context_packet(packet_chunks)
        prompt = build_structured_proposal_prompt(
            req,
            parsed,
            context_packet,
            selected_rules,
            response_rules,
        )
        try:
            raw = call_llm(prompt, "", task_type="general", agent_id="agent_sum_disc")
            proposals[req.number] = parse_structured_proposal_response(raw or "")
        except Exception:
            proposals[req.number] = build_fallback_structured_proposal(
                req,
                parsed,
                context_packet,
            )
    return proposals
```

- [ ] **Step 5: Update `RespondDiscoveryProposalWorker.run()`**

In `run()`, replace the `context_text` and `_draft_substantive_response_map()` block with:

```python
            self.progress.emit("Reading context files...")
            context_text_by_path = {
                path: read_document_text(path)
                for path in self.context_files
                if os.path.isfile(path)
            }

            self.progress.emit("Drafting proposed responses...")
            response_rules = load_respond_response_rules(self.file_number)
            proposal_map = _build_structured_proposal_map(
                parsed=parsed,
                selected_rules=self.selected_rules,
                context_text_by_path=context_text_by_path,
                response_rules=response_rules,
            )

            def propose(req, _parsed, _context, _rules, _response_rules):
                return proposal_map.get(
                    req.number,
                    build_fallback_structured_proposal(req, _parsed, ""),
                )

            review_state = generate_review_state(
                parsed,
                self.selected_rules,
                context_text="",
                response_rules=response_rules,
                callbacks=DraftCallbacks(structured_proposal=propose),
                fi_mode=self.fi_mode,
            )
```

Remove the old local `response_map` and `draft()` callback from this worker path.

- [ ] **Step 6: Run targeted worker helper test**

Run:

```powershell
python -m pytest tests/test_wizard/test_respond_discovery_page.py::RespondDiscoverySettingsPageTests::test_structured_proposal_map_uses_request_specific_context -q
```

Expected: test passes.

- [ ] **Step 7: Run wizard page tests**

Run:

```powershell
python -m pytest tests/test_wizard/test_respond_discovery_page.py -q
```

Expected: all tests in the file pass.

- [ ] **Step 8: Commit**

```powershell
git add icharlotte_core/ui/wizard/pages/respond_discovery_page.py tests/test_wizard/test_respond_discovery_page.py
git commit -m "feat(wizard): draft discovery responses per request"
```

---

### Task 6: Show Request Warning In Review UI

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`
- Test: `tests/test_wizard/test_respond_discovery_page.py`

- [ ] **Step 1: Write failing UI warning test**

Add this test to `RespondDiscoverySettingsPageTests`:

```python
    def test_review_warning_label_shows_current_request_reason(self):
        parsed = ParsedDiscovery(
            discovery_type="SI",
            propounding_party="Plaintiff Smith",
            responding_party="Defendant Jones",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[ParsedRequest(number="1", text="Identify witnesses.")],
        )
        state = ReviewState(
            [
                RequestReview(
                    number="1",
                    request_text="Identify witnesses.",
                    proposed_substantive_response="Unknown.",
                    needs_review=True,
                    review_reason="No specific context found.",
                )
            ]
        )

        page = RespondDiscoverySettingsPage(
            case_root="",
            file_number="1234.001",
            discovery_file=r"C:\case\srogg.pdf",
            detected_type="SI",
            parsed_discovery=parsed,
            review_state=state,
        )

        self.assertFalse(page.review_warning_label.isHidden())
        self.assertIn("No specific context found", page.review_warning_label.text())
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
python -m pytest tests/test_wizard/test_respond_discovery_page.py::RespondDiscoverySettingsPageTests::test_review_warning_label_shows_current_request_reason -q
```

Expected: fail with `AttributeError: 'RespondDiscoverySettingsPage' object has no attribute 'review_warning_label'`.

- [ ] **Step 3: Add warning label to `_build_review_widget()`**

In `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`, immediately after `self.request_text` is added to the review layout, add:

```python
        self.review_warning_label = QLabel("")
        self.review_warning_label.setWordWrap(True)
        self.review_warning_label.setStyleSheet("color: #8a5a00; font-weight: 600;")
        self.review_warning_label.hide()
        layout.addWidget(self.review_warning_label)
```

- [ ] **Step 4: Load warning state in `_load_current_review()`**

In `_load_current_review()`, after setting `self.response_edit`, add:

```python
        if review.needs_review and review.review_reason:
            self.review_warning_label.setText(f"Needs review: {review.review_reason}")
            self.review_warning_label.show()
        else:
            self.review_warning_label.setText("")
            self.review_warning_label.hide()
```

- [ ] **Step 5: Run targeted UI warning test**

Run:

```powershell
python -m pytest tests/test_wizard/test_respond_discovery_page.py::RespondDiscoverySettingsPageTests::test_review_warning_label_shows_current_request_reason -q
```

Expected: test passes.

- [ ] **Step 6: Run wizard page tests**

Run:

```powershell
python -m pytest tests/test_wizard/test_respond_discovery_page.py -q
```

Expected: all tests in the file pass.

- [ ] **Step 7: Commit**

```powershell
git add icharlotte_core/ui/wizard/pages/respond_discovery_page.py tests/test_wizard/test_respond_discovery_page.py
git commit -m "feat(wizard): show discovery response review warnings"
```

---

### Task 7: Regression And Integration Verification

**Files:**
- No new source files.
- Verify current touched test suite.

- [ ] **Step 1: Run discovery unit tests touched by the feature**

Run:

```powershell
python -m pytest tests/test_discovery/test_response_context_index.py tests/test_discovery/test_response_generation_engine.py tests/test_discovery/test_response_review_state.py -q
```

Expected: all selected tests pass.

- [ ] **Step 2: Run wizard response tests**

Run:

```powershell
python -m pytest tests/test_wizard/test_respond_discovery_page.py tests/test_wizard/test_respond_to_discovery_registry.py -q
```

Expected: all selected tests pass.

- [ ] **Step 3: Run legacy Respond tab smoke tests**

Run:

```powershell
python -m pytest tests/test_respond_tab.py -q
```

Expected: all tests pass, confirming the advanced Respond subtab still imports and its basic behavior is unchanged.

- [ ] **Step 4: Run compile check on touched modules**

Run:

```powershell
python -m py_compile icharlotte_core/discovery/response_context_index.py icharlotte_core/discovery/response_generation_engine.py icharlotte_core/discovery/response_review_state.py icharlotte_core/ui/wizard/pages/respond_discovery_page.py
```

Expected: command exits with code 0 and prints no syntax errors.

- [ ] **Step 5: Inspect diff for unintended broad changes**

Run:

```powershell
git diff --stat HEAD
git diff -- icharlotte_core/discovery/response_context_index.py icharlotte_core/discovery/response_generation_engine.py icharlotte_core/discovery/response_review_state.py icharlotte_core/ui/wizard/pages/respond_discovery_page.py tests/test_discovery/test_response_context_index.py tests/test_discovery/test_response_generation_engine.py tests/test_discovery/test_response_review_state.py tests/test_wizard/test_respond_discovery_page.py
```

Expected: diff only includes files named in this plan and no unrelated cleanup.

- [ ] **Step 6: Commit final verification note if any test-only adjustment was needed**

If Step 1 through Step 5 required a small test fix, commit that fix:

```powershell
git add tests/test_discovery/test_response_context_index.py tests/test_discovery/test_response_generation_engine.py tests/test_discovery/test_response_review_state.py tests/test_wizard/test_respond_discovery_page.py
git commit -m "test(wizard): cover speed first discovery response generation"
```

If no files changed after Step 5, do not create an empty commit.
