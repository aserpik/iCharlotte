# Respond to Discovery — Parallel Generation + Streaming Review Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Parallelize the Respond to Discovery wizard's proposal generation (cap 8 concurrent), stream results into the review screen as they arrive, drop the dead `proposed_objections` field from the LLM prompt, decouple parse from draft work, and add a per-request Regenerate button.

**Architecture:** Replace the single serial `RespondDiscoveryProposalWorker` with three new units: `DiscoveryParseWorker` (one-shot QThread for parsing), `ProposalTask` (QRunnable run in a thread pool), and `ProposalCoordinator` (QObject that fans tasks out and emits results). The review screen opens immediately with pending placeholders; each `proposal_ready` signal patches one row in-place. User edits are never auto-overwritten — a banner appears when fresh drafts collide with user typing.

**Tech Stack:** Python 3.12+, PySide6 (`QThread`, `QThreadPool`, `QRunnable`, `QObject`, `Signal`), pytest + pytestqt, unittest.mock. The discovery domain code in `icharlotte_core/discovery/` and the wizard page in `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`.

**Spec:** [docs/superpowers/specs/2026-05-26-respond-discovery-parallel-streaming-design.md](../specs/2026-05-26-respond-discovery-parallel-streaming-design.md)

---

## File Structure

**Files to create:**
- `icharlotte_core/discovery/proposal_coordinator.py` — `ProposalTask` (QRunnable), `WorkerSignals`, `ProposalCoordinator` (QObject).
- `icharlotte_core/discovery/discovery_parse_worker.py` — `DiscoveryParseWorker` (QThread).
- `tests/test_discovery/test_proposal_coordinator.py` — coordinator + task unit tests.
- `tests/test_discovery/test_discovery_parse_worker.py` — parse worker unit tests.

**Files to modify:**
- `icharlotte_core/discovery/response_review_state.py` — add `is_pending`, `pending_replacement` fields on `RequestReview`.
- `icharlotte_core/discovery/response_generation_engine.py` — drop dead field from prompt; add `_build_pending_review_state` and `_apply_proposal_to_review_state`.
- `icharlotte_core/llm_config.py` — `LLMConfig.discovery_response_max_concurrent()` accessor.
- `icharlotte_core/ui/wizard/pages/respond_discovery_page.py` — rewire to use parse worker + coordinator; remove `RespondDiscoveryProposalWorker`, `_build_structured_proposal_map`, `_draft_substantive_response_map`, `_generate_review_state_from_proposals`; add status indicators, pending state, Regenerate row, conflict banner.
- `tests/test_discovery/test_response_review_state.py` — cover new fields.
- `tests/test_discovery/test_response_generation_engine.py` — cover prompt change and new helpers.
- `tests/test_wizard/test_respond_discovery_page.py` — cover streaming behavior, edit conflict, regenerate.

**Why these boundaries:** The new `proposal_coordinator.py` is a self-contained concurrency primitive — it depends on the discovery domain models but not on Qt UI widgets, so it's testable in isolation with a synchronous fake task factory. `discovery_parse_worker.py` is tiny but lives in its own file so it can be imported by the page without dragging the coordinator's thread-pool state. The page stays the orchestrator: it owns the worker and coordinator instances, wires their signals to UI updates, and is the only place Qt widgets meet discovery state.

---

## Task 1: Drop dead `proposed_objections` field from the proposal prompt

**Why first:** Mechanical change with no API surface impact, immediately reduces wasted tokens, and proves the test harness runs before we touch the harder concurrency code.

**Files:**
- Modify: `icharlotte_core/discovery/response_generation_engine.py:250-309`
- Test: `tests/test_discovery/test_response_generation_engine.py`

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_discovery/test_response_generation_engine.py`:

```python
from icharlotte_core.discovery.response_generation_engine import (
    build_structured_proposal_prompt,
    parse_structured_proposal_response,
)
from icharlotte_core.discovery.response_rules import ResponseRules


class StructuredProposalPromptTests(unittest.TestCase):
    def _build_prompt(self):
        parsed = _parsed()
        request = parsed.requests[0]
        rules = built_in_rules_for("SI")
        return build_structured_proposal_prompt(
            request=request,
            parsed=parsed,
            context_packet="",
            selected_rules=rules,
            response_rules=ResponseRules(),
        )

    def test_prompt_omits_proposed_objections_from_schema(self):
        prompt = self._build_prompt()
        # The dead "proposed_objections" field used to live inside the JSON
        # schema block. We still mention objection RULES in the schema,
        # but the OBJECTIONS-TEXT field is gone.
        self.assertNotIn('"proposed_objections"', prompt)

    def test_prompt_explicitly_forbids_drafting_objection_text(self):
        prompt = self._build_prompt()
        self.assertIn("Do not draft objection text.", prompt)

    def test_parse_handles_payload_without_proposed_objections(self):
        # Old clients won't send the field. The parser must tolerate that.
        proposal = parse_structured_proposal_response(
            '{"request_number": "1", "proposed_substantive_response": "ok"}'
        )
        self.assertEqual(proposal.request_number, "1")
        self.assertEqual(proposal.proposed_substantive_response, "ok")
        self.assertEqual(proposal.proposed_objections, "")

    def test_parse_still_tolerates_legacy_proposed_objections_field(self):
        # Old persisted JSON may still carry it; we just ignore the value.
        proposal = parse_structured_proposal_response(
            '{"request_number": "1", "proposed_objections": "legacy", '
            '"proposed_substantive_response": "ok"}'
        )
        self.assertEqual(proposal.proposed_objections, "legacy")
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_discovery/test_response_generation_engine.py::StructuredProposalPromptTests -v`

Expected: `test_prompt_omits_proposed_objections_from_schema` and `test_prompt_explicitly_forbids_drafting_objection_text` FAIL. The two parse tests should PASS (the parser already handles a missing field).

- [ ] **Step 3: Update the prompt builder**

In `icharlotte_core/discovery/response_generation_engine.py`, replace the body of `build_structured_proposal_prompt` (lines 250–309). The change: remove `'  "proposed_objections": "ignored by application",'` from the JSON schema block, and add an explicit instruction line above the schema.

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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_discovery/test_response_generation_engine.py -v`

Expected: All tests in the file PASS, including pre-existing ones.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/discovery/response_generation_engine.py tests/test_discovery/test_response_generation_engine.py
git commit -m @'
discovery: drop dead proposed_objections field from proposal prompt

The application has always discarded the LLM's proposed_objections
output in favor of canned rule text. Remove the field from the JSON
schema in build_structured_proposal_prompt and add an explicit
instruction telling the model not to draft objection text.

Saves ~80 prompt tokens per request and removes a confusing
instruction. parse_structured_proposal_response still tolerates the
field if it shows up in legacy payloads.
'@
```

---

## Task 2: Add `is_pending` and `pending_replacement` fields to `RequestReview`

**Why now:** Pure dataclass change. Foundation for streaming review. Touches one file with full serialization symmetry. Easy to TDD.

**Files:**
- Modify: `icharlotte_core/discovery/response_review_state.py:10-51`
- Test: `tests/test_discovery/test_response_review_state.py`

- [ ] **Step 1: Inspect existing test file**

Read: `tests/test_discovery/test_response_review_state.py` to find a class to extend or a convention to match. (You'll append a new test class at the end.)

- [ ] **Step 2: Write the failing tests**

Append to `tests/test_discovery/test_response_review_state.py`:

```python
class RequestReviewPendingFieldsTests(unittest.TestCase):
    def test_default_request_review_is_not_pending(self):
        from icharlotte_core.discovery.response_review_state import RequestReview
        review = RequestReview(number="1", request_text="x")
        self.assertFalse(review.is_pending)
        self.assertIsNone(review.pending_replacement)

    def test_is_pending_round_trips_through_to_from_dict(self):
        from icharlotte_core.discovery.response_review_state import RequestReview
        review = RequestReview(number="1", request_text="x", is_pending=True)
        round_tripped = RequestReview.from_dict(review.to_dict())
        self.assertTrue(round_tripped.is_pending)

    def test_pending_replacement_is_not_serialized(self):
        # pending_replacement is session-only — it MUST NOT round-trip.
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
        )
        from icharlotte_core.discovery.response_review_state import RequestReview
        proposal = StructuredProposal(request_number="1")
        review = RequestReview(
            number="1", request_text="x", pending_replacement=proposal,
        )
        data = review.to_dict()
        self.assertNotIn("pending_replacement", data)
        round_tripped = RequestReview.from_dict(data)
        self.assertIsNone(round_tripped.pending_replacement)

    def test_legacy_dict_without_is_pending_loads_as_false(self):
        from icharlotte_core.discovery.response_review_state import RequestReview
        legacy = {
            "number": "1",
            "request_text": "x",
            "proposed_objections": "",
            "proposed_substantive_response": "",
            "selected_rule_ids": [],
            "selected_quick_objection_ids": [],
            "approved": False,
            "needs_review": False,
            "review_reason": "",
        }
        review = RequestReview.from_dict(legacy)
        self.assertFalse(review.is_pending)
        self.assertIsNone(review.pending_replacement)

    def test_all_approved_false_when_any_request_is_pending(self):
        from icharlotte_core.discovery.response_review_state import (
            RequestReview, ReviewState,
        )
        state = ReviewState(requests=[
            RequestReview(number="1", request_text="x", approved=True),
            RequestReview(
                number="2", request_text="y", approved=True, is_pending=True,
            ),
        ])
        self.assertFalse(state.all_approved())
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `python -m pytest tests/test_discovery/test_response_review_state.py::RequestReviewPendingFieldsTests -v`

Expected: All five tests FAIL — `is_pending` and `pending_replacement` are not yet defined.

- [ ] **Step 4: Add fields and update serialization**

In `icharlotte_core/discovery/response_review_state.py`, edit the `RequestReview` dataclass and its `to_dict`/`from_dict`, plus `ReviewState.all_approved`. The new imports at top:

```python
from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Iterable

from icharlotte_core.discovery.response_rule_library import ResponseRule

if TYPE_CHECKING:
    from icharlotte_core.discovery.response_generation_engine import (
        StructuredProposal,
    )
```

Replace `RequestReview` and `ReviewState.all_approved` with:

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
    is_pending: bool = False
    # Session-only: holds a freshly-generated proposal when the user has
    # already edited this request. Never serialized.
    pending_replacement: "StructuredProposal | None" = field(
        default=None, repr=False, compare=False,
    )

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
            "is_pending": bool(self.is_pending),
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
            is_pending=bool(data.get("is_pending", False)),
        )
```

And update `ReviewState.all_approved`:

```python
    def all_approved(self) -> bool:
        return all(
            item.approved and not item.is_pending
            for item in self.requests
        )
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest tests/test_discovery/test_response_review_state.py -v`

Expected: All tests in the file PASS.

- [ ] **Step 6: Commit**

```powershell
git add icharlotte_core/discovery/response_review_state.py tests/test_discovery/test_response_review_state.py
git commit -m @'
discovery: add is_pending and pending_replacement to RequestReview

is_pending marks rows whose proposal is still being generated by the
coordinator. Serialized so that resumed wizard sessions know which
rows to enqueue.

pending_replacement holds a freshly-generated proposal when the user
has already edited the row. Session-only — never serialized.

ReviewState.all_approved now requires every row to also be
non-pending, so the Finalize button can't be enabled while
generation is in flight.
'@
```

---

## Task 3: `LLMConfig.discovery_response_max_concurrent()` accessor

**Why now:** Tiny config helper. Lets the coordinator (next task) read its cap without coupling to JSON shape. Independent commit.

**Files:**
- Modify: `icharlotte_core/llm_config.py` (add one method to `LLMConfig`)
- Test: `tests/test_llm_config.py` (extend if it exists; else create)

- [ ] **Step 1: Check whether a test file exists**

Run: `python -c "import os; print(os.path.exists('tests/test_llm_config.py'))"`

If `True`, append. If `False`, create with the structure below.

- [ ] **Step 2: Write the failing test**

Append (or create) `tests/test_llm_config.py`:

```python
import unittest
from unittest.mock import patch


class LLMConfigDiscoveryResponseTests(unittest.TestCase):
    def _reset_singleton(self):
        from icharlotte_core.llm_config import LLMConfig
        LLMConfig._instance = None

    def setUp(self):
        self._reset_singleton()

    def tearDown(self):
        self._reset_singleton()

    def test_default_cap_is_eight_when_key_absent(self):
        from icharlotte_core.llm_config import LLMConfig
        config = LLMConfig()
        with patch.object(config, "_config", {}):
            self.assertEqual(config.discovery_response_max_concurrent(), 8)

    def test_reads_override_from_config_dict(self):
        from icharlotte_core.llm_config import LLMConfig
        config = LLMConfig()
        with patch.object(
            config,
            "_config",
            {"discovery_response": {"max_concurrent_proposals": 3}},
        ):
            self.assertEqual(config.discovery_response_max_concurrent(), 3)

    def test_invalid_value_falls_back_to_default(self):
        from icharlotte_core.llm_config import LLMConfig
        config = LLMConfig()
        with patch.object(
            config,
            "_config",
            {"discovery_response": {"max_concurrent_proposals": "bad"}},
        ):
            self.assertEqual(config.discovery_response_max_concurrent(), 8)

    def test_zero_or_negative_clamps_to_one(self):
        from icharlotte_core.llm_config import LLMConfig
        config = LLMConfig()
        with patch.object(
            config,
            "_config",
            {"discovery_response": {"max_concurrent_proposals": 0}},
        ):
            self.assertEqual(config.discovery_response_max_concurrent(), 1)
        with patch.object(
            config,
            "_config",
            {"discovery_response": {"max_concurrent_proposals": -5}},
        ):
            self.assertEqual(config.discovery_response_max_concurrent(), 1)
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `python -m pytest tests/test_llm_config.py::LLMConfigDiscoveryResponseTests -v`

Expected: All four tests FAIL with `AttributeError: 'LLMConfig' object has no attribute 'discovery_response_max_concurrent'`.

- [ ] **Step 4: Add the accessor method**

In `icharlotte_core/llm_config.py`, inside the `LLMConfig` class (after `_load_config` is fine), add:

```python
    def discovery_response_max_concurrent(self) -> int:
        """Return the max concurrent proposal tasks for the discovery wizard.

        Reads ``discovery_response.max_concurrent_proposals`` from the loaded
        config, defaulting to 8. Clamps non-positive values to 1.
        """
        section = self._config.get("discovery_response") if isinstance(
            self._config, dict
        ) else None
        default = 8
        if not isinstance(section, dict):
            return default
        raw = section.get("max_concurrent_proposals", default)
        try:
            value = int(raw)
        except (TypeError, ValueError):
            return default
        return max(1, value)
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest tests/test_llm_config.py::LLMConfigDiscoveryResponseTests -v`

Expected: All four tests PASS.

- [ ] **Step 6: Commit**

```powershell
git add icharlotte_core/llm_config.py tests/test_llm_config.py
git commit -m @'
llm_config: add discovery_response_max_concurrent() accessor

Reads the optional config key
`discovery_response.max_concurrent_proposals` (default 8, clamps to
>=1). Lets the new ProposalCoordinator size its thread pool without
coupling to the JSON shape.
'@
```

---

## Task 4: `ProposalTask` + `WorkerSignals`

**Why now:** Pure unit, no Qt event loop required for the body logic. Tests inject a fake `call_llm`. Sets up the contract the coordinator depends on.

**Files:**
- Create: `icharlotte_core/discovery/proposal_coordinator.py`
- Test: `tests/test_discovery/test_proposal_coordinator.py`

- [ ] **Step 1: Write the failing tests (just the task half — coordinator tests come in Task 5)**

Create `tests/test_discovery/test_proposal_coordinator.py`:

```python
import unittest
from unittest.mock import MagicMock

import pytest

pytest.importorskip("pytestqt")

from PySide6.QtCore import QObject, Signal
from PySide6.QtWidgets import QApplication

from icharlotte_core.discovery.proposal_coordinator import (
    ProposalTask,
    WorkerSignals,
)
from icharlotte_core.discovery.response_generation_engine import StructuredProposal
from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
from icharlotte_core.discovery.response_rules import ResponseRules


def _parsed():
    return ParsedDiscovery(
        discovery_type="SI",
        propounding_party="P",
        responding_party="D",
        set_number=1,
        set_word="ONE",
        case_number="1",
        requests=[ParsedRequest(number="1", text="Identify witnesses.")],
    )


class ProposalTaskTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._app = QApplication.instance() or QApplication([])

    def _make_signals(self):
        return WorkerSignals()

    def _make_task(self, signals, *, call_llm, override_instruction=""):
        parsed = _parsed()
        return ProposalTask(
            signals=signals,
            request=parsed.requests[0],
            parsed=parsed,
            context_packet="Some context.",
            selected_rules=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
            call_llm=call_llm,
            override_instruction=override_instruction,
        )

    def test_emits_proposal_ready_with_parsed_proposal(self):
        emitted = []
        signals = self._make_signals()
        signals.proposal_ready.connect(lambda num, prop: emitted.append((num, prop)))

        call_llm = MagicMock(
            return_value='{"request_number": "1", "proposed_substantive_response": "OK"}'
        )
        task = self._make_task(signals, call_llm=call_llm)
        task.run()

        self.assertEqual(len(emitted), 1)
        number, proposal = emitted[0]
        self.assertEqual(number, "1")
        self.assertEqual(proposal.proposed_substantive_response, "OK")
        call_llm.assert_called_once()

    def test_appends_override_instruction_to_prompt(self):
        signals = self._make_signals()
        signals.proposal_ready.connect(lambda *a: None)
        call_llm = MagicMock(
            return_value='{"request_number": "1", "proposed_substantive_response": "X"}'
        )
        task = self._make_task(
            signals,
            call_llm=call_llm,
            override_instruction="Lean harder on privilege.",
        )
        task.run()
        prompt = call_llm.call_args[0][0]
        self.assertIn("ADDITIONAL INSTRUCTIONS:", prompt)
        self.assertIn("Lean harder on privilege.", prompt)

    def test_repair_prompt_invoked_when_initial_response_is_unparseable(self):
        signals = self._make_signals()
        emitted = []
        signals.proposal_ready.connect(lambda num, prop: emitted.append((num, prop)))

        # First call returns garbage; second (repair) call returns valid JSON.
        call_llm = MagicMock(
            side_effect=[
                "not json at all",
                '{"request_number": "1", "proposed_substantive_response": "REPAIRED"}',
            ]
        )
        task = self._make_task(signals, call_llm=call_llm)
        task.run()

        self.assertEqual(call_llm.call_count, 2)
        repair_prompt = call_llm.call_args_list[1][0][0]
        self.assertIn("Repair this structured discovery proposal JSON.", repair_prompt)
        self.assertEqual(emitted[0][1].proposed_substantive_response, "REPAIRED")

    def test_fallback_used_when_repair_also_fails(self):
        signals = self._make_signals()
        emitted = []
        signals.proposal_ready.connect(lambda num, prop: emitted.append((num, prop)))

        call_llm = MagicMock(side_effect=["junk", "still junk"])
        task = self._make_task(signals, call_llm=call_llm)
        task.run()

        # One emission, fallback proposal, needs_review True.
        self.assertEqual(len(emitted), 1)
        _, proposal = emitted[0]
        self.assertTrue(proposal.needs_review)
        self.assertIn("could not be parsed", proposal.review_reason.lower())

    def test_fallback_used_when_call_llm_raises(self):
        signals = self._make_signals()
        emitted = []
        signals.proposal_ready.connect(lambda num, prop: emitted.append((num, prop)))

        call_llm = MagicMock(side_effect=RuntimeError("rate limited"))
        task = self._make_task(signals, call_llm=call_llm)
        task.run()

        self.assertEqual(len(emitted), 1)
        _, proposal = emitted[0]
        self.assertTrue(proposal.needs_review)
        self.assertIn("rate limited", proposal.review_reason)

    def test_emits_needs_review_when_context_packet_empty(self):
        signals = self._make_signals()
        emitted = []
        signals.proposal_ready.connect(lambda num, prop: emitted.append((num, prop)))

        call_llm = MagicMock(
            return_value='{"request_number": "1", "proposed_substantive_response": "OK"}'
        )
        parsed = _parsed()
        task = ProposalTask(
            signals=signals,
            request=parsed.requests[0],
            parsed=parsed,
            context_packet="",  # empty
            selected_rules=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
            call_llm=call_llm,
        )
        task.run()

        _, proposal = emitted[0]
        self.assertTrue(proposal.needs_review)
        self.assertIn("No specific context found", proposal.review_reason)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_discovery/test_proposal_coordinator.py::ProposalTaskTests -v`

Expected: All tests FAIL with `ModuleNotFoundError: No module named 'icharlotte_core.discovery.proposal_coordinator'`.

- [ ] **Step 3: Create the module with `WorkerSignals` and `ProposalTask`**

Create `icharlotte_core/discovery/proposal_coordinator.py`:

```python
"""Concurrent proposal generation for the Respond to Discovery wizard.

This module provides three primitives:

- ``WorkerSignals``: small QObject that owns ``proposal_ready`` and
  ``proposal_failed`` signals (QRunnable can't own signals directly).
- ``ProposalTask``: a QRunnable that builds the structured-proposal
  prompt for one request, calls the LLM (with a one-shot repair retry),
  and emits the parsed StructuredProposal.
- ``ProposalCoordinator``: a QObject that owns a QThreadPool, fans
  ProposalTasks out for a parsed discovery, and emits per-request and
  aggregate progress signals.

The coordinator is decoupled from Qt widgets so it can be unit-tested
with a synchronous fake ``task_factory``.
"""
from __future__ import annotations

from dataclasses import replace
from typing import Callable

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal

from icharlotte_core.discovery.response_context_index import (
    ContextChunk,
    format_context_packet,
    select_context_packet,
)
from icharlotte_core.discovery.response_drafter import (
    detect_inapplicable_fi,
    get_fi_fixed_response,
)
from icharlotte_core.discovery.response_generation_engine import (
    StructuredProposal,
    build_fallback_structured_proposal,
    build_structured_proposal_prompt,
    parse_structured_proposal_response,
)
from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
from icharlotte_core.discovery.response_rule_library import ResponseRule
from icharlotte_core.discovery.response_rules import ResponseRules
from icharlotte_core.discovery.response_type_detector import normalize_discovery_type


CallLLMFn = Callable[..., str]


class WorkerSignals(QObject):
    """Signal carrier for ProposalTask.

    QRunnable can't define signals directly; we route them through this
    helper. One WorkerSignals instance can be shared by many tasks (the
    coordinator does that).
    """

    proposal_ready = Signal(str, object)  # request_number, StructuredProposal
    proposal_failed = Signal(str, str)    # request_number, reason


class ProposalTask(QRunnable):
    """Generate one structured proposal for one parsed discovery request."""

    def __init__(
        self,
        signals: WorkerSignals,
        request: ParsedRequest,
        parsed: ParsedDiscovery,
        context_packet: str,
        selected_rules: list[ResponseRule],
        response_rules: ResponseRules,
        fi_mode: str = "custom",
        call_llm: CallLLMFn | None = None,
        override_instruction: str = "",
    ):
        super().__init__()
        self.signals = signals
        self.request = request
        self.parsed = parsed
        self.context_packet = context_packet or ""
        self.selected_rules = list(selected_rules or [])
        self.response_rules = response_rules or ResponseRules()
        self.fi_mode = fi_mode
        self._call_llm = call_llm
        self.override_instruction = (override_instruction or "").strip()
        self.setAutoDelete(True)

    def run(self) -> None:  # noqa: D401 - QRunnable contract
        try:
            proposal = self._generate_proposal()
        except Exception as exc:  # pragma: no cover - defensive
            proposal = build_fallback_structured_proposal(
                self.request,
                self.parsed,
                self.context_packet,
            )
            proposal = replace(
                proposal,
                needs_review=True,
                review_reason=f"Task raised: {str(exc)[:200]}",
            )
        self.signals.proposal_ready.emit(self.request.number, proposal)

    def _generate_proposal(self) -> StructuredProposal:
        call_llm = self._resolve_call_llm()
        prompt = build_structured_proposal_prompt(
            self.request,
            self.parsed,
            self.context_packet,
            self.selected_rules,
            self.response_rules,
        )
        if self.override_instruction:
            prompt = f"{prompt}\n\nADDITIONAL INSTRUCTIONS:\n{self.override_instruction}"

        try:
            raw = call_llm(prompt, "", task_type="general", agent_id="agent_sum_disc")
        except Exception as exc:
            fallback = build_fallback_structured_proposal(
                self.request, self.parsed, self.context_packet,
            )
            return replace(
                fallback,
                needs_review=True,
                review_reason=f"LLM call failed: {str(exc)[:200]}",
            )

        try:
            proposal = parse_structured_proposal_response(raw or "")
        except Exception:
            repaired = self._call_repair(call_llm, raw or "")
            if repaired is None:
                fallback = build_fallback_structured_proposal(
                    self.request, self.parsed, self.context_packet,
                )
                return replace(
                    fallback,
                    needs_review=True,
                    review_reason="Model response could not be parsed.",
                )
            proposal = repaired

        return self._ensure_context_warning(proposal)

    def _call_repair(self, call_llm: CallLLMFn, raw_text: str) -> StructuredProposal | None:
        repair_prompt = (
            "Repair this structured discovery proposal JSON. "
            "Return ONLY one valid JSON object with the same schema. "
            f"The request_number must be {self.request.number}.\n\n"
            f"INVALID RESPONSE:\n{raw_text}"
        )
        try:
            repaired_raw = call_llm(
                repair_prompt, "", task_type="general", agent_id="agent_sum_disc",
            )
            return parse_structured_proposal_response(repaired_raw or "")
        except Exception:
            return None

    def _ensure_context_warning(self, proposal: StructuredProposal) -> StructuredProposal:
        if self.context_packet.strip():
            return proposal
        reason = (proposal.review_reason or "").strip()
        missing_context = "No specific context found."
        if missing_context.lower() not in reason.lower():
            reason = f"{reason} {missing_context}".strip()
        return replace(proposal, needs_review=True, review_reason=reason)

    def _resolve_call_llm(self) -> CallLLMFn:
        if self._call_llm is not None:
            return self._call_llm
        from icharlotte_core.llm_config import call_llm
        return call_llm


def _should_skip_request(
    request: ParsedRequest,
    parsed: ParsedDiscovery,
    response_rules: ResponseRules,
    fi_mode: str,
) -> bool:
    """Mirror response_generation_engine semantics for FI fixed mode."""
    if normalize_discovery_type(parsed.discovery_type) != "FI" or fi_mode != "fixed":
        return False
    if detect_inapplicable_fi(request.number):
        return True
    return get_fi_fixed_response(request.number, response_rules) is not None


class ProposalCoordinator(QObject):
    """Fan out one ProposalTask per request through a QThreadPool."""

    proposal_ready = Signal(str, object)  # request_number, StructuredProposal
    progress = Signal(int, int)           # completed, total
    all_done = Signal()

    def __init__(
        self,
        max_concurrent: int = 8,
        task_factory: Callable[..., QRunnable] | None = None,
        parent: QObject | None = None,
    ):
        super().__init__(parent)
        self._pool = QThreadPool(self)
        self._pool.setMaxThreadCount(max(1, int(max_concurrent)))
        self._task_factory = task_factory
        self._signals = WorkerSignals()
        self._signals.proposal_ready.connect(self._on_proposal_ready)
        self._in_flight: set[str] = set()
        self._completed: set[str] = set()
        self._total: int = 0
        self._cancelled = False

    # Public API ------------------------------------------------------

    def start(
        self,
        parsed: ParsedDiscovery,
        selected_rules: list[ResponseRule],
        context_chunks: list[ContextChunk],
        response_rules: ResponseRules,
        fi_mode: str = "custom",
    ) -> list[str]:
        """Enqueue a task for every non-skipped request. Returns enqueued numbers."""
        self._cancelled = False
        self._completed.clear()
        self._in_flight.clear()
        enqueued: list[str] = []
        for req in parsed.requests:
            if _should_skip_request(req, parsed, response_rules, fi_mode):
                continue
            packet_chunks = select_context_packet(req, context_chunks)
            context_packet = format_context_packet(packet_chunks)
            task = self._build_task(
                request=req,
                parsed=parsed,
                context_packet=context_packet,
                selected_rules=selected_rules,
                response_rules=response_rules,
                fi_mode=fi_mode,
            )
            self._in_flight.add(req.number)
            enqueued.append(req.number)
            self._pool.start(task)
        self._total = len(enqueued)
        if self._total == 0:
            self.all_done.emit()
        return enqueued

    def regenerate(
        self,
        request_number: str,
        parsed: ParsedDiscovery,
        selected_rules: list[ResponseRule],
        context_chunks: list[ContextChunk],
        response_rules: ResponseRules,
        fi_mode: str = "custom",
        override_instruction: str = "",
    ) -> bool:
        """Re-enqueue one request. No-op if already in flight. Returns True if queued."""
        if request_number in self._in_flight:
            return False
        target = next(
            (req for req in parsed.requests if req.number == request_number), None,
        )
        if target is None:
            return False
        packet_chunks = select_context_packet(target, context_chunks)
        context_packet = format_context_packet(packet_chunks)
        task = self._build_task(
            request=target,
            parsed=parsed,
            context_packet=context_packet,
            selected_rules=selected_rules,
            response_rules=response_rules,
            fi_mode=fi_mode,
            override_instruction=override_instruction,
        )
        self._completed.discard(request_number)
        self._in_flight.add(request_number)
        self._total = max(self._total, len(self._completed) + len(self._in_flight))
        self._pool.start(task)
        return True

    def cancel(self) -> None:
        self._cancelled = True
        self._in_flight.clear()

    def is_done(self) -> bool:
        return not self._in_flight and self._total > 0

    # Internal --------------------------------------------------------

    def _build_task(
        self,
        request: ParsedRequest,
        parsed: ParsedDiscovery,
        context_packet: str,
        selected_rules: list[ResponseRule],
        response_rules: ResponseRules,
        fi_mode: str,
        override_instruction: str = "",
    ) -> QRunnable:
        if self._task_factory is not None:
            return self._task_factory(
                signals=self._signals,
                request=request,
                parsed=parsed,
                context_packet=context_packet,
                selected_rules=selected_rules,
                response_rules=response_rules,
                fi_mode=fi_mode,
                override_instruction=override_instruction,
            )
        return ProposalTask(
            signals=self._signals,
            request=request,
            parsed=parsed,
            context_packet=context_packet,
            selected_rules=selected_rules,
            response_rules=response_rules,
            fi_mode=fi_mode,
            override_instruction=override_instruction,
        )

    def _on_proposal_ready(self, request_number: str, proposal: StructuredProposal) -> None:
        if self._cancelled:
            return
        if request_number in self._in_flight:
            self._in_flight.remove(request_number)
        self._completed.add(request_number)
        self.proposal_ready.emit(request_number, proposal)
        self.progress.emit(len(self._completed), self._total)
        if not self._in_flight:
            self.all_done.emit()
```

- [ ] **Step 4: Run task tests to verify they pass**

Run: `python -m pytest tests/test_discovery/test_proposal_coordinator.py::ProposalTaskTests -v`

Expected: All six `ProposalTaskTests` PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/discovery/proposal_coordinator.py tests/test_discovery/test_proposal_coordinator.py
git commit -m @'
discovery: add ProposalTask and ProposalCoordinator skeleton

ProposalTask (QRunnable) builds the structured-proposal prompt for
one parsed discovery request, calls the LLM with a one-shot repair
retry on JSON parse failure, and emits the parsed StructuredProposal
via a shared WorkerSignals helper. Falls back to
build_fallback_structured_proposal on any exception, with the
exception text surfaced in review_reason.

ProposalCoordinator owns a QThreadPool and fans tasks out. Coordinator
tests come in a follow-up commit (task 5 of the plan).
'@
```

---

## Task 5: `ProposalCoordinator` test coverage

**Why now:** The coordinator is shipped (Task 4) but only the task half is tested. This task adds the coordinator's own unit tests with a synchronous fake task factory.

**Files:**
- Test: `tests/test_discovery/test_proposal_coordinator.py`

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_discovery/test_proposal_coordinator.py`:

```python
from PySide6.QtCore import QRunnable

from icharlotte_core.discovery.proposal_coordinator import ProposalCoordinator
from icharlotte_core.discovery.response_context_index import build_context_chunks


def _parsed_two_requests():
    return ParsedDiscovery(
        discovery_type="SI",
        propounding_party="P",
        responding_party="D",
        set_number=1,
        set_word="ONE",
        case_number="1",
        requests=[
            ParsedRequest(number="1", text="Identify witnesses."),
            ParsedRequest(number="2", text="State all facts."),
        ],
    )


class _SyncFakeTask(QRunnable):
    """Run the task body inline so tests don't need an event loop."""

    def __init__(self, *, signals, request, parsed, context_packet,
                 selected_rules, response_rules, fi_mode, override_instruction="",
                 proposal_factory):
        super().__init__()
        self.setAutoDelete(False)
        self._signals = signals
        self._request = request
        self._proposal_factory = proposal_factory

    def run(self):
        proposal = self._proposal_factory(self._request)
        self._signals.proposal_ready.emit(self._request.number, proposal)


def _factory(proposal_factory):
    def _make(**kwargs):
        return _SyncFakeTask(proposal_factory=proposal_factory, **kwargs)
    return _make


class ProposalCoordinatorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._app = QApplication.instance() or QApplication([])

    def _proposal_for(self, request):
        return StructuredProposal(
            request_number=request.number,
            proposed_substantive_response=f"Answer for {request.number}.",
        )

    def test_start_enqueues_one_task_per_request(self):
        parsed = _parsed_two_requests()
        emitted = []
        coordinator = ProposalCoordinator(
            max_concurrent=2,
            task_factory=_factory(self._proposal_for),
        )
        coordinator.proposal_ready.connect(
            lambda num, prop: emitted.append((num, prop))
        )
        enqueued = coordinator.start(
            parsed=parsed,
            selected_rules=[],
            context_chunks=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
        )
        self.assertEqual(enqueued, ["1", "2"])
        self.assertEqual([num for num, _ in emitted], ["1", "2"])

    def test_progress_signal_counts_completions(self):
        parsed = _parsed_two_requests()
        progress = []
        coordinator = ProposalCoordinator(
            max_concurrent=2,
            task_factory=_factory(self._proposal_for),
        )
        coordinator.progress.connect(lambda done, total: progress.append((done, total)))
        coordinator.start(
            parsed=parsed,
            selected_rules=[],
            context_chunks=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
        )
        self.assertEqual(progress, [(1, 2), (2, 2)])

    def test_all_done_emitted_once_after_last_task(self):
        parsed = _parsed_two_requests()
        all_done_count = []
        coordinator = ProposalCoordinator(
            max_concurrent=2,
            task_factory=_factory(self._proposal_for),
        )
        coordinator.all_done.connect(lambda: all_done_count.append(1))
        coordinator.start(
            parsed=parsed,
            selected_rules=[],
            context_chunks=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
        )
        self.assertEqual(all_done_count, [1])
        self.assertTrue(coordinator.is_done())

    def test_skipped_fi_fixed_request_is_not_enqueued(self):
        # In FI fixed mode, request 1.1 (a fixed-response number) is skipped.
        parsed = ParsedDiscovery(
            discovery_type="FI",
            propounding_party="P",
            responding_party="D",
            set_number=1,
            set_word="ONE",
            case_number="1",
            requests=[
                ParsedRequest(number="1.1", text="State your name."),
                ParsedRequest(number="17.1", text="Identify denials."),
            ],
        )
        rules = ResponseRules()
        # Make 1.1 look like it has a fixed response.
        rules.fi_15_1_response = None  # noop, but ensures attribute exists.

        # Patch get_fi_fixed_response to return text for "1.1".
        from icharlotte_core.discovery import proposal_coordinator as pc_module

        original = pc_module.get_fi_fixed_response
        pc_module.get_fi_fixed_response = lambda number, _rules: (
            "Fixed." if number == "1.1" else None
        )
        try:
            emitted = []
            coordinator = ProposalCoordinator(
                max_concurrent=2,
                task_factory=_factory(self._proposal_for),
            )
            coordinator.proposal_ready.connect(
                lambda num, prop: emitted.append(num)
            )
            enqueued = coordinator.start(
                parsed=parsed,
                selected_rules=[],
                context_chunks=[],
                response_rules=rules,
                fi_mode="fixed",
            )
        finally:
            pc_module.get_fi_fixed_response = original

        self.assertEqual(enqueued, ["17.1"])
        self.assertEqual(emitted, ["17.1"])

    def test_regenerate_re_enqueues_named_request(self):
        parsed = _parsed_two_requests()
        emitted = []
        coordinator = ProposalCoordinator(
            max_concurrent=2,
            task_factory=_factory(self._proposal_for),
        )
        coordinator.proposal_ready.connect(
            lambda num, prop: emitted.append(num)
        )
        coordinator.start(
            parsed=parsed,
            selected_rules=[],
            context_chunks=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
        )
        emitted.clear()
        queued = coordinator.regenerate(
            request_number="1",
            parsed=parsed,
            selected_rules=[],
            context_chunks=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
        )
        self.assertTrue(queued)
        self.assertEqual(emitted, ["1"])

    def test_regenerate_no_op_when_already_in_flight(self):
        parsed = _parsed_two_requests()
        # Use a factory that *doesn't* emit so the task stays "in flight".
        def _silent(**kwargs):
            class _Silent(QRunnable):
                def run(_self):
                    pass
            return _Silent()
        coordinator = ProposalCoordinator(max_concurrent=2, task_factory=_silent)
        coordinator.start(
            parsed=parsed,
            selected_rules=[],
            context_chunks=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
        )
        # Request "1" is still in flight (no emit happened).
        result = coordinator.regenerate(
            request_number="1",
            parsed=parsed,
            selected_rules=[],
            context_chunks=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
        )
        self.assertFalse(result)

    def test_cancel_swallows_late_results(self):
        parsed = _parsed_two_requests()
        emitted = []
        coordinator = ProposalCoordinator(
            max_concurrent=2,
            task_factory=_factory(self._proposal_for),
        )
        coordinator.proposal_ready.connect(
            lambda num, prop: emitted.append(num)
        )
        coordinator.cancel()
        coordinator.start(
            parsed=parsed,
            selected_rules=[],
            context_chunks=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
        )
        # Coordinator.start clears the cancelled flag, so this set still runs.
        self.assertEqual(emitted, ["1", "2"])

        # Now cancel mid-flight (simulated: cancel before emitting).
        emitted.clear()
        coordinator._cancelled = True
        # Manually emit a "late" result through the internal signals path.
        coordinator._signals.proposal_ready.emit(
            "1", StructuredProposal(request_number="1"),
        )
        self.assertEqual(emitted, [])

    def test_max_concurrent_caps_at_one(self):
        coordinator = ProposalCoordinator(max_concurrent=0)
        self.assertEqual(coordinator._pool.maxThreadCount(), 1)
        coordinator = ProposalCoordinator(max_concurrent=-3)
        self.assertEqual(coordinator._pool.maxThreadCount(), 1)
```

- [ ] **Step 2: Run tests to verify they pass**

Run: `python -m pytest tests/test_discovery/test_proposal_coordinator.py -v`

Expected: All `ProposalTaskTests` (from Task 4) AND all eight `ProposalCoordinatorTests` PASS.

- [ ] **Step 3: Commit**

```powershell
git add tests/test_discovery/test_proposal_coordinator.py
git commit -m @'
discovery: cover ProposalCoordinator behavior with sync fake tasks

Adds unit tests for start/regenerate/cancel/progress/all_done/skip
paths. Uses a synchronous fake task_factory so the tests don''t
depend on the real thread pool draining.
'@
```

---

## Task 6: `DiscoveryParseWorker`

**Why now:** Trivial extraction from the existing `RespondDiscoveryProposalWorker.run()`. Decouples parsing from drafting and gives the page a clean parse-then-draft seam.

**Files:**
- Create: `icharlotte_core/discovery/discovery_parse_worker.py`
- Test: `tests/test_discovery/test_discovery_parse_worker.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_discovery/test_discovery_parse_worker.py`:

```python
import unittest
from unittest.mock import patch, MagicMock

import pytest

pytest.importorskip("pytestqt")
from PySide6.QtWidgets import QApplication

from icharlotte_core.discovery.discovery_parse_worker import DiscoveryParseWorker
from icharlotte_core.discovery.response_parser import ParsedDiscovery


class DiscoveryParseWorkerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._app = QApplication.instance() or QApplication([])

    def _run_synchronously(self, worker):
        # Invoke the body directly instead of starting a thread.
        worker.run()

    def test_emits_parsed_discovery_on_success(self):
        emitted = []

        with patch(
            "icharlotte_core.discovery.discovery_parse_worker.read_document_text",
            return_value="some discovery text",
        ), patch(
            "icharlotte_core.discovery.discovery_parse_worker.call_llm",
            return_value="ignored",
        ), patch(
            "icharlotte_core.discovery.discovery_parse_worker.parse_llm_response",
            return_value=ParsedDiscovery(
                discovery_type="SI",
                propounding_party="P",
                responding_party="D",
                set_number=1,
                set_word="ONE",
                case_number="1",
                requests=[],
            ),
        ):
            worker = DiscoveryParseWorker(
                discovery_file="x.pdf",
                detected_type="SI",
            )
            worker.parse_finished.connect(
                lambda ok, payload: emitted.append((ok, payload))
            )
            self._run_synchronously(worker)

        self.assertEqual(len(emitted), 1)
        ok, parsed = emitted[0]
        self.assertTrue(ok)
        self.assertIsInstance(parsed, ParsedDiscovery)
        self.assertEqual(parsed.discovery_type, "SI")

    def test_emits_error_when_file_is_empty(self):
        emitted = []
        with patch(
            "icharlotte_core.discovery.discovery_parse_worker.read_document_text",
            return_value="",
        ):
            worker = DiscoveryParseWorker(
                discovery_file="x.pdf", detected_type="SI",
            )
            worker.parse_finished.connect(
                lambda ok, payload: emitted.append((ok, payload))
            )
            self._run_synchronously(worker)
        ok, payload = emitted[0]
        self.assertFalse(ok)
        self.assertIn("Could not read text", payload)

    def test_emits_error_when_llm_returns_nothing(self):
        emitted = []
        with patch(
            "icharlotte_core.discovery.discovery_parse_worker.read_document_text",
            return_value="text",
        ), patch(
            "icharlotte_core.discovery.discovery_parse_worker.call_llm",
            return_value="",
        ):
            worker = DiscoveryParseWorker(
                discovery_file="x.pdf", detected_type="SI",
            )
            worker.parse_finished.connect(
                lambda ok, payload: emitted.append((ok, payload))
            )
            self._run_synchronously(worker)
        ok, payload = emitted[0]
        self.assertFalse(ok)
        self.assertIn("parser did not return", payload)

    def test_emits_error_when_unexpected_exception(self):
        emitted = []
        with patch(
            "icharlotte_core.discovery.discovery_parse_worker.read_document_text",
            side_effect=RuntimeError("disk gone"),
        ):
            worker = DiscoveryParseWorker(
                discovery_file="x.pdf", detected_type="SI",
            )
            worker.parse_finished.connect(
                lambda ok, payload: emitted.append((ok, payload))
            )
            self._run_synchronously(worker)
        ok, payload = emitted[0]
        self.assertFalse(ok)
        self.assertIn("disk gone", payload)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_discovery/test_discovery_parse_worker.py -v`

Expected: All FAIL with `ModuleNotFoundError`.

- [ ] **Step 3: Create the worker**

Create `icharlotte_core/discovery/discovery_parse_worker.py`:

```python
"""One-shot QThread that parses an incoming discovery document via LLM."""
from __future__ import annotations

from PySide6.QtCore import QThread, Signal

from icharlotte_core.discovery.response_parser import (
    ParsedDiscovery,
    build_parse_prompt,
    parse_llm_response,
)
from icharlotte_core.llm_config import call_llm
from icharlotte_core.ui.wizard.pages.respond_discovery_page import (
    _normalize_and_filter_parsed_discovery,
    read_document_text,
)


class DiscoveryParseWorker(QThread):
    """Reads the discovery file, runs the parse LLM, normalizes, emits once."""

    # parse_finished(success: bool, parsed_or_error_message: ParsedDiscovery | str)
    parse_finished = Signal(bool, object)

    def __init__(
        self,
        discovery_file: str,
        detected_type: str,
        parent=None,
    ):
        super().__init__(parent)
        self.discovery_file = discovery_file
        self.detected_type = detected_type

    def run(self) -> None:
        try:
            discovery_text = read_document_text(self.discovery_file)
            if not discovery_text.strip():
                self.parse_finished.emit(False, "Could not read text from the discovery file.")
                return

            raw = call_llm(
                build_parse_prompt(discovery_text),
                "",
                task_type="extraction",
                agent_id="agent_sum_disc",
            )
            if not raw:
                self.parse_finished.emit(False, "The parser did not return a response.")
                return

            parsed = parse_llm_response(raw)
            parsed = _normalize_and_filter_parsed_discovery(
                parsed, self.detected_type, self.discovery_file,
            )
            self.parse_finished.emit(True, parsed)
        except Exception as exc:
            self.parse_finished.emit(False, str(exc))
```

> **Note on the import:** `_normalize_and_filter_parsed_discovery` and `read_document_text` currently live in `respond_discovery_page.py`. Importing UI module-level functions from a worker creates an awkward dependency. We'll address that in Task 8 by moving both functions into `discovery_parse_worker.py` (or into a new `_io.py` helper module). For now, the import works because the page module has no Qt-application-construction side effects at import time.

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_discovery/test_discovery_parse_worker.py -v`

Expected: All four tests PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/discovery/discovery_parse_worker.py tests/test_discovery/test_discovery_parse_worker.py
git commit -m @'
discovery: add DiscoveryParseWorker (one-shot QThread)

Extracted from RespondDiscoveryProposalWorker.run()''s parsing phase.
Reads the discovery file, runs the extraction-task LLM, normalizes
the result, emits parse_finished(success, parsed_or_error) exactly
once. Lives in its own module so the wizard page can parse once and
re-draft many times without re-parsing.

Currently imports read_document_text and
_normalize_and_filter_parsed_discovery from the wizard page; that
dependency will be inverted in the page-rewiring task.
'@
```

---

## Task 7: `_build_pending_review_state` and `_apply_proposal_to_review_state` helpers

**Why now:** These are the two functions the page calls. Pure, side-effect-free, easy to TDD before plumbing through Qt.

**Files:**
- Modify: `icharlotte_core/discovery/response_generation_engine.py` (append two helpers)
- Test: `tests/test_discovery/test_response_generation_engine.py`

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_discovery/test_response_generation_engine.py`:

```python
class PendingReviewStateTests(unittest.TestCase):
    def _fi_parsed(self):
        return ParsedDiscovery(
            discovery_type="FI",
            propounding_party="P",
            responding_party="D",
            set_number=1,
            set_word="ONE",
            case_number="1",
            requests=[
                ParsedRequest(number="1.1", text="State your name."),
                ParsedRequest(number="17.1", text="Identify denials."),
                ParsedRequest(
                    number="3.7",
                    text="Form interrogatory not applicable.",
                ),
            ],
        )

    def test_pending_state_marks_non_skipped_requests_as_pending(self):
        from icharlotte_core.discovery.response_generation_engine import (
            _build_pending_review_state,
        )
        from icharlotte_core.discovery.response_rules import ResponseRules

        parsed = _parsed()  # SI, two requests
        state = _build_pending_review_state(parsed, ResponseRules(), "custom")
        self.assertEqual(len(state.requests), 2)
        for review in state.requests:
            self.assertTrue(review.is_pending)
            self.assertEqual(review.proposed_substantive_response, "")
            self.assertEqual(review.proposed_objections, "")
            self.assertEqual(review.review_reason, "Generating...")

    def test_pending_state_fills_fixed_fi_responses_immediately(self):
        from icharlotte_core.discovery.response_generation_engine import (
            _build_pending_review_state,
        )
        from icharlotte_core.discovery.response_rules import ResponseRules
        rules = ResponseRules()
        rules.fi_objections_by_number = {"17.1": "Objection text."}
        rules.fi_15_1_response = None  # ensure attribute exists for FI lookups

        # Make 17.1 look like it has a fixed substantive response.
        from icharlotte_core.discovery import response_generation_engine as eng

        original = eng.get_fi_fixed_response
        eng.get_fi_fixed_response = lambda number, _rules: (
            "Identification body." if number == "17.1" else None
        )
        try:
            parsed = self._fi_parsed()
            state = eng._build_pending_review_state(parsed, rules, "fixed")
        finally:
            eng.get_fi_fixed_response = original

        rows = {r.number: r for r in state.requests}
        # 17.1 has a fixed response — not pending.
        self.assertFalse(rows["17.1"].is_pending)
        self.assertIn("Identification", rows["17.1"].proposed_substantive_response)
        # 3.7 is inapplicable — not pending, gets the inapplicable response.
        self.assertFalse(rows["3.7"].is_pending)
        self.assertIn("not applicable", rows["3.7"].proposed_substantive_response.lower())
        # 1.1 has no fixed response and is applicable — pending.
        self.assertTrue(rows["1.1"].is_pending)

    def test_apply_proposal_overwrites_pending_row(self):
        from icharlotte_core.discovery.response_generation_engine import (
            _apply_proposal_to_review_state,
            _build_pending_review_state,
            StructuredProposal,
        )
        from icharlotte_core.discovery.response_rules import ResponseRules

        parsed = _parsed()
        state = _build_pending_review_state(parsed, ResponseRules(), "custom")
        proposal = StructuredProposal(
            request_number="1",
            proposed_substantive_response="Drafted answer.",
        )

        review = _apply_proposal_to_review_state(
            review_state=state,
            req_number="1",
            proposal=proposal,
            parsed=parsed,
            selected_rules=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
        )

        self.assertIsNotNone(review)
        self.assertFalse(review.is_pending)
        self.assertEqual(review.proposed_substantive_response, "Drafted answer.")
        # The state object itself reflects the change.
        self.assertFalse(state.requests[0].is_pending)

    def test_apply_proposal_returns_none_for_unknown_request_number(self):
        from icharlotte_core.discovery.response_generation_engine import (
            _apply_proposal_to_review_state,
            _build_pending_review_state,
            StructuredProposal,
        )
        from icharlotte_core.discovery.response_rules import ResponseRules

        parsed = _parsed()
        state = _build_pending_review_state(parsed, ResponseRules(), "custom")
        result = _apply_proposal_to_review_state(
            review_state=state,
            req_number="999",
            proposal=StructuredProposal(request_number="999"),
            parsed=parsed,
            selected_rules=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
        )
        self.assertIsNone(result)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_discovery/test_response_generation_engine.py::PendingReviewStateTests -v`

Expected: All four tests FAIL with `ImportError` for the new helpers.

- [ ] **Step 3: Add the helpers**

In `icharlotte_core/discovery/response_generation_engine.py`, append at the bottom:

```python
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
    # Preserve approval state (user may have approved a pending placeholder)
    # and quick-objection selections, since those are user-driven UI state.
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
            get_fi_fixed_response,
        )
        if not detect_inapplicable_fi(parsed_request.number) and (
            get_fi_fixed_response(parsed_request.number, response_rules) is None
        ):
            if proposal.needs_review:
                new_review.needs_review = True
            if proposal.review_reason.strip():
                new_review.review_reason = proposal.review_reason.strip()

    return new_review
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_discovery/test_response_generation_engine.py -v`

Expected: All tests in the file PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/discovery/response_generation_engine.py tests/test_discovery/test_response_generation_engine.py
git commit -m @'
discovery: add helpers for streaming review-state updates

_build_pending_review_state produces a ReviewState where every
non-skipped request is is_pending=True with "Generating..." as the
review_reason. FI fixed-mode rows (fixed responses and inapplicable
items) are filled immediately so the user can review them while the
LLM works on the rest.

_apply_proposal_to_review_state merges one StructuredProposal into a
named row, clears is_pending, and preserves user-driven selections
(quick-objection IDs).
'@
```

---

## Task 8: Rewire `RespondDiscoverySettingsPage` to use parse worker + coordinator

**Why now:** All primitives exist. This task replaces `RespondDiscoveryProposalWorker` usage end-to-end and removes the now-dead code.

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`
- Test: `tests/test_wizard/test_respond_discovery_page.py`

- [ ] **Step 1: Move `read_document_text` and `_normalize_and_filter_parsed_discovery` to a helper module to break the circular dependency**

Create `icharlotte_core/discovery/_io.py`:

```python
"""Pure I/O helpers reused by the wizard page and the parse worker."""
from __future__ import annotations

import os

from icharlotte_core.discovery.form_interrogatory_selection import (
    complete_selected_form_interrogatories,
    extract_selected_form_interrogatory_numbers,
)
from icharlotte_core.discovery.response_parser import ParsedDiscovery
from icharlotte_core.discovery.response_type_detector import normalize_discovery_type

try:
    import fitz
except ImportError:  # pragma: no cover - depends on local install
    fitz = None


def read_document_text(path: str) -> str:
    """Extract text from a supported context or discovery file."""
    if not path or not os.path.isfile(path):
        return ""
    lower = path.lower()
    if lower.endswith(".pdf"):
        if not fitz:
            return ""
        doc = fitz.open(path)
        try:
            return "\n".join(page.get_text() for page in doc)
        finally:
            doc.close()
    if lower.endswith(".docx"):
        from docx import Document
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs)
    if lower.endswith(".txt"):
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            return fh.read()
    return ""


def read_first_page_text(path: str) -> str:
    """Read first-page text for type detection without parsing the whole PDF."""
    if not path or not path.lower().endswith(".pdf") or not fitz:
        return ""
    if not os.path.isfile(path):
        return ""
    doc = fitz.open(path)
    try:
        if len(doc) == 0:
            return ""
        return doc[0].get_text()
    finally:
        doc.close()


def normalize_and_filter_parsed_discovery(
    parsed: ParsedDiscovery,
    detected_type: str,
    discovery_file: str,
) -> ParsedDiscovery:
    """Canonicalize discovery type and keep only checked FROG items."""
    normalized_detected = normalize_discovery_type(detected_type)
    parsed.discovery_type = normalized_detected or normalize_discovery_type(
        parsed.discovery_type
    )
    if parsed.discovery_type != "FI":
        return parsed
    selected_numbers = extract_selected_form_interrogatory_numbers(discovery_file)
    return complete_selected_form_interrogatories(
        parsed, discovery_file, selected_numbers,
    )
```

- [ ] **Step 2: Update `discovery_parse_worker.py` to import from `_io`**

Edit `icharlotte_core/discovery/discovery_parse_worker.py`, replace the import:

```python
from icharlotte_core.discovery._io import (
    normalize_and_filter_parsed_discovery,
    read_document_text,
)
```

And in `run()`, change `_normalize_and_filter_parsed_discovery(...)` → `normalize_and_filter_parsed_discovery(...)`.

- [ ] **Step 3: Update `respond_discovery_page.py` to re-export from `_io`**

In `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`, replace the local definitions of `read_document_text`, `read_first_page_text`, and `_normalize_and_filter_parsed_discovery` (around lines 95–131 and 1282–1299) with thin re-exports so existing imports in tests keep working:

```python
from icharlotte_core.discovery._io import (
    normalize_and_filter_parsed_discovery as _normalize_and_filter_parsed_discovery,
    read_document_text,
    read_first_page_text,
)
```

Delete the now-redundant `fitz` import at the top of the page file (the `_io` module owns that).

- [ ] **Step 4: Run discovery + wizard tests to confirm the refactor is clean**

Run: `python -m pytest tests/test_discovery tests/test_wizard -v`

Expected: All tests PASS (or pre-existing skips/xfails). If a wizard test imports `_build_structured_proposal_map`, `_draft_substantive_response_map`, or `RespondDiscoveryProposalWorker` and that test is now broken, leave it broken — Step 7 of this task will update those tests.

- [ ] **Step 5: Commit the refactor**

```powershell
git add icharlotte_core/discovery/_io.py icharlotte_core/discovery/discovery_parse_worker.py icharlotte_core/ui/wizard/pages/respond_discovery_page.py
git commit -m @'
discovery: extract _io helpers to break worker/page cycle

read_document_text, read_first_page_text, and
normalize_and_filter_parsed_discovery moved to a new
icharlotte_core/discovery/_io module. The wizard page re-exports
them so existing call sites keep working. DiscoveryParseWorker now
imports directly from _io instead of reaching into the UI module.
'@
```

- [ ] **Step 6: Write the failing page tests for streaming wiring**

Append to `tests/test_wizard/test_respond_discovery_page.py` (after the existing `RespondDiscoverySettingsPageTests`):

```python
from icharlotte_core.discovery.response_generation_engine import (
    StructuredProposal,
)


class _MakeParsedSI:
    def __call__(self, n=2):
        from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
        return ParsedDiscovery(
            discovery_type="SI",
            propounding_party="P",
            responding_party="D",
            set_number=1,
            set_word="ONE",
            case_number="1",
            requests=[
                ParsedRequest(number=str(i + 1), text=f"Request {i + 1}.")
                for i in range(n)
            ],
        )


class StreamingReviewStateTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._app = QApplication.instance() or QApplication([])

    def _make_page(self):
        return RespondDiscoverySettingsPage(
            case_root="",
            file_number="1234.001",
            discovery_file=r"C:\case\srogg.pdf",
            detected_type="SI",
        )

    def test_review_screen_opens_with_all_requests_pending(self):
        page = self._make_page()
        parsed = _MakeParsedSI()(n=3)
        page._open_review_screen_with_pending(parsed)

        self.assertIsNotNone(page.review_state)
        self.assertEqual(len(page.review_state.requests), 3)
        for review in page.review_state.requests:
            self.assertTrue(review.is_pending)

    def test_apply_proposal_clears_pending_on_matching_row(self):
        page = self._make_page()
        parsed = _MakeParsedSI()(n=2)
        page._open_review_screen_with_pending(parsed)
        page.parsed_discovery = parsed

        proposal = StructuredProposal(
            request_number="1",
            proposed_substantive_response="Drafted answer for 1.",
        )
        page._on_proposal_ready("1", proposal)

        self.assertFalse(page.review_state.requests[0].is_pending)
        self.assertEqual(
            page.review_state.requests[0].proposed_substantive_response,
            "Drafted answer for 1.",
        )
        # The second row is still pending.
        self.assertTrue(page.review_state.requests[1].is_pending)

    def test_proposal_does_not_overwrite_user_edits(self):
        page = self._make_page()
        parsed = _MakeParsedSI()(n=2)
        page._open_review_screen_with_pending(parsed)
        page.parsed_discovery = parsed

        # User edits the first row's response while it's still pending.
        page.review_state.requests[0].proposed_substantive_response = "USER TYPED THIS"

        proposal = StructuredProposal(
            request_number="1",
            proposed_substantive_response="Drafted answer.",
        )
        page._on_proposal_ready("1", proposal)

        # The user text survives; the new draft is stashed in pending_replacement.
        self.assertEqual(
            page.review_state.requests[0].proposed_substantive_response,
            "USER TYPED THIS",
        )
        self.assertIsNotNone(page.review_state.requests[0].pending_replacement)
        self.assertEqual(
            page.review_state.requests[0].pending_replacement.proposed_substantive_response,
            "Drafted answer.",
        )
        # is_pending is cleared either way.
        self.assertFalse(page.review_state.requests[0].is_pending)

    def test_finalize_disabled_until_no_pending_remain(self):
        page = self._make_page()
        parsed = _MakeParsedSI()(n=2)
        page._open_review_screen_with_pending(parsed)
        page.parsed_discovery = parsed
        page._show_review()

        # No proposals delivered yet → Finalize is disabled.
        self.assertFalse(page.finalize_btn.isEnabled())

        # Deliver both proposals.
        page._on_proposal_ready(
            "1",
            StructuredProposal(request_number="1", proposed_substantive_response="A1"),
        )
        page._on_proposal_ready(
            "2",
            StructuredProposal(request_number="2", proposed_substantive_response="A2"),
        )
        page._on_coordinator_all_done()
        # All rows arrived; Finalize is enabled once user approves them via UI.
        # We only check the "no longer disabled by pending" half here.
        self.assertTrue(page.finalize_btn.isEnabled())
```

- [ ] **Step 7: Run the new tests to verify they fail**

Run: `python -m pytest tests/test_wizard/test_respond_discovery_page.py::StreamingReviewStateTests -v`

Expected: FAIL — `_open_review_screen_with_pending` and `_on_proposal_ready` don't exist yet.

- [ ] **Step 8: Rewire the page**

In `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`:

**8a. Add imports near the top:**

```python
from icharlotte_core.discovery.discovery_parse_worker import DiscoveryParseWorker
from icharlotte_core.discovery.proposal_coordinator import ProposalCoordinator
from icharlotte_core.discovery.response_context_index import (
    build_context_chunks,
    format_context_packet,
    select_context_packet,
)
from icharlotte_core.discovery.response_generation_engine import (
    _apply_proposal_to_review_state,
    _build_pending_review_state,
    StructuredProposal,
    build_fallback_structured_proposal,
    build_structured_proposal_prompt,
    generate_review_state,
    parse_structured_proposal_response,
)
from icharlotte_core.llm_config import LLMConfig
```

(Some of these may already be imported. Avoid duplicates.)

**8b. In `RespondDiscoverySettingsPage.__init__`, add coordinator/worker state:**

```python
        self._parse_worker: DiscoveryParseWorker | None = None
        self._coordinator: ProposalCoordinator | None = None
        self._context_chunks: list = []
        self._loaded_response_rules = None
```

(Place near the existing `self._proposal_worker = None` line, then remove that line.)

**8c. Replace `_generate_proposals` and `_on_proposals_finished` with the new flow:**

```python
    def _generate_proposals(self) -> None:
        self.next_btn.setEnabled(False)
        self.status_label.setText("Parsing discovery and drafting proposed responses...")
        if self.parsed_discovery is not None:
            self._kick_off_drafting(self.parsed_discovery)
            return

        self._parse_worker = DiscoveryParseWorker(
            discovery_file=self.discovery_file,
            detected_type=self.detected_type,
            parent=self,
        )
        self._parse_worker.parse_finished.connect(self._on_parse_finished)
        self._parse_worker.start()

    def _on_parse_finished(self, success: bool, payload: object) -> None:
        self._parse_worker = None
        if not success:
            self.next_btn.setEnabled(True)
            self.status_label.setText(str(payload))
            return
        self.parsed_discovery = payload  # ParsedDiscovery
        self._kick_off_drafting(self.parsed_discovery)

    def _kick_off_drafting(self, parsed) -> None:
        self.next_btn.setEnabled(True)
        response_rules = load_respond_response_rules(self.file_number)
        self._loaded_response_rules = response_rules
        context_text_by_path = {
            path: read_document_text(path)
            for path in self.context_files
            if os.path.isfile(path)
        }
        self._context_chunks = build_context_chunks(context_text_by_path)
        self.review_state = _build_pending_review_state(
            parsed, response_rules, self.fi_mode,
        )
        self._open_review_screen_with_pending(parsed)

        if self._coordinator is not None:
            self._coordinator.cancel()
        max_concurrent = LLMConfig().discovery_response_max_concurrent()
        self._coordinator = ProposalCoordinator(
            max_concurrent=max_concurrent, parent=self,
        )
        self._coordinator.proposal_ready.connect(self._on_proposal_ready)
        self._coordinator.progress.connect(self._on_coordinator_progress)
        self._coordinator.all_done.connect(self._on_coordinator_all_done)
        self._coordinator.start(
            parsed=parsed,
            selected_rules=self.selected_rules(),
            context_chunks=self._context_chunks,
            response_rules=response_rules,
            fi_mode=self.fi_mode,
        )

    def _open_review_screen_with_pending(self, parsed) -> None:
        if self.review_state is None:
            response_rules = self._loaded_response_rules or load_respond_response_rules(
                self.file_number,
            )
            self.review_state = _build_pending_review_state(
                parsed, response_rules, self.fi_mode,
            )
        self._show_review()

    def _on_proposal_ready(self, req_number: str, proposal) -> None:
        if not self.review_state or not self.parsed_discovery:
            return
        target_index = next(
            (i for i, r in enumerate(self.review_state.requests)
             if r.number == req_number),
            -1,
        )
        if target_index < 0:
            return
        previous = self.review_state.requests[target_index]
        user_edited = (
            (previous.proposed_substantive_response or "").strip() != ""
            or (previous.proposed_objections or "").strip() != ""
        )
        if user_edited:
            # Stash for the conflict banner; do not overwrite.
            previous.pending_replacement = proposal
            previous.is_pending = False
            self._refresh_review_row(target_index)
            return

        response_rules = self._loaded_response_rules or load_respond_response_rules(
            self.file_number,
        )
        _apply_proposal_to_review_state(
            review_state=self.review_state,
            req_number=req_number,
            proposal=proposal,
            parsed=self.parsed_discovery,
            selected_rules=self.selected_rules(),
            response_rules=response_rules,
            fi_mode=self.fi_mode,
        )
        self._refresh_review_row(target_index)

    def _on_coordinator_progress(self, completed: int, total: int) -> None:
        if hasattr(self, "review_status_label") and self.review_status_label:
            self.review_status_label.setText(f"Generated {completed} / {total}")

    def _on_coordinator_all_done(self) -> None:
        self._refresh_finalize_button()

    def _refresh_review_row(self, target_index: int) -> None:
        if target_index == self._current_review_index:
            self._load_current_review()
        self._refresh_finalize_button()

    def _refresh_finalize_button(self) -> None:
        if not hasattr(self, "finalize_btn"):
            return
        any_pending = self.review_state and any(
            r.is_pending for r in self.review_state.requests
        )
        self.finalize_btn.setEnabled(not any_pending)
```

**8d. Delete dead code:**

Remove these from `respond_discovery_page.py`:
- The entire `RespondDiscoveryProposalWorker` class (lines 832–913 in the original).
- `_build_structured_proposal_map` (916–963).
- `_build_structured_proposal_repair_prompt` (979–985).
- `_ensure_context_warning` (988–1002) — its logic moved into `ProposalTask`.
- `_generate_review_state_from_proposals` (1005–1027) — replaced by `_build_pending_review_state` + `_apply_proposal_to_review_state`.
- `_apply_fixed_fi_proposal_warnings` (1030–1053) — its logic moved into `_apply_proposal_to_review_state`.
- `_callbacks_from_proposal_map` (1056–1095).
- `_draft_substantive_response_map`, `_build_combined_substantive_prompt`, `_parse_combined_response_map` (1098–1183) — no longer reached.

Keep `_generate_review_from_parsed` only if any test/caller still uses it; otherwise delete. (Grep for `_generate_review_from_parsed` first; the wizard registry should now call `_kick_off_drafting` directly.)

- [ ] **Step 9: Update the existing wizard tests that referenced removed symbols**

Edit `tests/test_wizard/test_respond_discovery_page.py`:

- Remove `_build_structured_proposal_map` and `_draft_substantive_response_map` from the imports (line 17–24).
- If `RespondDiscoveryWorker` is still used in tests (it is — it's the *final assembly* worker, different class), keep it.
- Any test that explicitly invoked `_build_structured_proposal_map` should be deleted. Read the file to identify them; delete with confidence — the new tests in step 6 cover the replacement behavior.

- [ ] **Step 10: Run all wizard + discovery tests**

Run: `python -m pytest tests/test_discovery tests/test_wizard -v`

Expected: All PASS.

- [ ] **Step 11: Commit**

```powershell
git add icharlotte_core/ui/wizard/pages/respond_discovery_page.py tests/test_wizard/test_respond_discovery_page.py
git commit -m @'
wizard: stream proposals into review screen as they arrive

Replace the single serial RespondDiscoveryProposalWorker with a
DiscoveryParseWorker + ProposalCoordinator pair. The wizard now:

- Parses the discovery file once on entry, caches the result.
- Opens the review screen immediately with is_pending=True rows.
- Fans up to N concurrent ProposalTask jobs out (N from
  llm_preferences.json discovery_response.max_concurrent_proposals,
  default 8).
- Patches each row in-place as proposal_ready fires.
- Never overwrites user edits — stashes new drafts in
  pending_replacement for a later banner action.
- Disables Finalize while any row is_pending.

Removes RespondDiscoveryProposalWorker,
_build_structured_proposal_map, _draft_substantive_response_map,
_build_combined_substantive_prompt, _parse_combined_response_map,
_ensure_context_warning, _generate_review_state_from_proposals,
_apply_fixed_fi_proposal_warnings, and _callbacks_from_proposal_map
from the wizard page. Their logic now lives in ProposalTask,
ProposalCoordinator, _build_pending_review_state, and
_apply_proposal_to_review_state.
'@
```

---

## Task 9: Review screen UI — status indicators, status bar, pending visuals

**Why now:** Plumbing is in place; this task makes the streaming visible to the user.

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/respond_discovery_page.py` (review widget builder + load methods)
- Test: `tests/test_wizard/test_respond_discovery_page.py`

- [ ] **Step 1: Write the failing UI tests**

Append to `tests/test_wizard/test_respond_discovery_page.py`:

```python
class ReviewScreenStatusIndicatorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._app = QApplication.instance() or QApplication([])

    def _make_page_with_pending(self, n=3):
        from icharlotte_core.discovery.response_parser import (
            ParsedDiscovery, ParsedRequest,
        )
        parsed = ParsedDiscovery(
            discovery_type="SI",
            propounding_party="P",
            responding_party="D",
            set_number=1, set_word="ONE", case_number="1",
            requests=[
                ParsedRequest(number=str(i + 1), text=f"Request {i + 1}.")
                for i in range(n)
            ],
        )
        page = RespondDiscoverySettingsPage(
            case_root="", file_number="1234.001",
            discovery_file=r"C:\case\srogg.pdf",
            detected_type="SI",
        )
        page.parsed_discovery = parsed
        page._open_review_screen_with_pending(parsed)
        return page, parsed

    def test_pending_row_shows_generating_placeholder_in_response_pane(self):
        page, _ = self._make_page_with_pending()
        page._current_review_index = 0
        page._load_current_review()
        self.assertEqual(page.response_edit.toPlainText(), "")
        self.assertTrue(page.response_edit.isReadOnly())
        self.assertIn("Generating", page.review_warning_label.text())

    def test_completed_row_enables_edits(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
        )
        page, parsed = self._make_page_with_pending()
        page._on_proposal_ready(
            "1",
            StructuredProposal(
                request_number="1", proposed_substantive_response="Done."
            ),
        )
        page._current_review_index = 0
        page._load_current_review()
        self.assertFalse(page.response_edit.isReadOnly())
        self.assertEqual(page.response_edit.toPlainText(), "Done.")

    def test_status_bar_reflects_progress(self):
        page, parsed = self._make_page_with_pending(n=4)
        page._on_coordinator_progress(2, 4)
        self.assertIn("2", page.review_status_label.text())
        self.assertIn("4", page.review_status_label.text())

    def test_request_number_header_includes_status_icon(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
        )
        page, parsed = self._make_page_with_pending(n=2)
        page._current_review_index = 0
        page._load_current_review()
        self.assertIn("⏳", page.request_label.text())  # hourglass
        page._on_proposal_ready(
            "1",
            StructuredProposal(
                request_number="1",
                proposed_substantive_response="Done.",
                needs_review=True,
                review_reason="weak context",
            ),
        )
        page._load_current_review()
        self.assertIn("⚠", page.request_label.text())  # warning
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_wizard/test_respond_discovery_page.py::ReviewScreenStatusIndicatorTests -v`

Expected: FAIL — `review_status_label` doesn't exist; pending placeholder isn't set; icons aren't in the header.

- [ ] **Step 3: Add the status bar and pending visuals**

In `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`:

**3a. In `_build_review_widget`, before the nav row, add a status label:**

Find:

```python
        layout.addLayout(quick_row)

        nav = QHBoxLayout()
```

Replace with:

```python
        layout.addLayout(quick_row)

        self.review_status_label = QLabel("")
        self.review_status_label.setStyleSheet("color: #666; font-style: italic;")
        layout.addWidget(self.review_status_label)

        nav = QHBoxLayout()
```

**3b. In `_load_current_review`, add status-icon logic at the top of the existing method (after the `review = self._current_review()` block) and toggle read-only on the editors:**

Replace the body of `_load_current_review` (the existing `_load_current_review` you already have) with:

```python
    def _load_current_review(self) -> None:
        review = self._current_review()
        count = len(self.review_state.requests) if self.review_state else 0
        if not review:
            self.request_label.setText("No parsed requests.")
            return
        icon = self._status_icon_for(review)
        self.review_title.setText(
            f"Review Discovery Responses ({self._current_review_index + 1} of {count})"
        )
        self.request_label.setText(f"{icon} Request No. {review.number}".strip())
        self.request_text.setText(review.request_text)

        is_pending = bool(review.is_pending)
        self.objection_edit.setPlainText(
            "" if is_pending else review.proposed_objections
        )
        self.response_edit.setPlainText(
            "" if is_pending else review.proposed_substantive_response
        )
        self.objection_edit.setReadOnly(is_pending)
        self.response_edit.setReadOnly(is_pending)
        self.objection_edit.setPlaceholderText("Generating..." if is_pending else "")
        self.response_edit.setPlaceholderText("Generating..." if is_pending else "")

        review_reason = (review.review_reason or "").strip()
        if is_pending:
            self.review_warning_label.setText("Generating...")
            self.review_warning_label.show()
        elif review.needs_review:
            warning_text = (
                f"Needs review: {review_reason}" if review_reason else "Needs review."
            )
            self.review_warning_label.setText(warning_text)
            self.review_warning_label.show()
        else:
            self.review_warning_label.setText("")
            self.review_warning_label.hide()

        for rule_id, cb in self._quick_checks.items():
            cb.blockSignals(True)
            cb.setChecked(rule_id in review.selected_quick_objection_ids)
            cb.blockSignals(False)
        self.prev_btn.setEnabled(self._current_review_index > 0)
        self.next_review_btn.setEnabled(self._current_review_index < count - 1)

    def _status_icon_for(self, review) -> str:
        if review.is_pending:
            return "⏳"  # ⏳
        if review.approved:
            return "✓"  # ✓
        if review.needs_review:
            return "⚠"  # ⚠
        # "Edited" = non-default content but not yet approved.
        has_content = bool(
            (review.proposed_objections or "").strip()
            or (review.proposed_substantive_response or "").strip()
        )
        if has_content:
            return "✏"  # ✏
        return ""
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_respond_discovery_page.py::ReviewScreenStatusIndicatorTests -v`

Expected: All four PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/ui/wizard/pages/respond_discovery_page.py tests/test_wizard/test_respond_discovery_page.py
git commit -m @'
wizard: surface streaming progress in the review screen

- Pending rows show "Generating..." placeholders with editors set to
  read-only.
- Request header shows a status icon (hourglass, warning, edit pen,
  check) reflecting is_pending / needs_review / edited / approved.
- New review_status_label under the quick-response row shows
  "Generated X / N" as ProposalCoordinator progresses.
- Warning banner shows "Generating..." while pending, replacing the
  prior empty state.
'@
```

---

## Task 10: Regenerate button + edit-conflict banner

**Why now:** Final user-visible affordance. Leverages the existing coordinator hook and `pending_replacement` field.

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/respond_discovery_page.py` (review widget builder + new methods)
- Test: `tests/test_wizard/test_respond_discovery_page.py`

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_wizard/test_respond_discovery_page.py`:

```python
class RegenerateAndConflictTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._app = QApplication.instance() or QApplication([])

    def _make_page(self, n=2):
        from icharlotte_core.discovery.response_parser import (
            ParsedDiscovery, ParsedRequest,
        )
        parsed = ParsedDiscovery(
            discovery_type="SI",
            propounding_party="P",
            responding_party="D",
            set_number=1, set_word="ONE", case_number="1",
            requests=[
                ParsedRequest(number=str(i + 1), text=f"Request {i + 1}.")
                for i in range(n)
            ],
        )
        page = RespondDiscoverySettingsPage(
            case_root="", file_number="1234.001",
            discovery_file=r"C:\case\srogg.pdf",
            detected_type="SI",
        )
        page.parsed_discovery = parsed
        page._open_review_screen_with_pending(parsed)
        return page, parsed

    def test_regenerate_button_calls_coordinator_with_current_request(self):
        page, parsed = self._make_page()
        # Install a real coordinator with a no-op stub.
        regen_calls = []
        class _Stub:
            def regenerate(self, request_number, **kwargs):
                regen_calls.append((request_number, kwargs.get("override_instruction", "")))
                return True
            def cancel(self): pass
            def start(self, **kwargs): pass
            def is_done(self): return False
        page._coordinator = _Stub()
        page._context_chunks = []
        page._loaded_response_rules = None

        page._current_review_index = 0
        page.regenerate_instruction_edit.setText("more privilege")
        page._on_regenerate_clicked()

        self.assertEqual(regen_calls, [("1", "more privilege")])

    def test_regenerate_marks_row_pending_until_proposal_arrives(self):
        page, parsed = self._make_page()
        class _Stub:
            def regenerate(self, **kwargs): return True
            def cancel(self): pass
            def start(self, **kwargs): pass
            def is_done(self): return False
        page._coordinator = _Stub()
        page._context_chunks = []
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
        )
        page._on_proposal_ready(
            "1",
            StructuredProposal(
                request_number="1", proposed_substantive_response="Initial."
            ),
        )
        self.assertFalse(page.review_state.requests[0].is_pending)

        page._current_review_index = 0
        page.regenerate_instruction_edit.setText("")
        page._on_regenerate_clicked()
        self.assertTrue(page.review_state.requests[0].is_pending)

    def test_conflict_banner_view_apply_replaces_user_text(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
        )
        page, parsed = self._make_page()
        page.review_state.requests[0].proposed_substantive_response = "USER TEXT"
        page._on_proposal_ready(
            "1",
            StructuredProposal(
                request_number="1", proposed_substantive_response="DRAFT"
            ),
        )
        # Banner is exposed via page.conflict_banner_visible_for(req_number).
        page._current_review_index = 0
        page._load_current_review()
        self.assertTrue(page.conflict_banner.isVisible())

        page._apply_pending_replacement()
        self.assertEqual(
            page.review_state.requests[0].proposed_substantive_response,
            "DRAFT",
        )
        self.assertIsNone(page.review_state.requests[0].pending_replacement)
        self.assertFalse(page.conflict_banner.isVisible())

    def test_conflict_banner_discard_keeps_user_text(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
        )
        page, parsed = self._make_page()
        page.review_state.requests[0].proposed_substantive_response = "USER TEXT"
        page._on_proposal_ready(
            "1",
            StructuredProposal(
                request_number="1", proposed_substantive_response="DRAFT"
            ),
        )
        page._current_review_index = 0
        page._load_current_review()
        page._discard_pending_replacement()
        self.assertEqual(
            page.review_state.requests[0].proposed_substantive_response,
            "USER TEXT",
        )
        self.assertIsNone(page.review_state.requests[0].pending_replacement)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_wizard/test_respond_discovery_page.py::RegenerateAndConflictTests -v`

Expected: All four FAIL — none of those widgets/methods exist yet.

- [ ] **Step 3: Add the Regenerate row and conflict banner to the review widget**

In `_build_review_widget` in `respond_discovery_page.py`, after the `layout.addLayout(quick_row)` line and before `self.review_status_label = QLabel("")`, add:

```python
        # Conflict banner — shown when a new draft arrives for a row the
        # user has already edited.
        self.conflict_banner = QWidget()
        banner_layout = QHBoxLayout(self.conflict_banner)
        banner_layout.setContentsMargins(8, 6, 8, 6)
        self.conflict_banner.setStyleSheet(
            "background-color: #fff7d6; border: 1px solid #d4b85a; border-radius: 4px;"
        )
        banner_label = QLabel("New draft available — your edits are preserved.")
        banner_label.setStyleSheet("font-weight: 600;")
        banner_layout.addWidget(banner_label, 1)
        self.conflict_view_btn = QPushButton("View")
        self.conflict_view_btn.clicked.connect(self._view_pending_replacement)
        self.conflict_apply_btn = QPushButton("Apply")
        self.conflict_apply_btn.clicked.connect(self._apply_pending_replacement)
        self.conflict_discard_btn = QPushButton("Discard")
        self.conflict_discard_btn.clicked.connect(self._discard_pending_replacement)
        for btn in (self.conflict_view_btn, self.conflict_apply_btn, self.conflict_discard_btn):
            banner_layout.addWidget(btn)
        self.conflict_banner.hide()
        layout.addWidget(self.conflict_banner)

        # Regenerate row.
        regenerate_row = QHBoxLayout()
        from PySide6.QtWidgets import QLineEdit
        self.regenerate_instruction_edit = QLineEdit()
        self.regenerate_instruction_edit.setPlaceholderText(
            "Optional instructions for regeneration (e.g., lean harder on privilege)"
        )
        regenerate_row.addWidget(self.regenerate_instruction_edit, 1)
        self.regenerate_btn = QPushButton("Regenerate")
        self.regenerate_btn.clicked.connect(self._on_regenerate_clicked)
        regenerate_row.addWidget(self.regenerate_btn)
        layout.addLayout(regenerate_row)
```

(The `QLineEdit` import can also go at the top of the file with the other QtWidgets imports.)

**Add the handler methods in `RespondDiscoverySettingsPage`:**

```python
    def _on_regenerate_clicked(self) -> None:
        review = self._current_review()
        if not review or not self.parsed_discovery or not self._coordinator:
            return
        instruction = self.regenerate_instruction_edit.text().strip()
        response_rules = self._loaded_response_rules or load_respond_response_rules(
            self.file_number,
        )
        queued = self._coordinator.regenerate(
            request_number=review.number,
            parsed=self.parsed_discovery,
            selected_rules=self.selected_rules(),
            context_chunks=self._context_chunks,
            response_rules=response_rules,
            fi_mode=self.fi_mode,
            override_instruction=instruction,
        )
        if not queued:
            return
        review.is_pending = True
        review.review_reason = "Regenerating..."
        review.needs_review = False
        review.approved = False
        review.pending_replacement = None
        self.regenerate_instruction_edit.clear()
        self._load_current_review()
        self._refresh_finalize_button()

    def _view_pending_replacement(self) -> None:
        review = self._current_review()
        if not review or not review.pending_replacement:
            return
        QMessageBox.information(
            self,
            f"New draft for Request No. {review.number}",
            review.pending_replacement.proposed_substantive_response
            or "(no substantive text in draft)",
        )

    def _apply_pending_replacement(self) -> None:
        review = self._current_review()
        if not review or not review.pending_replacement:
            return
        response_rules = self._loaded_response_rules or load_respond_response_rules(
            self.file_number,
        )
        _apply_proposal_to_review_state(
            review_state=self.review_state,
            req_number=review.number,
            proposal=review.pending_replacement,
            parsed=self.parsed_discovery,
            selected_rules=self.selected_rules(),
            response_rules=response_rules,
            fi_mode=self.fi_mode,
        )
        # _apply_proposal_to_review_state replaces the row; re-fetch.
        new_review = self._current_review()
        if new_review:
            new_review.pending_replacement = None
        self._load_current_review()

    def _discard_pending_replacement(self) -> None:
        review = self._current_review()
        if not review:
            return
        review.pending_replacement = None
        self._load_current_review()
```

**Update `_load_current_review` to toggle the conflict banner:**

At the end of the existing `_load_current_review` body (after the prev/next button enable logic), add:

```python
        has_pending_replacement = review.pending_replacement is not None
        if hasattr(self, "conflict_banner"):
            self.conflict_banner.setVisible(has_pending_replacement)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_respond_discovery_page.py::RegenerateAndConflictTests -v`

Expected: All four PASS.

- [ ] **Step 5: Run the full test suite to catch regressions**

Run: `python -m pytest tests/test_discovery tests/test_wizard -v`

Expected: All PASS.

- [ ] **Step 6: Commit**

```powershell
git add icharlotte_core/ui/wizard/pages/respond_discovery_page.py tests/test_wizard/test_respond_discovery_page.py
git commit -m @'
wizard: per-request Regenerate + edit-conflict banner

A new Regenerate row under the response editor lets users re-queue
the current request with optional one-line override instructions.
The row goes pending until the coordinator returns a fresh draft.

When a fresh draft lands on a row the user has already typed in, the
draft is stashed in pending_replacement and a yellow banner appears
with View / Apply / Discard buttons. User text is never overwritten
automatically.
'@
```

---

## Task 11: Manual verification on a real discovery set

**Why last:** Per CLAUDE.md, any feature change requires a real-app run before we call it done. This task is no-code; it produces a manual-verification note.

- [ ] **Step 1: Pick a test case**

Find a case folder in `Z:\Shared\Current Clients` (or any local case) with a 35+ request set — typically a Form Interrogatories, Special Interrogatories, or RPD set. Note the file number.

- [ ] **Step 2: Launch the app and run the wizard**

```powershell
python iCharlotte.py
```

Steps:
- Open the case.
- Navigate to the Wizard tab.
- Choose Respond to Discovery.
- Pick the discovery file.
- Configure rules.
- Pick context files.

- [ ] **Step 3: Time and observe**

Record on the spec doc or in a session note:
- Wall-clock time from "Next: Context Files" click to "review screen appears".
- Wall-clock time from review screen open to "all rows non-pending".
- Whether the status bar updates smoothly (no UI freezes).
- Whether the per-row status icons transition correctly (⏳ → ⚠/✓/✏).
- Whether the finalize button stays disabled until all rows complete.

Expected: ~5–8× speedup vs. the prior serial version on Gemini Flash (50-request set: minutes → 30–60s).

- [ ] **Step 4: Trigger a deliberate failure**

Briefly unset `GEMINI_API_KEY` (or point at a bogus context PDF that yields zero text) and confirm:
- The wizard doesn't crash.
- Affected rows show ⚠ with a useful `review_reason`.
- The user can still finalize after manually editing the failed rows.

Restore the API key.

- [ ] **Step 5: Trigger an edit conflict**

- Start the wizard, get to the review screen.
- While rows are still pending, click into the first pending row and type some text into the response editor.
- Wait for the proposal to arrive.
- Confirm:
  - The user text is preserved (not overwritten).
  - The yellow banner shows.
  - View shows the new draft in a dialog.
  - Apply replaces the user text with the new draft.
  - Discard keeps the user text.

- [ ] **Step 6: Verify the final Word document**

After approving every row and clicking Finalize, confirm:
- The output `.docx` is created at the expected path under `<case>\.icharlotte\wizard_previews\respond_to_discovery\`.
- `validate_discovery_response_docx` passes (no error dialog).
- The document opens cleanly in Word with correct numbering, captions, and objections.

- [ ] **Step 7: Note results in MEMORY.md**

Per the global instructions, record findings under a new topic file:

Edit `C:\Users\ASerpik.DESKTOP-MRIMK0D\.claude\projects\C--geminiterminal2\memory\MEMORY.md` index to add:

```
- `respond_discovery_streaming.md` — parallel proposal generation, streaming review, edit-conflict banner, regenerate button
```

Then create `C:\Users\ASerpik.DESKTOP-MRIMK0D\.claude\projects\C--geminiterminal2\memory\respond_discovery_streaming.md` with:
- Wall-clock numbers (before / after).
- Any quirks observed (e.g., specific rate-limit thresholds hit, model behaviors).
- Configuration tweaks recommended for future use.

- [ ] **Step 8: Commit the memory note**

```powershell
git add ../../Users/ASerpik.DESKTOP-MRIMK0D/.claude/projects/C--geminiterminal2/memory/MEMORY.md ../../Users/ASerpik.DESKTOP-MRIMK0D/.claude/projects/C--geminiterminal2/memory/respond_discovery_streaming.md
git commit -m "memory: note from respond-to-discovery streaming verification"
```

(If the memory directory is outside the repo working tree, skip the commit and just save the file — the global memory file is user-scoped, not project-scoped.)

---

## Self-Review (Plan vs. Spec)

**Coverage check** — every spec requirement maps to a task:

| Spec requirement | Task |
|---|---|
| Drop dead `proposed_objections` field from prompt | Task 1 |
| Add `is_pending` to `RequestReview` | Task 2 |
| Add `pending_replacement` to `RequestReview` (session-only) | Task 2 |
| `LLMConfig.discovery_response_max_concurrent()` | Task 3 |
| `WorkerSignals`, `ProposalTask` (with repair retry + fallback) | Task 4 |
| `ProposalCoordinator` (thread pool, start/regenerate/cancel) | Tasks 4 (code) + 5 (tests) |
| `DiscoveryParseWorker` extracted | Task 6 |
| `_build_pending_review_state`, `_apply_proposal_to_review_state` | Task 7 |
| Page rewiring: parse worker + coordinator | Task 8 |
| Remove `RespondDiscoveryProposalWorker`, `_build_structured_proposal_map`, etc. | Task 8 |
| Status indicators (⏳ ⚠ ✓ ✏), status bar, pending visuals | Task 9 |
| Per-request Regenerate button + override instruction | Task 10 |
| Edit-conflict banner (View / Apply / Discard) | Task 10 |
| Manual verification on real set | Task 11 |
| `ReviewState.all_approved` blocks while any pending | Task 2 |

No spec requirement is unaccounted for.

**Placeholder scan** — no TBDs, no "implement appropriate error handling", no "similar to Task N", no references to undefined types. Every code block is complete.

**Type consistency** — `ProposalCoordinator.regenerate` signature in Task 4 (`request_number, parsed, selected_rules, context_chunks, response_rules, fi_mode, override_instruction`) matches the call site in Task 10's `_on_regenerate_clicked`. `_apply_proposal_to_review_state` signature in Task 7 (`review_state, req_number, proposal, parsed, selected_rules, response_rules, fi_mode`) matches the call sites in Task 8 (`_on_proposal_ready`) and Task 10 (`_apply_pending_replacement`). `parse_finished` signal carries `(bool, object)` consistently across Task 6 emission and Task 8 consumption.

**Scope check** — single coherent PR touching one feature flow. ~700 lines of new code, ~400 deleted. Falls cleanly into 11 tasks, each independently commit-able and testable.
