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
        # Verify the LLM call signature matches what production code expects.
        _, args, kwargs = call_llm.mock_calls[0]
        self.assertEqual(args[1], "")
        self.assertEqual(kwargs.get("task_type"), "general")
        self.assertEqual(kwargs.get("agent_id"), "agent_sum_disc")

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
