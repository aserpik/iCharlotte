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


# ---------------------------------------------------------------------------
# ProposalCoordinator tests
# ---------------------------------------------------------------------------

from PySide6.QtCore import QRunnable

from icharlotte_core.discovery.proposal_coordinator import ProposalCoordinator


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


def _install_inline_pool(coordinator):
    """Patch coordinator._pool so task.run() is invoked synchronously."""
    coordinator._pool.start = lambda task: task.run()


class ProposalCoordinatorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._app = QApplication.instance() or QApplication([])

    def _proposal_for(self, request):
        return StructuredProposal(
            request_number=request.number,
            proposed_substantive_response=f"Answer for {request.number}.",
        )

    def _new_coordinator(self, *, task_factory=None, max_concurrent=2):
        coordinator = ProposalCoordinator(
            max_concurrent=max_concurrent,
            task_factory=task_factory,
        )
        _install_inline_pool(coordinator)
        return coordinator

    def test_start_enqueues_one_task_per_request(self):
        parsed = _parsed_two_requests()
        emitted = []
        coordinator = self._new_coordinator(task_factory=_factory(self._proposal_for))
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
        coordinator = self._new_coordinator(task_factory=_factory(self._proposal_for))
        coordinator.progress.connect(
            lambda done, total: progress.append((done, total))
        )
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
        coordinator = self._new_coordinator(task_factory=_factory(self._proposal_for))
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
        # In FI fixed mode, requests with a fixed response are skipped.
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

        # Patch get_fi_fixed_response to return text for "1.1".
        from icharlotte_core.discovery import proposal_coordinator as pc_module

        original = pc_module.get_fi_fixed_response
        pc_module.get_fi_fixed_response = lambda number, _rules: (
            "Fixed." if number == "1.1" else None
        )
        try:
            emitted = []
            coordinator = self._new_coordinator(task_factory=_factory(self._proposal_for))
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
        coordinator = self._new_coordinator(task_factory=_factory(self._proposal_for))
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

        # Use a silent factory that produces tasks that do NOT emit.
        def _silent(**kwargs):
            class _Silent(QRunnable):
                def run(_self):
                    pass
            return _Silent()

        coordinator = self._new_coordinator(task_factory=_silent)
        coordinator.start(
            parsed=parsed,
            selected_rules=[],
            context_chunks=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
        )
        # Request "1" is still in flight (silent tasks never emit).
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
        coordinator = self._new_coordinator(task_factory=_factory(self._proposal_for))
        coordinator.proposal_ready.connect(
            lambda num, prop: emitted.append(num)
        )
        # Pre-cancel — start() must clear the flag.
        coordinator.cancel()
        coordinator.start(
            parsed=parsed,
            selected_rules=[],
            context_chunks=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
        )
        self.assertEqual(emitted, ["1", "2"])

        # Now simulate a late-arriving result after cancel().
        emitted.clear()
        coordinator._cancelled = True
        coordinator._signals.proposal_ready.emit(
            "1", StructuredProposal(request_number="1"),
        )
        self.assertEqual(emitted, [])

    def test_max_concurrent_caps_at_one(self):
        coordinator = ProposalCoordinator(max_concurrent=0)
        self.assertEqual(coordinator._pool.maxThreadCount(), 1)
        coordinator = ProposalCoordinator(max_concurrent=-3)
        self.assertEqual(coordinator._pool.maxThreadCount(), 1)
