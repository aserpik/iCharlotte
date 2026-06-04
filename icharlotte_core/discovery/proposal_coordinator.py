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
    normalize_rpd_substantive_response,
    parse_structured_proposal_response,
)
from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
from icharlotte_core.discovery.response_rule_library import ResponseRule
from icharlotte_core.discovery.response_rules import ResponseRules
from icharlotte_core.discovery.response_type_detector import normalize_discovery_type


CallLLMFn = Callable[..., str]

_AGENT_ID = "agent_sum_disc"


class WorkerSignals(QObject):
    """Signal carrier for ProposalTask.

    QRunnable can't define signals directly; we route them through this
    helper. One WorkerSignals instance can be shared by many tasks (the
    coordinator does that).
    """

    proposal_ready = Signal(str, object)  # request_number, StructuredProposal


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
            raw = call_llm(prompt, "", task_type="general", agent_id=_AGENT_ID)
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

        return self._ensure_context_warning(self._normalize_proposal(proposal))

    def _normalize_proposal(self, proposal: StructuredProposal) -> StructuredProposal:
        """Snap RPD substantive text to a canonical statement at creation time.

        Keeps the conflict-banner preview consistent with what apply will store.
        Idempotent; a no-op for non-RPD discovery types.
        """
        if normalize_discovery_type(self.parsed.discovery_type) != "RPD":
            return proposal
        snapped = normalize_rpd_substantive_response(
            proposal.proposed_substantive_response,
        )
        if snapped == proposal.proposed_substantive_response:
            return proposal
        return replace(proposal, proposed_substantive_response=snapped)

    def _call_repair(self, call_llm: CallLLMFn, raw_text: str) -> StructuredProposal | None:
        repair_prompt = (
            "Repair this structured discovery proposal JSON. "
            "Return ONLY one valid JSON object with the same schema. "
            f"The request_number must be {self.request.number}.\n\n"
            f"INVALID RESPONSE:\n{raw_text}"
        )
        try:
            repaired_raw = call_llm(
                repair_prompt, "", task_type="general", agent_id=_AGENT_ID,
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
        """Enqueue a task for every non-skipped request. Returns enqueued numbers.

        ``_total`` is set BEFORE any task is dispatched so that synchronously-
        completing tasks (e.g., a test factory using DirectConnection) see the
        correct denominator in the ``progress`` signal and don't fire
        ``all_done`` prematurely.
        """
        self._cancelled = False
        self._completed.clear()
        self._in_flight.clear()

        enqueued: list[str] = [
            req.number for req in parsed.requests
            if not _should_skip_request(req, parsed, response_rules, fi_mode)
        ]
        self._total = len(enqueued)
        self._in_flight.update(enqueued)

        if self._total == 0:
            self.all_done.emit()
            return enqueued

        enqueued_set = set(enqueued)
        for req in parsed.requests:
            if req.number not in enqueued_set:
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
            self._pool.start(task)
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
        """Re-enqueue one request. No-op if already in flight. Returns True if queued.

        The denominator (``_total``) is bumped via
        ``max(_total, _completed + _in_flight)``. Repeated regeneration of
        the same request can creep the denominator upward over time even
        though the amount of work is unchanged — acceptable for UI
        progress display, where a slowly-growing total is preferable to
        making the progress bar appear to go backwards.
        """
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
        """Mark the coordinator cancelled; ignore any further task results.

        Tasks already submitted to the thread pool may still run to
        completion — Qt's QThreadPool does not support cooperative
        cancellation of in-flight QRunnables. Their results are discarded
        in ``_on_proposal_ready`` via the ``_cancelled`` flag.

        ``all_done`` is NOT emitted by cancel(). Callers needing a
        cleanup signal should rely on their own teardown path.

        ``start()`` clears ``_cancelled`` so the coordinator can be
        re-used for a fresh run. If you call ``start()`` while old
        cancelled tasks are still draining, the freshly-cleared flag
        means those stale results would be processed as new — call
        ``self._pool.waitForDone()`` between ``cancel()`` and a fresh
        ``start()`` if that ordering matters.
        """
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
