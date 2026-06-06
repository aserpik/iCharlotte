"""Fail-closed Phase 1 orchestrator scaffold for agentic legal research."""
from __future__ import annotations

from collections.abc import Callable
from typing import Any

from .models import (
    CourtListenerMode,
    DeepResearchRequest,
    ResearchPacket,
    ResearchRun,
    ResearchStatus,
    ResearchStep,
    SourcePolicy,
    normalize_source_policy,
)


def _has_selected_source(policy: SourcePolicy) -> bool:
    return (
        policy.firm_authority
        or policy.local_corpus
        or policy.courtlistener_mode != CourtListenerMode.OFF
        or policy.ca_leginfo
        or policy.ca_courts_recent
    )


def run_deep_research(
    request: DeepResearchRequest,
    llm_callback: Callable[..., Any],
    source_registry: Any,
    status_callback: Callable[[str], Any] | None = None,
) -> ResearchRun:
    """Prepare a deep-research run and fail closed until adapters exist."""
    if status_callback is not None:
        status_callback("Preparing deep research request...")

    policy = normalize_source_policy(request.source_policy)
    warnings: list[str] = []
    if not _has_selected_source(policy):
        warnings.append("No selected research source is usable.")
    if source_registry is None:
        warnings.append("No deep-research source adapters were provided.")

    step = ResearchStep(
        phase="initialization",
        input_summary=request.question,
        decision="Research did not run because Phase 1 has no retrieval adapters.",
        warnings=list(warnings),
    )

    return ResearchRun(
        request=request,
        status=ResearchStatus.FAILED,
        steps=[step],
        warnings=list(warnings),
        packet=ResearchPacket(warnings=list(warnings)),
        diagnostics={"source_policy": policy},
    )
