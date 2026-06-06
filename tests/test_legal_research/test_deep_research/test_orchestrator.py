from icharlotte_core.legal_research.deep_research import (
    CourtListenerMode,
    DeepResearchRequest,
    ResearchStatus,
    SourcePolicy,
)
from icharlotte_core.legal_research.deep_research.orchestrator import (
    run_deep_research,
)


def test_orchestrator_scaffold_fails_closed_without_sources():
    request = DeepResearchRequest(question="What is the summary judgment standard?")

    run = run_deep_research(
        request,
        llm_callback=lambda *_args, **_kwargs: None,
        source_registry=None,
    )

    assert run.request is request
    assert run.status == ResearchStatus.FAILED
    assert "No deep-research source adapters were provided." in run.warnings
    assert run.packet.warnings == run.warnings
    assert run.steps[0].phase == "initialization"
    assert run.steps[0].input_summary == request.question
    assert run.steps[0].warnings == run.warnings


def test_orchestrator_reports_unusable_source_policy():
    request = DeepResearchRequest(
        question="Can the court limit disproportionate discovery?",
        source_policy=SourcePolicy(
            firm_authority=False,
            local_corpus=False,
            courtlistener_mode=CourtListenerMode.OFF,
            ca_leginfo=False,
            ca_courts_recent=False,
        ),
    )

    run = run_deep_research(
        request,
        llm_callback=lambda *_args, **_kwargs: None,
        source_registry=object(),
    )

    assert run.status == ResearchStatus.FAILED
    assert "No selected research source is usable." in run.warnings


def test_orchestrator_reports_status_and_does_not_call_llm_callback():
    request = DeepResearchRequest(question="What authorities govern issue sanctions?")
    statuses = []
    llm_calls = []

    run = run_deep_research(
        request,
        llm_callback=lambda *_args, **_kwargs: llm_calls.append("called"),
        source_registry=object(),
        status_callback=statuses.append,
    )

    assert run.status == ResearchStatus.FAILED
    assert statuses == ["Preparing deep research request..."]
    assert llm_calls == []


def test_run_deep_research_is_public_package_export():
    from icharlotte_core.legal_research.deep_research import run_deep_research as exported

    assert exported is run_deep_research
