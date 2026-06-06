"""GenerateMotionWorker researches each selected outline subsection (parity
with oppose), not just the top-level grounds."""
from __future__ import annotations

import pytest

pytest.importorskip("PySide6")


def test_worker_researches_subsection_leaves(monkeypatch, tmp_path):
    import icharlotte_core.motion_generation.samples as samples
    import icharlotte_core.ui.wizard.pages.generate_motion_page as gm
    from icharlotte_core.opposition.models import (
        DraftDocument, MotionMetadata, OutlineNode,
    )

    # Stub everything around the research call so run() is fast + offline.
    monkeypatch.setattr(gm, "extract_context_bundle", lambda files: ("ctx", []))
    monkeypatch.setattr(gm, "_make_local_corpus", lambda: object())  # truthy → research runs, no token needed
    monkeypatch.setattr(gm, "draft_motion",
                        lambda *a, **k: DraftDocument(title="M", body_text="Body."))
    monkeypatch.setattr(gm, "extract_citations", lambda body: [])
    monkeypatch.setattr(gm, "assemble_motion_preview", lambda **k: k.get("output_path"))
    monkeypatch.setattr(gm.DiscoveryAssembler, "find_caption_page", staticmethod(lambda p: ""))
    monkeypatch.setattr(samples, "load_exemplars", lambda tid: [])

    captured = {}

    def fake_research(targets, **kwargs):
        captured["targets"] = list(targets)
        captured["max_workers"] = kwargs.get("max_workers")
        captured["cache_dir"] = kwargs.get("cache_dir")
        captured["prompt_namespace"] = kwargs.get("prompt_namespace")
        return []

    monkeypatch.setattr(gm, "research_arguments", fake_research)

    settings = {
        "motion_type_id": "compel",
        "motion_type_name": "Motion to Compel",
        "target_files": [],
        "metadata": MotionMetadata(
            motion_type="Motion to Compel",
            relief_requested="Compel further responses",
            principal_arguments=["Boilerplate objections are improper"],
        ).to_dict(),
        "outline": [
            OutlineNode(text="Argument", selected=True, children=[
                OutlineNode(text="Responses were evasive and incomplete", selected=True),
            ]).to_dict(),
        ],
    }
    worker = gm.GenerateMotionWorker(case_path=str(tmp_path), file_number="123", settings=settings)

    results = {}
    worker.finished_result.connect(lambda ok, payload: results.update(ok=ok, payload=payload))
    worker.run()  # synchronous (not via QThread.start)

    assert results.get("ok") is True
    draft = results["payload"]
    # Top-level ground AND the selected subsection leaf were both researched.
    assert "Boilerplate objections are improper" in captured["targets"]
    assert any("evasive" in t.lower() for t in captured["targets"])
    # Local corpus runs can use a slightly higher worker count because they do
    # not hit CourtListener search rate limits.
    assert captured["max_workers"] == 3
    assert captured["cache_dir"] and "generate_motion" in captured["cache_dir"]
    assert captured.get("prompt_namespace") == "generate_motion"
    assert draft.diagnostics["research"]["target_count"] == len(captured["targets"])
    assert draft.diagnostics["research"]["source"] == "local_corpus"
    assert draft.diagnostics["research"]["workers"] == 3
    assert draft.diagnostics["output"]["preview_path"]
    assert draft.diagnostics["phase_seconds"]["total"] >= 0


def test_worker_adds_replacement_candidates_for_flagged_citations(monkeypatch, tmp_path):
    import icharlotte_core.motion_generation.samples as samples
    import icharlotte_core.ui.wizard.pages.generate_motion_page as gm
    from icharlotte_core.opposition.citation_parser import Citation
    from icharlotte_core.opposition.models import (
        CitationVerification, DraftDocument, MotionMetadata, OutlineNode,
    )

    monkeypatch.setattr(gm, "extract_context_bundle", lambda files: ("ctx", []))
    monkeypatch.setattr(gm, "_make_local_corpus", lambda: object())
    monkeypatch.setattr(gm, "research_arguments", lambda *a, **k: [])
    monkeypatch.setattr(
        gm, "draft_motion",
        lambda *a, **k: DraftDocument(
            title="M",
            body_text="The claim fails under Smith v. Jones (2010) 50 Cal.4th 100.",
        ),
    )
    monkeypatch.setattr(
        gm,
        "extract_citations",
        lambda body: [
            Citation(
                kind="case",
                raw_text="Smith v. Jones (2010) 50 Cal.4th 100",
                normalized="Smith v. Jones (2010) 50 Cal.4th 100",
                reporter_citation="50 Cal.4th 100",
                proposition="The claim fails.",
            )
        ],
    )
    monkeypatch.setattr(gm, "pool_membership_check", lambda citations, retrieved: (citations, []))

    class FakeVerifier:
        def verify_all(self, citations, on_progress=None):
            return [
                CitationVerification(
                    citation_text=citations[0].raw_text,
                    normalized_citation=citations[0].normalized,
                    kind="case",
                    verdict="NOT_SUPPORTED",
                    proposition=citations[0].proposition,
                    note="Wrong rule.",
                )
            ]

    monkeypatch.setattr(
        gm,
        "build_local_opposition_verifier",
        lambda **kwargs: FakeVerifier(),
    )
    monkeypatch.setattr(
        gm,
        "find_replacement_candidates",
        lambda **kwargs: [
            CitationVerification(
                citation_text="Brown v. Davis (2015) 60 Cal.App.4th 200",
                normalized_citation="Brown v. Davis (2015) 60 Cal.App.4th 200",
                kind="case",
                verdict="SUPPORTED",
                note="Direct replacement.",
            )
        ],
    )
    monkeypatch.setattr(gm, "assemble_motion_preview", lambda **k: k.get("output_path"))
    monkeypatch.setattr(gm.DiscoveryAssembler, "find_caption_page", staticmethod(lambda p: ""))
    monkeypatch.setattr(samples, "load_exemplars", lambda tid: [])

    settings = {
        "motion_type_id": "demurrer",
        "motion_type_name": "Demurrer",
        "target_files": [],
        "metadata": MotionMetadata(
            motion_type="Demurrer",
            relief_requested="Sustain demurrer",
            principal_arguments=["The claim fails."],
        ).to_dict(),
        "outline": [OutlineNode(text="Argument", selected=True).to_dict()],
    }
    worker = gm.GenerateMotionWorker(case_path=str(tmp_path), file_number="123", settings=settings)

    results = {}
    worker.finished_result.connect(lambda ok, payload: results.update(ok=ok, payload=payload))
    worker.run()

    assert results.get("ok") is True
    citation = results["payload"].citations[0]
    assert citation.verdict == "NOT_SUPPORTED"
    assert len(citation.replacement_candidates) == 1
    assert citation.replacement_candidates[0]["verdict"] == "SUPPORTED"
    assert results["payload"].diagnostics["citations"]["replacement_candidates"] == 1


def test_worker_keeps_live_courtlistener_research_conservative(monkeypatch, tmp_path):
    import icharlotte_core.motion_generation.samples as samples
    import icharlotte_core.ui.wizard.pages.generate_motion_page as gm
    from icharlotte_core.opposition.models import (
        DraftDocument, MotionMetadata, OutlineNode,
    )

    monkeypatch.setattr(gm, "extract_context_bundle", lambda files: ("ctx", []))
    monkeypatch.setattr(gm, "_make_local_corpus", lambda: None)
    monkeypatch.setattr(gm.os.environ, "get", lambda key, default="": "TOKEN" if key == "COURTLISTENER_API_TOKEN" else default)
    monkeypatch.setattr(
        "icharlotte_core.legal_research.sources.courtlistener.CourtListenerClient",
        lambda token: object(),
    )
    monkeypatch.setattr(gm, "draft_motion",
                        lambda *a, **k: DraftDocument(title="M", body_text="Body."))
    monkeypatch.setattr(gm, "extract_citations", lambda body: [])
    monkeypatch.setattr(gm, "assemble_motion_preview", lambda **k: k.get("output_path"))
    monkeypatch.setattr(gm.DiscoveryAssembler, "find_caption_page", staticmethod(lambda p: ""))
    monkeypatch.setattr(samples, "load_exemplars", lambda tid: [])

    captured = {}

    def fake_research(targets, **kwargs):
        captured["max_workers"] = kwargs.get("max_workers")
        return []

    monkeypatch.setattr(gm, "research_arguments", fake_research)

    settings = {
        "motion_type_id": "compel",
        "motion_type_name": "Motion to Compel",
        "target_files": [],
        "metadata": MotionMetadata(
            motion_type="Motion to Compel",
            relief_requested="Compel further responses",
            principal_arguments=["Boilerplate objections are improper"],
        ).to_dict(),
        "outline": [
            OutlineNode(text="Argument", selected=True, children=[
                OutlineNode(text="Responses were evasive and incomplete", selected=True),
            ]).to_dict(),
        ],
    }
    worker = gm.GenerateMotionWorker(case_path=str(tmp_path), file_number="123", settings=settings)

    results = {}
    worker.finished_result.connect(lambda ok, payload: results.update(ok=ok, payload=payload))
    worker.run()

    assert results.get("ok") is True
    assert captured["max_workers"] == 2


def test_worker_uses_courtlistener_when_local_corpus_is_stale(monkeypatch, tmp_path):
    import icharlotte_core.motion_generation.samples as samples
    import icharlotte_core.ui.wizard.pages.generate_motion_page as gm
    from icharlotte_core.opposition.models import (
        DraftDocument, MotionMetadata, OutlineNode,
    )

    class StaleCorpus:
        def corpus_metadata(self):
            return {
                "source_counts": {"cap": 10},
                "max_decision_date": "2017-05-26",
            }

    live_client = object()
    monkeypatch.setattr(gm, "extract_context_bundle", lambda files: ("ctx", []))
    monkeypatch.setattr(gm, "_make_local_corpus", lambda: StaleCorpus())
    monkeypatch.setattr(gm.os.environ, "get", lambda key, default="": "TOKEN" if key == "COURTLISTENER_API_TOKEN" else default)
    monkeypatch.setattr(
        "icharlotte_core.legal_research.sources.courtlistener.CourtListenerClient",
        lambda token: live_client,
    )
    monkeypatch.setattr(gm, "draft_motion",
                        lambda *a, **k: DraftDocument(title="M", body_text="Body."))
    monkeypatch.setattr(gm, "extract_citations", lambda body: [])
    monkeypatch.setattr(gm, "assemble_motion_preview", lambda **k: k.get("output_path"))
    monkeypatch.setattr(gm.DiscoveryAssembler, "find_caption_page", staticmethod(lambda p: ""))
    monkeypatch.setattr(samples, "load_exemplars", lambda tid: [])

    captured = {}

    def fake_research(targets, **kwargs):
        captured["client"] = kwargs.get("cl_client")
        captured["max_workers"] = kwargs.get("max_workers")
        return []

    monkeypatch.setattr(gm, "research_arguments", fake_research)

    settings = {
        "motion_type_id": "compel",
        "motion_type_name": "Motion to Compel",
        "target_files": [],
        "metadata": MotionMetadata(
            motion_type="Motion to Compel",
            relief_requested="Compel further responses",
            principal_arguments=["Boilerplate objections are improper"],
        ).to_dict(),
        "outline": [
            OutlineNode(text="Argument", selected=True, children=[
                OutlineNode(text="Responses were evasive and incomplete", selected=True),
            ]).to_dict(),
        ],
    }
    worker = gm.GenerateMotionWorker(case_path=str(tmp_path), file_number="123", settings=settings)

    results = {}
    worker.finished_result.connect(lambda ok, payload: results.update(ok=ok, payload=payload))
    worker.run()

    assert results.get("ok") is True
    assert captured["client"] is live_client
    assert captured["max_workers"] == 2
    assert results["payload"].diagnostics["research"]["source"] == "courtlistener"


def test_analysis_worker_passes_motion_name(monkeypatch):
    import icharlotte_core.ui.wizard.pages.generate_motion_page as gm
    from icharlotte_core.opposition.models import MotionMetadata

    captured = {}

    def fake_analyze(config, target_text, *, llm_callback, context_text="", motion_name=""):
        captured["motion_name"] = motion_name
        return MotionMetadata(motion_type=motion_name or "X",
                              relief_requested="r", principal_arguments=["g1"])

    monkeypatch.setattr(gm, "analyze_target", fake_analyze)
    monkeypatch.setattr(gm, "extract_context_bundle", lambda files: ("some facts", []))
    # Harmless stub LLM closures so any outline generation (a later task) makes no real call.
    monkeypatch.setattr(gm, "_make_llms", lambda: ((lambda s, u: ""), (lambda s, u: ""),
                                                   (lambda p: (lambda s, u: ""))))

    settings = {
        "motion_type_id": "generic",
        "motion_type_name": "Motion in Limine to Exclude Witnesses",
        "target_files": ["x.pdf"], "user_relief": "", "user_arguments": [],
    }
    worker = gm.GenerateMotionAnalysisWorker(settings=settings)
    results = {}
    worker.finished_analysis.connect(lambda ok, payload: results.update(ok=ok, payload=payload))
    worker.run()

    assert results["ok"] is True
    assert captured["motion_name"] == "Motion in Limine to Exclude Witnesses"


def test_analysis_worker_emits_nested_outline(monkeypatch):
    import icharlotte_core.ui.wizard.pages.generate_motion_page as gm
    from icharlotte_core.opposition.models import MotionMetadata

    monkeypatch.setattr(gm, "extract_context_bundle", lambda files: ("facts", []))
    monkeypatch.setattr(
        gm, "analyze_target",
        lambda *a, **k: MotionMetadata(
            motion_type=k.get("motion_name") or "M",
            relief_requested="r", principal_arguments=["g1", "g2"]),
    )

    def fake_outline_llm(system, user):
        return ('{"outline": [{"text": "Argument", '
                '"children": [{"text": "Sub A"}, {"text": "Sub B"}]}]}')

    # _make_llms() returns (analysis_llm, draft_llm, make_pass_llm); the worker
    # uses analysis_llm for both analyze_target and generate_motion_outline.
    monkeypatch.setattr(
        gm, "_make_llms",
        lambda: (fake_outline_llm, fake_outline_llm, (lambda p: fake_outline_llm)),
    )

    settings = {
        "motion_type_id": "generic",
        "motion_type_name": "Motion in Limine to Exclude Witnesses",
        "target_files": ["x.pdf"], "user_relief": "", "user_arguments": [],
    }
    worker = gm.GenerateMotionAnalysisWorker(settings=settings)
    out = {}
    worker.finished_analysis.connect(lambda ok, payload: out.update(ok=ok, payload=payload))
    worker.run()

    assert out["ok"] is True
    outline = out["payload"]["outline"]
    arg = [n for n in outline if n.text == "Argument"]
    assert arg and len(arg[0].children) >= 2


def test_context_cache_extracts_once_then_reextracts_on_change(monkeypatch, tmp_path):
    """The analysis pass and the draft pass over the SAME files share one
    extraction; editing the file selection invalidates the cache."""
    import icharlotte_core.ui.wizard.pages.generate_motion_page as gm

    # Reset the module-level single-entry cache for test isolation.
    gm._context_cache["key"] = None
    gm._context_cache["result"] = None

    calls = []

    def fake_bundle(files):
        calls.append(list(files))
        return ("text-" + str(len(calls)), [])

    monkeypatch.setattr(gm, "extract_context_bundle", fake_bundle)

    f1 = tmp_path / "doc.pdf"
    f1.write_text("one")

    # Analysis pass + draft pass over the same file → bundle runs ONCE.
    t1, _ = gm.extract_context_cached([str(f1)])
    t2, _ = gm.extract_context_cached([str(f1)])
    assert t1 == t2 == "text-1"
    assert len(calls) == 1

    # Change the selection → cache invalidates → bundle runs again.
    f2 = tmp_path / "other.pdf"
    f2.write_text("two")
    t3, _ = gm.extract_context_cached([str(f1), str(f2)])
    assert t3 == "text-2"
    assert len(calls) == 2

    # Editing a file's contents (new mtime/size) also invalidates.
    f1.write_text("one-modified-much-longer")
    t4, _ = gm.extract_context_cached([str(f1)])
    assert t4 == "text-3"
    assert len(calls) == 3
