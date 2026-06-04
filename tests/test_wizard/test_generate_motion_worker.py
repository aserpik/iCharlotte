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
    # Top-level ground AND the selected subsection leaf were both researched.
    assert "Boilerplate objections are improper" in captured["targets"]
    assert any("evasive" in t.lower() for t in captured["targets"])
    # Parity knobs.
    assert captured["max_workers"] == 2
    assert captured["cache_dir"] and "generate_motion" in captured["cache_dir"]


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
