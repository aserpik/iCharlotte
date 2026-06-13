from pathlib import Path


def test_motion_database_style_samples_loads_selected_folder_samples(tmp_path):
    from icharlotte_core.motion_drafting.style_samples import load_motion_database_style_samples

    selected = tmp_path / "Replies" / "Motion to Compel"
    nested = selected / "Separate Statements"
    nested.mkdir(parents=True)
    first = selected / "reply.docx"
    second = nested / "reply_sample.pdf"
    ignored = selected / "notes.csv"
    first.write_text("CAPTION\nARGUMENT\nDatabase reply style one.", encoding="utf-8")
    second.write_text("DISCUSSION\nDatabase reply style two.", encoding="utf-8")
    ignored.write_text("ignore me", encoding="utf-8")

    samples = load_motion_database_style_samples(
        str(selected),
        max_samples=3,
        extract_fn=lambda path: Path(path).read_text(encoding="utf-8"),
        cache_dir=str(tmp_path / ".cache"),
    )

    assert samples == [
        "ARGUMENT\nDatabase reply style one.",
        "DISCUSSION\nDatabase reply style two.",
    ]


def test_motion_database_style_samples_caps_results(tmp_path):
    from icharlotte_core.motion_drafting.style_samples import load_motion_database_style_samples

    selected = tmp_path / "Motions" / "Motion - Compel"
    selected.mkdir(parents=True)
    for index in range(5):
        (selected / f"sample-{index}.txt").write_text(
            f"ARGUMENT\nStyle sample {index}.",
            encoding="utf-8",
        )

    samples = load_motion_database_style_samples(str(selected), max_samples=3)

    assert len(samples) == 3
    assert samples[0] == "ARGUMENT\nStyle sample 0."


def test_generate_worker_passes_motion_database_samples_first(monkeypatch, tmp_path):
    import icharlotte_core.motion_generation.samples as configured_samples
    import icharlotte_core.motion_drafting.style_samples as db_samples
    import icharlotte_core.ui.wizard.pages.generate_motion_page as gm
    from icharlotte_core.opposition.models import DraftDocument, MotionMetadata, OutlineNode

    captured = {}
    monkeypatch.setattr(gm, "extract_context_bundle", lambda files: ("ctx", []))
    monkeypatch.setattr(gm, "_make_local_corpus", lambda: None)
    monkeypatch.delenv("COURTLISTENER_API_TOKEN", raising=False)
    monkeypatch.setattr(gm, "research_arguments", lambda *a, **k: [])
    monkeypatch.setattr(configured_samples, "load_exemplars", lambda _tid: ["CONFIGURED STYLE"])
    monkeypatch.setattr(gm, "_firm_style_exemplars", lambda *a, **k: ["FIRM STYLE"])
    monkeypatch.setattr(db_samples, "load_motion_database_style_samples", lambda _path: ["DATABASE STYLE"])
    monkeypatch.setattr(gm, "identify_key_legal_issues", lambda *a, **k: [])
    monkeypatch.setattr(gm, "extract_citations", lambda body: [])
    monkeypatch.setattr(gm, "assemble_motion_preview", lambda **k: k.get("output_path"))
    monkeypatch.setattr(gm.DiscoveryAssembler, "find_caption_page", staticmethod(lambda p: ""))

    def fake_draft_motion(*_args, **kwargs):
        captured["style_exemplars"] = kwargs.get("style_exemplars")
        return DraftDocument(title="Motion", body_text="Draft body.")

    monkeypatch.setattr(gm, "draft_motion", fake_draft_motion)

    settings = {
        "motion_type_id": "compel",
        "motion_type_name": "Motion to Compel",
        "motion_type_source_path": str(tmp_path / "MOTION DATABASE" / "Motion - Compel"),
        "target_files": [],
        "metadata": MotionMetadata(
            motion_type="Motion to Compel",
            relief_requested="Compel further responses",
            principal_arguments=["Boilerplate objections are improper"],
        ).to_dict(),
        "outline": [OutlineNode(text="Argument", selected=True).to_dict()],
    }
    worker = gm.GenerateMotionWorker(case_path=str(tmp_path), file_number="123", settings=settings)

    results = {}
    worker.finished_result.connect(lambda ok, payload: results.update(ok=ok, payload=payload))
    worker.run()

    assert results.get("ok") is True
    assert captured["style_exemplars"] == ["DATABASE STYLE", "FIRM STYLE", "CONFIGURED STYLE"]
    assert results["payload"].diagnostics["style"]["motion_database_samples"] == 1


def test_oppose_worker_passes_motion_database_samples_first(monkeypatch, tmp_path):
    import icharlotte_core.motion_drafting.style_samples as db_samples
    import icharlotte_core.ui.wizard.pages.oppose_motion_page as page
    from icharlotte_core.opposition.models import DraftDocument, MotionMetadata, OutlineNode

    captured = {}
    monkeypatch.setattr(page, "extract_document_text", lambda _path: type("R", (), {"success": True, "text": "motion", "error": ""})())
    monkeypatch.setattr(page, "extract_context_bundle", lambda _paths: ("context", []))
    monkeypatch.setattr(page, "_make_local_corpus", lambda: None)
    monkeypatch.delenv("COURTLISTENER_API_TOKEN", raising=False)
    monkeypatch.setattr(page, "research_arguments", lambda *a, **k: [])
    monkeypatch.setattr(page.StyleExampleRegistry, "load", classmethod(lambda cls, path: cls(path=path, examples=[])))
    monkeypatch.setattr(page, "_firm_style_exemplars", lambda *a, **k: ["FIRM OPP STYLE"])
    monkeypatch.setattr(db_samples, "load_motion_database_style_samples", lambda _path: ["DATABASE OPP STYLE"])
    monkeypatch.setattr(page, "extract_citations", lambda body: [])
    monkeypatch.setattr(page, "assemble_opposition_preview", lambda **_kwargs: None)
    monkeypatch.setattr(page, "validate_opposition_docx", lambda _path: type("V", (), {"has_errors": False})())
    monkeypatch.setattr(page.DiscoveryAssembler, "find_caption_page", staticmethod(lambda _path: ""))

    def fake_draft_memorandum(**kwargs):
        captured["style_exemplars"] = kwargs.get("style_exemplars")
        return DraftDocument(title="Opposition", body_text="Draft body.")

    monkeypatch.setattr(page, "draft_memorandum", fake_draft_memorandum)

    worker = page.OpposeMotionWorker(
        case_path=str(tmp_path),
        file_number="123",
        settings={
            "motion_file": "/tmp/motion.pdf",
            "context_files": [],
            "motion_type_source_path": str(tmp_path / "MOTION DATABASE" / "Oppositions" / "Motion to Compel"),
            "metadata": MotionMetadata(motion_type="Motion to Compel").to_dict(),
            "outline": [OutlineNode(text="Argument", selected=True).to_dict()],
        },
    )

    results = []
    worker.finished_result.connect(lambda ok, payload: results.append((ok, payload)))
    worker.run()

    assert results and results[0][0] is True
    assert captured["style_exemplars"] == ["DATABASE OPP STYLE", "FIRM OPP STYLE"]
