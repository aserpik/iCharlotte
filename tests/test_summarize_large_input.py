"""Regression tests for large-input Summarize Documents behavior."""

import sys
from pathlib import Path
from types import SimpleNamespace


PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "Scripts"))

import summarize  # noqa: E402


def test_summarize_chunks_large_document_before_llm(monkeypatch, tmp_path):
    large_text = (
        "The witness discussed the incident, treatment, symptoms, work limits, "
        "and medical history.\n"
    ) * 2600

    def fake_extract(self, path, progress_callback=None):
        return SimpleNamespace(
            success=True,
            text=large_text,
            char_count=len(large_text),
            page_count=130,
            ocr_pages=[],
            ocr_percentage=0.0,
            error=None,
        )

    calls = []

    def fake_call(self, prompt, text, task_type=None, agent_id=None, pass_name=None, **kwargs):
        calls.append({
            "prompt": prompt,
            "text_len": len(text),
            "task_type": task_type,
            "agent_id": agent_id,
            "pass_name": pass_name,
        })
        if text and len(text) > 80_000:
            return None
        if "partial summaries" in prompt.lower():
            return "Final consolidated summary."
        return "Chunk summary."

    saved = {}

    def fake_save(output_path, content, title, logger):
        saved["content"] = content
        saved["title"] = title
        saved["output_path"] = output_path
        return output_path

    monkeypatch.setattr(summarize.DocumentProcessor, "extract_with_dynamic_ocr", fake_extract)
    monkeypatch.setattr(summarize.LLMCaller, "call", fake_call)
    monkeypatch.setattr(summarize, "safe_append_to_docx", fake_save)

    input_path = tmp_path / "Long Document.pdf"
    input_path.write_bytes(b"%PDF-1.4\n")
    logger = summarize.AgentLogger("SummarizeTest", log_to_file=False)
    ok = summarize.process_document(
        str(input_path),
        logger,
        output_path_override=str(tmp_path / "AI_OUTPUT.docx"),
    )

    assert ok is True
    summary_calls = [
        c for c in calls
        if c["agent_id"] == "agent_summarize" and c["pass_name"] == "summary"
    ]
    assert len(summary_calls) >= 2
    assert all(c["text_len"] <= 80_000 for c in summary_calls)
    assert saved["content"] == "Final consolidated summary."


def test_summarize_documents_trusts_existing_pdf_text_layer(monkeypatch, tmp_path):
    captured_configs = []

    def fake_init(self, ocr_config=None, logger=None):
        captured_configs.append(ocr_config)

    def fake_extract(self, path, progress_callback=None):
        return SimpleNamespace(
            success=True,
            text="Document text extracted from an existing OCR layer.",
            char_count=52,
            page_count=2,
            ocr_pages=[],
            ocr_percentage=0.0,
            error=None,
        )

    def fake_call(self, prompt, text, task_type=None, agent_id=None, pass_name=None, **kwargs):
        return "Summary."

    def fake_save(output_path, content, title, logger):
        return output_path

    monkeypatch.setattr(summarize.DocumentProcessor, "__init__", fake_init)
    monkeypatch.setattr(summarize.DocumentProcessor, "extract_with_dynamic_ocr", fake_extract)
    monkeypatch.setattr(summarize.LLMCaller, "call", fake_call)
    monkeypatch.setattr(summarize, "safe_append_to_docx", fake_save)

    input_path = tmp_path / "Text Layered.pdf"
    input_path.write_bytes(b"%PDF-1.4\n")
    logger = summarize.AgentLogger("SummarizeTest", log_to_file=False)

    ok = summarize.process_document(
        str(input_path),
        logger,
        output_path_override=str(tmp_path / "AI_OUTPUT.docx"),
    )

    assert ok is True
    assert captured_configs
    assert getattr(captured_configs[0], "skip_sparse_ocr_in_text_layer_pdf", False) is True
