from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[2]
SCRIPTS = ROOT / "Scripts"
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

import Scripts.summarize_discovery as sd  # noqa: E402


def test_verbatim_extracted_responses_preserve_complete_source_response():
    source_text = """
SPECIAL INTERROGATORY NO. 7:
Identify all injuries you claim resulted from the incident.

RESPONSE TO SPECIAL INTERROGATORY NO. 7:
Subject to and without waiving the foregoing objections, Responding Party states as follows: Plaintiff sustained injuries to her cervical spine, lumbar spine, and left shoulder, experienced headaches three to four times per week for approximately six months, treated with Harbor Physical Therapy from March 4, 2024 through June 18, 2024, and continues to experience intermittent numbness in her left hand when lifting objects heavier than ten pounds.

SPECIAL INTERROGATORY NO. 8:
Identify all witnesses.

RESPONSE TO SPECIAL INTERROGATORY NO. 8:
Responding Party identifies Maria Lopez, Allen Smith, and Officer Janet Reed as witnesses.
"""

    section_a = sd.build_verbatim_extracted_responses(
        source_text,
        fallback_discovery_type="Special Interrogatories",
    )

    assert (
        "SROG No. 7: Subject to and without waiving the foregoing objections, "
        "Responding Party states as follows: Plaintiff sustained injuries to "
        "her cervical spine, lumbar spine, and left shoulder, experienced "
        "headaches three to four times per week for approximately six months, "
        "treated with Harbor Physical Therapy from March 4, 2024 through "
        "June 18, 2024, and continues to experience intermittent numbness in "
        "her left hand when lifting objects heavier than ten pounds."
    ) in section_a
    assert "Identify all injuries you claim resulted from the incident" not in section_a
    assert "SROG No. 8: Responding Party identifies Maria Lopez" in section_a


def test_consolidation_replaces_shortened_section_a_with_verbatim_section():
    original_doc_text = """
1. Plaintiff's Responses to Special Interrogatories

SECTION A: EXTRACTED RESPONSES
SROG No. 7: Responding Party states as follows: Plaintiff sustained injuries to her cervical spine, lumbar spine, and left shoulder, experienced headaches three to four times per week for approximately six months, treated with Harbor Physical Therapy from March 4, 2024 through June 18, 2024, and continues to experience intermittent numbness in her left hand when lifting objects heavier than ten pounds.

SECTION B: NARRATIVE SUMMARY
Plaintiff claims neck, back, and shoulder injuries.
"""
    llm_consolidated = """
**Master Discovery Summary for Plaintiff**

SECTION A: CONSOLIDATED EXTRACTED RESPONSES
SROG No. 7: Plaintiff claims neck, back, shoulder, and headache injuries.

SECTION B: CONSOLIDATED NARRATIVE SUMMARY
**Claimed Injuries.** Plaintiff claims injuries.
"""

    verbatim_section = sd.collect_consolidated_extracted_responses(original_doc_text)
    protected = sd.replace_consolidated_extracted_responses(
        llm_consolidated,
        verbatim_section,
    )

    assert "Plaintiff claims neck, back, shoulder, and headache injuries." not in protected
    assert (
        "SROG No. 7: Responding Party states as follows: Plaintiff sustained "
        "injuries to her cervical spine, lumbar spine, and left shoulder, "
        "experienced headaches three to four times per week for approximately "
        "six months, treated with Harbor Physical Therapy from March 4, 2024 "
        "through June 18, 2024, and continues to experience intermittent "
        "numbness in her left hand when lifting objects heavier than ten pounds."
    ) in protected
    assert "SECTION B: CONSOLIDATED NARRATIVE SUMMARY" in protected


def test_process_document_saves_verbatim_section_a_when_llm_extraction_is_shortened(
    tmp_path,
    monkeypatch,
):
    source_text = """
SPECIAL INTERROGATORY NO. 7:
Identify all injuries you claim resulted from the incident.

RESPONSE TO SPECIAL INTERROGATORY NO. 7:
Responding Party states as follows: Plaintiff sustained injuries to her cervical spine, lumbar spine, and left shoulder, experienced headaches three to four times per week for approximately six months, treated with Harbor Physical Therapy from March 4, 2024 through June 18, 2024, and continues to experience intermittent numbness in her left hand when lifting objects heavier than ten pounds.
"""
    input_path = tmp_path / "responses.pdf"
    input_path.write_bytes(b"%PDF-1.4\n%fake")
    saved = {}

    class FakeProcessor:
        def __init__(self, *args, **kwargs):
            pass

        def extract_with_dynamic_ocr(self, path):
            return SimpleNamespace(
                success=True,
                text=source_text,
                char_count=len(source_text),
                page_count=1,
            )

    class FakeCaller:
        def __init__(self, *args, **kwargs):
            pass

        def call(self, prompt, text, **kwargs):
            pass_name = kwargs.get("pass_name")
            if pass_name == "extraction":
                return (
                    "RESPONDING_PARTY: Plaintiff Jane Doe\n"
                    "DISCOVERY_TYPE: Special Interrogatories\n"
                    "SROG No. 7: Plaintiff claims neck, back, shoulder, "
                    "headache, treatment, and numbness issues."
                )
            if pass_name == "summary":
                return (
                    "RESPONDING_PARTY: Plaintiff Jane Doe\n"
                    "DISCOVERY_TYPE: Special Interrogatories\n"
                    "We have now received Plaintiff's responses."
                )
            if pass_name == "cross_check":
                return "We have now received Plaintiff's responses."
            return "We have now received Plaintiff's responses."

    class FakeLogger:
        def info(self, *args, **kwargs):
            pass

        def warning(self, *args, **kwargs):
            pass

        def error(self, *args, **kwargs):
            pass

        def progress(self, *args, **kwargs):
            pass

        def pass_start(self, *args, **kwargs):
            pass

        def pass_complete(self, *args, **kwargs):
            pass

        def pass_failed(self, *args, **kwargs):
            pass

        def output_file(self, *args, **kwargs):
            pass

    def fake_save(extraction_content, summary_content, output_path, *args, **kwargs):
        saved["extraction_content"] = extraction_content
        saved["summary_content"] = summary_content
        saved["output_path"] = output_path
        return True

    monkeypatch.setattr(sd, "DocumentProcessor", FakeProcessor)
    monkeypatch.setattr(sd, "LLMCaller", FakeCaller)
    monkeypatch.setattr(sd, "save_to_docx", fake_save)
    monkeypatch.setattr(sd, "consolidate_file", lambda *args, **kwargs: True)

    ok = sd.process_document(
        str(input_path),
        FakeLogger(),
        output_path_override=str(tmp_path / "out.docx"),
    )

    assert ok is True
    assert "Plaintiff claims neck, back, shoulder" not in saved["extraction_content"]
    assert (
        "SROG No. 7: Responding Party states as follows: Plaintiff sustained "
        "injuries to her cervical spine, lumbar spine, and left shoulder, "
        "experienced headaches three to four times per week for approximately "
        "six months, treated with Harbor Physical Therapy from March 4, 2024 "
        "through June 18, 2024, and continues to experience intermittent "
        "numbness in her left hand when lifting objects heavier than ten pounds."
    ) in saved["extraction_content"]
