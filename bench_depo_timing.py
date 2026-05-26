"""
Time each step of summarize_deposition end-to-end on a single PDF.

Mimics the UI's default config (all topics, 5 bullets, neutral bias,
cross-check off, no context docs) so the run is realistic.
"""

import os
import sys
import time
from pathlib import Path

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "Scripts"))

from icharlotte_core.agent_logger import AgentLogger
from icharlotte_core.document_processor import DocumentProcessor, OCRConfig
from icharlotte_core.llm_config import LLMCaller
from icharlotte_core.deposition import session_manager
from Scripts import summarize_deposition as sd


INPUT_PATH = r"Z:\Shared\Current Clients\5800 - AMTRUST\072 - Forney\DISCOVERY\TRANSCRIPTS\Combined Transcripts\Forney - Pamela Forney Depo 09-24-25 - Full_165765512_1.PDF"


def fmt(seconds: float) -> str:
    if seconds < 60:
        return f"{seconds:6.2f}s"
    m, s = divmod(seconds, 60)
    return f"{int(m):d}m {s:05.2f}s"


def main():
    if not os.path.exists(INPUT_PATH):
        print(f"Input not found: {INPUT_PATH}", flush=True)
        sys.exit(1)

    file_number = sd.extract_file_number(INPUT_PATH)
    logger = AgentLogger("Deposition-Bench", file_number=file_number)
    llm = LLMCaller(logger=logger)

    timings = []
    t_overall = time.perf_counter()

    # -----------------------------------------------------------------------
    # Step 1: Text extraction (PDF -> text, adaptive OCR)
    # -----------------------------------------------------------------------
    print("\n[1/5] Extracting text...", flush=True)
    t0 = time.perf_counter()
    processor = DocumentProcessor(ocr_config=OCRConfig(adaptive=True), logger=logger)
    result = processor.extract_with_dynamic_ocr(INPUT_PATH)
    if not result.success:
        print(f"  FAILED: {result.error}", flush=True)
        sys.exit(2)
    text = result.text
    extract_elapsed = time.perf_counter() - t0
    timings.append(("Text extraction (PDF+OCR)", extract_elapsed))
    print(f"  -> {len(text):,} chars in {fmt(extract_elapsed)}", flush=True)

    # Cache text to a session file (same as Phase 1 does)
    paths = session_manager.compute_session_paths(INPUT_PATH)
    paths.cached_text_path.write_text(text, encoding="utf-8")

    # -----------------------------------------------------------------------
    # Step 2: Topic discovery LLM call
    # -----------------------------------------------------------------------
    print("\n[2/5] Topic discovery (LLM)...", flush=True)
    with open(sd.TOPIC_DISCOVERY_PROMPT_FILE, "r", encoding="utf-8") as f:
        topic_prompt = f.read()
    t0 = time.perf_counter()
    topic_response = llm.call(topic_prompt, text, task_type="summary")
    topic_elapsed = time.perf_counter() - t0
    topics = sd._parse_topic_response(topic_response)
    timings.append(("Topic discovery LLM", topic_elapsed))
    print(f"  -> {len(topics)} topics in {fmt(topic_elapsed)}", flush=True)
    for t in topics:
        print(f"     {t['rank']:>2}. {t['title']}", flush=True)

    # -----------------------------------------------------------------------
    # Step 3: Deponent metadata extraction (cheap regex, included for parity)
    # -----------------------------------------------------------------------
    print("\n[3/5] Deponent metadata...", flush=True)
    t0 = time.perf_counter()
    deponent_name = sd.DeponentExtractor.extract_deponent_name(text) or \
        os.path.splitext(os.path.basename(INPUT_PATH))[0]
    deposition_date = sd.DeponentExtractor.extract_deposition_date(text)
    deponent_type = sd.DeponentExtractor.detect_deponent_type(text, deponent_name)
    meta_elapsed = time.perf_counter() - t0
    timings.append(("Deponent metadata (regex)", meta_elapsed))
    print(f"  -> Name: {deponent_name!r}, Date: {deposition_date!r}, Type: {deponent_type!r} "
          f"in {fmt(meta_elapsed)}", flush=True)

    # -----------------------------------------------------------------------
    # Step 4: Narrative summary LLM call (topic-locked)
    # -----------------------------------------------------------------------
    print("\n[4/5] Narrative summary (LLM)...", flush=True)
    with open(sd.NARRATIVE_PROMPT_FILE, "r", encoding="utf-8") as f:
        base_prompt = f.read()

    topic_titles = [t["title"] for t in topics]
    cfg = {"audience": "neutral", "tone": "recitation"}
    audience_directive = sd._resolve_audience_directive(cfg)
    tone_directive = sd._resolve_tone_directive(cfg)
    summary_prompt = sd._build_topic_locked_prompt(
        base_prompt,
        topic_list=topic_titles,
        bullets_per_topic=5,
        deponent_label=deponent_type or "Deponent",
        custom_rules="",
        audience_directive=audience_directive,
        tone_directive=tone_directive,
        context_documents=[],
    )
    t0 = time.perf_counter()
    summary = llm.call(summary_prompt, text, task_type="summary")
    summary_elapsed = time.perf_counter() - t0
    timings.append(("Narrative summary LLM", summary_elapsed))
    print(f"  -> {len(summary):,} chars in {fmt(summary_elapsed)}", flush=True)

    # -----------------------------------------------------------------------
    # Step 5: Save to .docx (real output dir under case NOTES/AI OUTPUT)
    # -----------------------------------------------------------------------
    print("\n[5/5] Saving .docx...", flush=True)
    output_dir = sd.get_output_directory(INPUT_PATH)
    output_file = os.path.join(output_dir, "Deposition_summaries.docx")
    os.makedirs(output_dir, exist_ok=True)
    t0 = time.perf_counter()
    ok = sd.save_to_docx(summary, output_file, deponent_name, deposition_date, logger)
    save_elapsed = time.perf_counter() - t0
    timings.append(("Save to .docx", save_elapsed))
    if not ok:
        print("  SAVE FAILED", flush=True)
        sys.exit(3)
    print(f"  -> {output_file} in {fmt(save_elapsed)}", flush=True)

    # Cleanup the bench session files (cached text + .json never written)
    try:
        if paths.cached_text_path.exists():
            paths.cached_text_path.unlink()
    except OSError:
        pass

    # -----------------------------------------------------------------------
    # Timing summary
    # -----------------------------------------------------------------------
    overall = time.perf_counter() - t_overall
    print("\n" + "=" * 60, flush=True)
    print("TIMING SUMMARY", flush=True)
    print("=" * 60, flush=True)
    print(f"{'Step':<35} {'Time':>12} {'% of total':>10}", flush=True)
    print("-" * 60, flush=True)
    for name, secs in timings:
        pct = (secs / overall) * 100 if overall > 0 else 0
        print(f"{name:<35} {fmt(secs):>12} {pct:>9.1f}%", flush=True)
    print("-" * 60, flush=True)
    print(f"{'TOTAL':<35} {fmt(overall):>12}", flush=True)
    print("=" * 60, flush=True)
    print(f"\nOutput: {output_file}", flush=True)


if __name__ == "__main__":
    main()
