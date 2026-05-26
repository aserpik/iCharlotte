"""
Compare two model configurations for summarize_deposition:
  Run A: topic discovery = gemini-3.1-flash-lite-preview, summary = gemini-3.1-pro-preview
  Run B: topic discovery = gemini-3.1-pro-preview,        summary = gemini-3.1-pro-preview

Text extraction happens once (deterministic) and is shared by both runs.
Outputs (timing + topics + summary) are written to logs/bench_depo_models/.
"""

import os
import sys
import time
import json
from pathlib import Path

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "Scripts"))

from icharlotte_core.agent_logger import AgentLogger
from icharlotte_core.document_processor import DocumentProcessor, OCRConfig
from icharlotte_core.llm_config import LLMCaller
from Scripts import summarize_deposition as sd


INPUT_PATH = r"Z:\Shared\Current Clients\5800 - AMTRUST\072 - Forney\DISCOVERY\TRANSCRIPTS\Combined Transcripts\Forney - Pamela Forney Depo 09-24-25 - Full_165765512_1.PDF"

BENCH_DIR = Path(ROOT) / "logs" / "bench_depo_models"
BENCH_DIR.mkdir(parents=True, exist_ok=True)

RUNS = [
    {"label": "A_flashlite_topics_pro_summary",
     "topic_model": "gemini-3.1-flash-lite-preview",
     "summary_model": "gemini-3.1-pro-preview"},
    {"label": "B_pro_topics_pro_summary",
     "topic_model": "gemini-3.1-pro-preview",
     "summary_model": "gemini-3.1-pro-preview"},
]


def fmt(seconds: float) -> str:
    if seconds < 60:
        return f"{seconds:6.2f}s"
    m, s = divmod(seconds, 60)
    return f"{int(m):d}m {s:05.2f}s"


def run_pipeline(text: str, topic_model: str, summary_model: str, logger, label: str):
    """Run topic discovery + narrative summary with explicit model overrides.

    Returns a dict with timings, topics, and summary text.
    """
    llm = LLMCaller(logger=logger)
    out = {"label": label, "topic_model": topic_model, "summary_model": summary_model}

    # --- Step 2: topic discovery ---
    with open(sd.TOPIC_DISCOVERY_PROMPT_FILE, "r", encoding="utf-8") as f:
        topic_prompt = f.read()
    print(f"\n[{label}] Topic discovery via {topic_model}...", flush=True)
    t0 = time.perf_counter()
    topic_response = llm.call(topic_prompt, text, task_type="summary", model_override=topic_model)
    out["topic_elapsed"] = time.perf_counter() - t0
    topics = sd._parse_topic_response(topic_response)
    out["topics"] = topics
    print(f"  -> {len(topics)} topics in {fmt(out['topic_elapsed'])}", flush=True)

    # --- Step 4: narrative summary ---
    with open(sd.NARRATIVE_PROMPT_FILE, "r", encoding="utf-8") as f:
        base_prompt = f.read()
    deponent_name = sd.DeponentExtractor.extract_deponent_name(text) or "Deponent"
    deponent_type = sd.DeponentExtractor.detect_deponent_type(text, deponent_name)
    cfg = {"audience": "neutral", "tone": "recitation"}
    audience_directive = sd._resolve_audience_directive(cfg)
    tone_directive = sd._resolve_tone_directive(cfg)
    topic_titles = [t["title"] for t in topics]
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
    print(f"[{label}] Narrative summary via {summary_model}...", flush=True)
    t0 = time.perf_counter()
    summary = llm.call(summary_prompt, text, task_type="summary", model_override=summary_model)
    out["summary_elapsed"] = time.perf_counter() - t0
    out["summary"] = summary
    out["summary_chars"] = len(summary) if summary else 0
    print(f"  -> {out['summary_chars']:,} chars in {fmt(out['summary_elapsed'])}", flush=True)
    return out


def main():
    if not os.path.exists(INPUT_PATH):
        print(f"Input not found: {INPUT_PATH}", flush=True)
        sys.exit(1)

    logger = AgentLogger("DepoModelBench", file_number=sd.extract_file_number(INPUT_PATH))

    # --- One-time text extraction (shared by both runs) ---
    print("Extracting text (shared)...", flush=True)
    t0 = time.perf_counter()
    processor = DocumentProcessor(ocr_config=OCRConfig(adaptive=True), logger=logger)
    result = processor.extract_with_dynamic_ocr(INPUT_PATH)
    if not result.success:
        print(f"FAILED: {result.error}", flush=True)
        sys.exit(2)
    text = result.text
    extract_elapsed = time.perf_counter() - t0
    print(f"  -> {len(text):,} chars in {fmt(extract_elapsed)}", flush=True)

    # --- Run both pipelines ---
    results = []
    for cfg in RUNS:
        r = run_pipeline(text, cfg["topic_model"], cfg["summary_model"], logger, cfg["label"])
        results.append(r)

    # --- Persist outputs ---
    for r in results:
        base = BENCH_DIR / r["label"]
        (base.with_suffix(".topics.json")).write_text(
            json.dumps(r["topics"], indent=2), encoding="utf-8")
        (base.with_suffix(".summary.md")).write_text(r["summary"] or "", encoding="utf-8")

    # --- Timing comparison ---
    print("\n" + "=" * 78, flush=True)
    print("TIMING COMPARISON", flush=True)
    print("=" * 78, flush=True)
    print(f"{'Run':<40} {'Step 2':>12} {'Step 4':>12} {'Total':>12}", flush=True)
    print("-" * 78, flush=True)
    for r in results:
        total = r["topic_elapsed"] + r["summary_elapsed"]
        print(f"{r['label']:<40} {fmt(r['topic_elapsed']):>12} "
              f"{fmt(r['summary_elapsed']):>12} {fmt(total):>12}", flush=True)
    print("=" * 78, flush=True)

    # --- Output-shape comparison ---
    print("\nOUTPUT SHAPE", flush=True)
    print("-" * 78, flush=True)
    for r in results:
        print(f"{r['label']}:", flush=True)
        print(f"  topics:        {len(r['topics'])}", flush=True)
        print(f"  summary chars: {r['summary_chars']:,}", flush=True)
        if r["summary"]:
            bullets = sum(1 for line in r["summary"].splitlines()
                          if line.strip().startswith(("*", "-", "•")))
            headings = sum(1 for line in r["summary"].splitlines()
                           if line.strip().startswith("#"))
            print(f"  headings:      {headings}", flush=True)
            print(f"  bullets:       {bullets}", flush=True)

    # --- Topic title comparison ---
    print("\nTOPIC TITLES (side by side, in original rank order)", flush=True)
    print("-" * 78, flush=True)
    a_titles = [t["title"] for t in results[0]["topics"]]
    b_titles = [t["title"] for t in results[1]["topics"]]
    a_set = set(a_titles)
    b_set = set(b_titles)
    common = a_set & b_set
    only_a = a_set - b_set
    only_b = b_set - a_set
    width = max(len(t) for t in (a_titles + b_titles)) + 2
    for i in range(max(len(a_titles), len(b_titles))):
        a = a_titles[i] if i < len(a_titles) else ""
        b = b_titles[i] if i < len(b_titles) else ""
        print(f"  {i+1:>2}. {a:<{width}}| {b}", flush=True)
    print(f"\n  identical titles: {len(common)} / "
          f"max({len(a_titles)},{len(b_titles)}) = {max(len(a_titles), len(b_titles))}", flush=True)
    if only_a:
        print(f"  only in A: {sorted(only_a)}", flush=True)
    if only_b:
        print(f"  only in B: {sorted(only_b)}", flush=True)

    print(f"\nOutputs saved to: {BENCH_DIR}", flush=True)


if __name__ == "__main__":
    main()
