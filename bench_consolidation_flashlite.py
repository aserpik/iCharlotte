"""Run consolidation on the existing pre-consolidation snapshots using gemini-3.1-flash-lite-preview."""

import os
import sys
import time
import glob
from docx import Document
from docx.shared import Pt

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "Scripts"))

from Scripts import summarize_discovery as sd
from icharlotte_core.agent_logger import AgentLogger
from icharlotte_core.llm_config import LLMCaller
from icharlotte_core.document_processor import DocumentProcessor

BENCH_DIR = os.path.join(ROOT, "logs", "bench_consolidation")
MODEL = "gemini-3.1-flash-lite-preview"

SNAPSHOT_MAP = {
    "SPEC ROGS": "pre_Discovery_Responses_Defendant_Philip_Michael_Houlihan_dba_Creekside_Group_Properties.docx",
    "FROGS":     "pre_Discovery_Responses_Defendant_Philip_Michael_Houlihan_an_individual_doing_business_as_Creekside_Group_Properties.docx",
}


def consolidate(snap_path, out_path, model, logger):
    proc = DocumentProcessor(logger=logger)
    extract = proc.extract_text(snap_path)
    if not extract.success:
        return None, 0.0
    with open(sd.CONSOLIDATE_PROMPT_FILE, "r", encoding="utf-8") as f:
        prompt = f.read()
    llm = LLMCaller(logger=logger)
    t0 = time.time()
    text = llm.call(prompt, extract.text, task_type="summary", model_override=model)
    elapsed = time.time() - t0
    if not text:
        return None, elapsed
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.0
    sd.add_markdown_to_doc(doc, text)
    doc.save(out_path)
    return text, elapsed


def main():
    logger = AgentLogger("Bench-FlashLite", file_number="5800.048")
    results = []
    for label, fname in SNAPSHOT_MAP.items():
        snap = os.path.join(BENCH_DIR, fname)
        out  = os.path.join(BENCH_DIR, f"consolidated_{MODEL.replace('.','_')}__{label.replace(' ','_')}.docx")
        proc = DocumentProcessor(logger=logger)
        input_chars = len(proc.extract_text(snap).text)
        print(f"\n=== {label} ({input_chars} chars) ===")
        text, elapsed = consolidate(snap, out, MODEL, logger)
        chars = len(text) if text else 0
        print(f"  {MODEL}: {elapsed:.2f}s   {chars} chars  -> {os.path.basename(out)}")
        results.append((label, input_chars, elapsed, chars, out))

    print("\n========== Flash-Lite RESULTS ==========")
    for label, inc, el, oc, p in results:
        print(f"  {label:<12} input={inc:>6} chars   elapsed={el:>6.2f}s   output={oc:>6} chars")


if __name__ == "__main__":
    main()
