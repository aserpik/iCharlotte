"""
Benchmark: compare consolidation step using gemini-3.1-pro-preview vs gemini-3.1-flash-preview.

Strategy:
  1. Monkey-patch summarize_discovery.consolidate_file to snapshot the pre-consolidation
     .docx instead of mutating it.
  2. Run the full agent on each PDF (output redirected to a clean bench dir to avoid
     polluting the real AI OUTPUT).
  3. For each snapshot, run consolidation twice — once with each model override —
     measuring wall-clock time and saving the resulting .docx for side-by-side review.
"""

import os
import sys
import time
import shutil
from docx import Document
from docx.shared import Pt

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "Scripts"))

from Scripts import summarize_discovery as sd
from icharlotte_core.agent_logger import AgentLogger
from icharlotte_core.llm_config import LLMCaller
from icharlotte_core.document_processor import DocumentProcessor

PDFS = [
    r"Z:\Shared\Current Clients\5800 - AMTRUST\048 - Campi\DISCOVERY\RESPONSES\tPltf\fHoulihans\2025 11 03 Def Phillip M Houlihans Rsp to SPEC ROGS Set 1 By Plt Patricia Campi.pdf",
    r"Z:\Shared\Current Clients\5800 - AMTRUST\048 - Campi\DISCOVERY\RESPONSES\tPltf\fHoulihans\2025 11 03 Def Phillip M Houlihans Rsp to FROGS Set 1 By Plt Patricia Campi.pdf",
]

MODELS = ["gemini-3.1-pro-preview", "gemini-3.1-flash-preview"]

BENCH_DIR = os.path.join(ROOT, "logs", "bench_consolidation")
os.makedirs(BENCH_DIR, exist_ok=True)

# Per-run state captured by the monkey-patched consolidate stub
_last_snapshot = {"path": None}


def snapshot_consolidate(file_path, logger, show_diff=False):
    """Replacement for sd.consolidate_file: copies the saved .docx to a snapshot path
    and skips the real LLM consolidation. We'll do that ourselves with model overrides."""
    snap = os.path.join(BENCH_DIR, "pre_" + os.path.basename(file_path))
    shutil.copy2(file_path, snap)
    _last_snapshot["path"] = snap
    logger.info(f"[bench] snapshot saved: {snap}")
    return True


def redirect_output_dir(input_path):
    """Force all agent .docx output for the bench under BENCH_DIR/agent_out."""
    out = os.path.join(BENCH_DIR, "agent_out")
    os.makedirs(out, exist_ok=True)
    return out


def consolidate_with_model(snap_path, output_path, model, logger):
    """Run the consolidation prompt against snap_path text using a specific model.
    Returns (consolidated_text, elapsed_seconds)."""
    processor = DocumentProcessor(logger=logger)
    extract = processor.extract_text(snap_path)
    if not extract.success:
        return None, 0.0

    with open(sd.CONSOLIDATE_PROMPT_FILE, "r", encoding="utf-8") as f:
        prompt = f.read()

    llm = LLMCaller(logger=logger)
    t0 = time.time()
    consolidated = llm.call(prompt, extract.text, task_type="summary", model_override=model)
    elapsed = time.time() - t0

    if not consolidated:
        return None, elapsed

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.0
    sd.add_markdown_to_doc(doc, consolidated)
    doc.save(output_path)
    return consolidated, elapsed


def main():
    # Monkey-patch
    sd.consolidate_file = snapshot_consolidate
    sd.get_output_directory = redirect_output_dir

    results = []  # list of dicts: pdf, model, elapsed, chars, output_path

    for pdf in PDFS:
        pdf_name = os.path.basename(pdf)
        print(f"\n=== Processing: {pdf_name} ===")
        _last_snapshot["path"] = None

        file_num = sd.extract_file_number(pdf)
        logger = AgentLogger("Bench", file_number=file_num)

        t_agent_start = time.time()
        ok = sd.process_document(pdf, logger)
        agent_elapsed = time.time() - t_agent_start
        if not ok or not _last_snapshot["path"]:
            print(f"  AGENT FAILED on {pdf_name}")
            continue
        snap = _last_snapshot["path"]
        print(f"  Agent (no consolidation): {agent_elapsed:.2f}s  ->  snapshot: {os.path.basename(snap)}")

        # Read snapshot input size for context
        proc = DocumentProcessor(logger=logger)
        snap_text = proc.extract_text(snap)
        input_chars = len(snap_text.text) if snap_text.success else -1
        print(f"  Snapshot input: {input_chars} chars")

        for model in MODELS:
            tag = model.replace(".", "_")
            out_path = os.path.join(
                BENCH_DIR, f"consolidated_{tag}__{os.path.splitext(pdf_name)[0]}.docx"
            )
            print(f"  Running consolidation with {model} ...")
            text, elapsed = consolidate_with_model(snap, out_path, model, logger)
            chars = len(text) if text else 0
            print(f"    {model}: {elapsed:.2f}s   {chars} chars   -> {os.path.basename(out_path)}")
            results.append({
                "pdf": pdf_name,
                "input_chars": input_chars,
                "model": model,
                "elapsed_s": elapsed,
                "out_chars": chars,
                "output_path": out_path,
            })

    print("\n\n========== RESULTS ==========")
    print(f"{'PDF':<60} {'Model':<28} {'Elapsed':>10} {'In chars':>10} {'Out chars':>10}")
    for r in results:
        print(f"{r['pdf'][:60]:<60} {r['model']:<28} {r['elapsed_s']:>9.2f}s {r['input_chars']:>10} {r['out_chars']:>10}")

    # Write a simple manifest
    manifest_path = os.path.join(BENCH_DIR, "manifest.txt")
    with open(manifest_path, "w", encoding="utf-8") as f:
        f.write(f"{'PDF':<60} {'Model':<28} {'Elapsed':>10} {'In chars':>10} {'Out chars':>10}  Output\n")
        for r in results:
            f.write(f"{r['pdf'][:60]:<60} {r['model']:<28} {r['elapsed_s']:>9.2f}s {r['input_chars']:>10} {r['out_chars']:>10}  {r['output_path']}\n")
    print(f"\nManifest written to {manifest_path}")


if __name__ == "__main__":
    main()
