"""Head-to-head: gemini-3-pro-preview vs gemini-3.1-pro-preview on the summarize.py
prompt + Edinger IME text. Three runs each, alternating, to spread API conditions.
"""
import os, sys, time, statistics

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

from icharlotte_core.document_processor import DocumentProcessor
from icharlotte_core.llm_config import LLMCaller
from icharlotte_core.agent_logger import AgentLogger

PDF = r"Z:\Shared\Current Clients\5800 - AMTRUST\072 - Forney\EXPERTS\Reports\Plf's Experts\Edinger IME.pdf"
PROMPT_FILE = os.path.join(ROOT, "Scripts", "SUMMARIZE_PROMPT.txt")
MODELS = ["gemini-3-pro-preview", "gemini-3.1-pro-preview"]
RUNS = 3

def main():
    logger = AgentLogger("BenchSum", file_number="5800.072", log_to_file=False)
    proc = DocumentProcessor(logger=logger)
    res = proc.extract_with_dynamic_ocr(PDF)
    text = res.text
    print(f"Input: {len(text)} chars, {res.page_count} pages")
    with open(PROMPT_FILE, "r", encoding="utf-8") as f:
        prompt = f.read()
    llm = LLMCaller(logger=logger)

    # Alternate models across runs to spread API conditions
    runs = []
    for i in range(RUNS):
        for m in MODELS:
            print(f"\n--- Run {i+1}/{RUNS}: {m} ---")
            t0 = time.time()
            out = llm.call(prompt, text, task_type="summary", model_override=m)
            elapsed = time.time() - t0
            chars = len(out) if out else 0
            print(f"  {elapsed:.2f}s   {chars} chars   ok={bool(out)}")
            runs.append((m, elapsed, chars))

    print("\n\n========== SUMMARY ==========")
    for m in MODELS:
        elapsed_list = [e for mm, e, _ in runs if mm == m]
        chars_list   = [c for mm, _, c in runs if mm == m]
        print(f"\n{m}")
        print(f"  Times:  {[f'{e:.1f}s' for e in elapsed_list]}")
        print(f"  Mean:   {statistics.mean(elapsed_list):.2f}s   "
              f"Min: {min(elapsed_list):.2f}s   Max: {max(elapsed_list):.2f}s")
        print(f"  Chars:  {chars_list}   "
              f"(mean {statistics.mean(chars_list):.0f})")

if __name__ == "__main__":
    main()
