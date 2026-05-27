"""Dump text from all four consolidated outputs side-by-side for comparison."""

import os, sys
ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
from icharlotte_core.document_processor import extract_docx_text

BENCH_DIR = os.path.join(ROOT, "logs", "bench_consolidation")

PAIRS = {
    "SPEC ROGS": {
        "pro":   "consolidated_gemini-3_1-pro-preview__2025 11 03 Def Phillip M Houlihans Rsp to SPEC ROGS Set 1 By Plt Patricia Campi.docx",
        "flash": "consolidated_gemini-3_1-flash-lite-preview__SPEC_ROGS.docx",
    },
    "FROGS": {
        "pro":   "consolidated_gemini-3_1-pro-preview__2025 11 03 Def Phillip M Houlihans Rsp to FROGS Set 1 By Plt Patricia Campi.docx",
        "flash": "consolidated_gemini-3_1-flash-lite-preview__FROGS.docx",
    },
}

out_path = os.path.join(BENCH_DIR, "side_by_side.txt")
with open(out_path, "w", encoding="utf-8") as f:
    for label, m in PAIRS.items():
        for tier, fn in m.items():
            p = os.path.join(BENCH_DIR, fn)
            text = extract_docx_text(p)
            f.write("=" * 80 + "\n")
            f.write(f"=== {label} | {tier.upper()} | {len(text)} chars ===\n")
            f.write("=" * 80 + "\n")
            f.write(text + "\n\n")
print(f"Wrote {out_path}")
