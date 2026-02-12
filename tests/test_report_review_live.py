"""
Live test of Report Review Engine on an associate's draft report.

Runs both passes:
  Pass 1: Structural scan (formatting, completeness)
  Pass 2: LLM review with Track Changes applied to the document

Usage:
    python -m tests.test_report_review_live "path/to/report.docx"
"""

import os
import sys
import time
import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("test_report_review_live")


def run_review(doc_path: str, only_sections=None, skip_structural=False, skip_final_validation=False):
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()

    if not os.path.exists(doc_path):
        print(f"ERROR: File not found: {doc_path}")
        return False

    print(f"Opening: {doc_path}")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True

    # Close any existing copy of this document without saving
    doc_name = os.path.basename(doc_path)
    try:
        for i in range(word.Documents.Count, 0, -1):
            d = word.Documents(i)
            if d.Name == doc_name:
                print(f"  Closing existing copy of {doc_name} (no save)...")
                d.Close(0)  # wdDoNotSaveChanges
    except Exception as e:
        print(f"  Warning closing existing: {e}")

    # Open fresh
    doc = word.Documents.Open(os.path.abspath(doc_path))
    print(f"  Opened: {doc.Name}, {doc.Paragraphs.Count} paragraphs")

    if only_sections:
        print(f"  Filtering to sections: {only_sections}")

    try:
        from icharlotte_core.report_reviewer import ReportReviewer
        import re

        # Detect case path from document path
        case_path = None
        parts = doc_path.replace('/', '\\').split('\\')
        for i, part in enumerate(parts):
            if re.search(r'\d{3,4}\s*-\s*', part):
                case_path = '\\'.join(parts[:i + 1])
                break

        print(f"Case path: {case_path or 'not detected'}")

        reviewer = ReportReviewer(doc_com=doc, case_path=case_path)

        def on_progress(msg):
            print(f"  >> {msg}")

        print("\n" + "=" * 60)
        print("STARTING REPORT REVIEW")
        print("=" * 60)

        t0 = time.time()
        structural, total_changes = reviewer.run(
            selection_range=None,
            include_case_data=bool(case_path),
            extra_doc_paths=None,
            progress_callback=on_progress,
            only_sections=only_sections,
            skip_structural=skip_structural,
            skip_final_validation=skip_final_validation,
        )
        elapsed = time.time() - t0

        print("\n" + "=" * 60)
        print(f"REVIEW COMPLETE — {elapsed:.1f}s")
        print(f"Total tracked changes applied: {total_changes}")
        if structural:
            errors = [f for f in structural.findings if f.severity == "ERROR"]
            warns = [f for f in structural.findings if f.severity == "WARN"]
            print(f"Structural findings: {len(errors)} errors, {len(warns)} warnings")
        print("=" * 60)
        print("\nThe document is open in Word with Track Changes.")
        print("Review the changes, then Accept/Reject as needed.")

        return True

    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Live report review test")
    parser.add_argument("doc_path", help="Path to report .docx")
    parser.add_argument("--sections", "-s", nargs="+",
                        help="Only review these sections (partial match, e.g. 'DISCOVERY' 'LIABILITY')")
    parser.add_argument("--skip-structural", action="store_true",
                        help="Skip Pass 1 structural scan")
    args = parser.parse_args()

    ok = run_review(args.doc_path, only_sections=args.sections,
                    skip_structural=args.skip_structural)
    sys.exit(0 if ok else 1)
