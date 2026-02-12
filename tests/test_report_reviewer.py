"""
Test the Report Review Engine — Pass 1 structural scan + prompt building.

Runs against a gold standard report via COM to verify:
1. Section detection (fuzzy matching)
2. Structural scan findings
3. Prompt building (no LLM call)
4. ReportReviewDialog instantiation

Usage:
    python -m tests.test_report_reviewer
"""

import os
import sys
import time

# ── Test 1: fuzzy_match_section_heading ──

def test_fuzzy_match():
    from icharlotte_core.word_validator import fuzzy_match_section_heading

    cases = [
        ("FACTUAL BACKGROUND", "FACTUAL BACKGROUND"),
        ("FACTUAL BACKGROUND AND SUMMARY", "FACTUAL BACKGROUND"),
        ("PROCEDURAL STATUS", "PROCEDURAL STATUS"),
        ("SETTLEMENT", "SETTLEMENT STATUS"),
        ("SETTLEMENT STATUS", "SETTLEMENT STATUS"),
        ("DISCOVERY", "DISCOVERY"),
        ("EXPERTS", "EXPERTS"),
        ("EVALUATION OF LIABILITY", "EVALUATION OF LIABILITY"),
        ("MEDICAL RECORD REVIEW", "MEDICAL RECORD REVIEW"),
        ("FURTHER CASE HANDLING", "FURTHER CASE HANDLING"),
        ("INVESTIGATION", "INVESTIGATION"),
        ("EVALUATION OF EXPOSURE", "EVALUATION OF EXPOSURE"),
        # Should NOT match
        ("THIS IS A RANDOM HEADING", None),
        ("VERY TRULY YOURS", None),
        ("RE: SMITH VS JONES", None),
    ]

    passed = 0
    failed = 0
    for input_text, expected in cases:
        result = fuzzy_match_section_heading(input_text)
        if result == expected:
            passed += 1
            print(f"  PASS: '{input_text}' -> '{result}'")
        else:
            failed += 1
            print(f"  FAIL: '{input_text}' -> '{result}' (expected '{expected}')")

    print(f"\nFuzzy match: {passed} passed, {failed} failed")
    return failed == 0


# ── Test 2: Section detection + structural scan on a real document ──

def test_structural_scan():
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print("  SKIP: win32com not available")
        return True

    pythoncom.CoInitialize()
    try:
        # Find a gold standard report
        test_doc = os.path.join(
            os.path.dirname(os.path.dirname(__file__)),
            "autodownloadreports", "gold_1_Haydel.docx"
        )
        if not os.path.exists(test_doc):
            print(f"  SKIP: Test document not found: {test_doc}")
            return True

        print(f"  Opening: {os.path.basename(test_doc)}")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(test_doc), ReadOnly=True)

        try:
            from icharlotte_core.report_reviewer import ReportReviewer

            reviewer = ReportReviewer(doc_com=doc)

            # Test section detection
            print("  Detecting sections...")
            sections = reviewer.detect_sections()
            print(f"  Found {len(sections)} sections:")
            for s in sections:
                content_len = len(s.text.strip())
                match_note = "" if s.raw_name == s.name else f" (raw: '{s.raw_name}')"
                print(f"    {s.name}: {content_len} chars, para {s.para_index}{match_note}")

            if len(sections) < 5:
                print("  FAIL: Expected at least 5 sections in gold standard")
                return False

            # Test structural scan
            print("\n  Running structural scan...")
            result = reviewer.run_structural_scan(sections)
            result.print_summary()

            errors = [f for f in result.findings if f.severity == "ERROR"]
            warns = [f for f in result.findings if f.severity == "WARN"]
            passes = [f for f in result.findings if f.severity == "PASS"]
            print(f"\n  Summary: {len(passes)} PASS, {len(warns)} WARN, {len(errors)} ERROR")

            # Gold standard should have most sections present
            found_sections = {s.name for s in sections}
            expected_core = {"FACTUAL BACKGROUND", "PROCEDURAL STATUS", "DISCOVERY",
                             "EVALUATION OF LIABILITY", "EVALUATION OF EXPOSURE"}
            missing_core = expected_core - found_sections
            if missing_core:
                print(f"  FAIL: Gold standard missing core sections: {missing_core}")
                return False

            print("  PASS: Structural scan completed successfully")
            return True

        finally:
            doc.Close(False)

    except Exception as e:
        print(f"  ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        pythoncom.CoUninitialize()


# ── Test 3: Prompt builder (no LLM call) ──

def test_prompt_builder():
    from icharlotte_core.report_reviewer import ReportReviewer

    # Create reviewer with no doc (just testing prompt building)
    class FakeDoc:
        class Content:
            Text = "Full document text here"
            End = 1000

    reviewer = ReportReviewer(doc_com=FakeDoc())
    reviewer._style_guide = {
        "general": {"hedging_phrases": ["We believe", "It appears that"]},
        "sections": {},
    }
    reviewer._section_examples = {
        "FACTUAL BACKGROUND": ["Example section text from gold standard..."],
    }

    prompt = reviewer._build_review_prompt(
        section_text="This is the associate's draft text for factual background.",
        section_name="FACTUAL BACKGROUND",
        full_doc_text="Full document context here...",
        case_data={"sections": {"FACTUAL BACKGROUND": "Additional case facts..."}},
        extra_texts=["Complaint text extracted from PDF..."],
    )

    # Verify prompt contains key elements
    checks = [
        ("section name", "FACTUAL BACKGROUND" in prompt),
        ("hedging phrases", "We believe" in prompt),
        ("example block", "EXAMPLE 1" in prompt),
        ("case data", "CASE FILE DATA" in prompt),
        ("extra docs", "ADDITIONAL REFERENCE" in prompt),
        ("associate draft", "associate's draft text" in prompt),
        ("redline instructions", "Track Changes" in prompt),
        ("paragraph preservation", "PRESERVE THE EXACT PARAGRAPH STRUCTURE" in prompt),
    ]

    passed = 0
    failed = 0
    for name, ok in checks:
        if ok:
            passed += 1
            print(f"  PASS: prompt contains {name}")
        else:
            failed += 1
            print(f"  FAIL: prompt missing {name}")

    print(f"\n  Prompt length: {len(prompt)} chars")
    print(f"  Prompt builder: {passed} passed, {failed} failed")
    return failed == 0


# ── Test 4: Dialog instantiation ──

def test_dialog():
    try:
        from PySide6.QtWidgets import QApplication
        app = QApplication.instance()
        if not app:
            app = QApplication(sys.argv)

        from icharlotte_core.ui.report_review_dialog import ReportReviewDialog

        # Test with no selection, no case
        dlg = ReportReviewDialog(has_selection=False, case_detected=False)
        opts = dlg.get_options()
        assert opts["scope"] == "full", f"Expected 'full', got '{opts['scope']}'"
        assert not opts["include_case_data"], "Should not include case data by default"
        assert opts["extra_doc_paths"] == [], "Should have no extra docs"
        print("  PASS: Dialog (no selection, no case)")

        # Test with selection and case
        dlg2 = ReportReviewDialog(has_selection=True, case_detected=True)
        opts2 = dlg2.get_options()
        assert opts2["scope"] == "selection", f"Expected 'selection', got '{opts2['scope']}'"
        print("  PASS: Dialog (with selection, with case)")

        return True

    except Exception as e:
        print(f"  ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False


# ── Main ──

if __name__ == "__main__":
    print("=" * 60)
    print("Report Review Engine — Test Suite")
    print("=" * 60)

    results = {}

    print("\n[1] Fuzzy Match Section Heading")
    results["fuzzy_match"] = test_fuzzy_match()

    print("\n[2] Prompt Builder")
    results["prompt_builder"] = test_prompt_builder()

    print("\n[3] Dialog Instantiation")
    results["dialog"] = test_dialog()

    print("\n[4] Structural Scan (COM + real document)")
    results["structural_scan"] = test_structural_scan()

    print("\n" + "=" * 60)
    total = len(results)
    passed = sum(1 for v in results.values() if v)
    print(f"RESULTS: {passed}/{total} tests passed")
    for name, ok in results.items():
        status = "PASS" if ok else "FAIL"
        print(f"  [{status}] {name}")
    print("=" * 60)

    sys.exit(0 if passed == total else 1)
