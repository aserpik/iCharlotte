"""
Test script to verify the contradiction detection matching logic works correctly.
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from detect_contradictions import _matches_selection

# Test cases: (name_to_match, selected_summaries, expected_result)
TEST_CASES = [
    # Document Registry names vs CaseDataManager variable names
    ("depo_summary_ron_dadash", ["Deposition of Plaintiff Ron Dadash"], True),
    ("depo_extraction_ron_dadash", ["Deposition of Plaintiff Ron Dadash"], True),
    ("discovery_summary_2025_09_30_r_dudash_resp_to_frog_1_", ["2025-09-30 R.Dudash Resp to FROG(1)"], True),
    ("discovery_summary_2025_11_03_r_dudash_resp_to_srog_1_", ["Plaintiff Ron Dudash's Responses to Special Interrogatories"], True),
    ("summary_judgment_against_dudash", ["Judgment against Dudash"], True),

    # Exact matches
    ("Judgment against Dudash.pdf", ["Judgment against Dudash.pdf"], True),
    ("Survey.pdf", ["Survey.pdf"], True),

    # Case insensitive
    ("DEPO_SUMMARY_RON_DADASH", ["deposition of plaintiff ron dadash"], True),

    # Substring matches
    ("discovery_summary_frog", ["frog"], True),
    ("depo_summary_test", ["depo_summary_test"], True),

    # Should NOT match - different documents
    ("depo_summary_ron_dadash", ["Judgment against Dudash"], False),
    ("discovery_summary_frog", ["Deposition of Plaintiff Ron Dadash"], False),
    ("summary_unrelated_doc", ["Deposition of Plaintiff Ron Dadash"], False),

    # None means all summaries
    ("any_summary_name", None, True),

    # Multiple selections - should match if ANY matches
    ("depo_summary_ron_dadash", ["Judgment against Dudash", "Deposition of Plaintiff Ron Dadash"], True),

    # CRITICAL: FROG should NOT match SROG selection (different discovery types)
    ("discovery_summary_2025_09_30_r_dudash_resp_to_frog_1_", ["Plaintiff Ron Dudash's Responses to Special Interrogatories"], False),

    # SROG should NOT match FROG selection
    ("discovery_summary_2025_11_03_r_dudash_resp_to_srog_1_", ["2025-09-30 R.Dudash Resp to FROG(1)"], False),
]

def run_tests():
    """Run all test cases and report results."""
    passed = 0
    failed = 0

    print("Testing _matches_selection function")
    print("=" * 70)

    for name, selections, expected in TEST_CASES:
        result = _matches_selection(name, selections)
        status = "PASS" if result == expected else "FAIL"

        if result == expected:
            passed += 1
        else:
            failed += 1

        sel_str = str(selections)[:50] + "..." if len(str(selections)) > 50 else str(selections)
        print(f"[{status}] name='{name[:40]}...' selections={sel_str}")
        if result != expected:
            print(f"       Expected: {expected}, Got: {result}")

    print("=" * 70)
    print(f"Results: {passed} passed, {failed} failed out of {len(TEST_CASES)} tests")

    return failed == 0

if __name__ == "__main__":
    success = run_tests()
    sys.exit(0 if success else 1)
