"""
Live test of the redline engine against a real Word document.
Tests whether _apply_redlines correctly produces surgical tracked changes.

Run: python tests/test_redline_live.py
"""
import sys
import os
import time
import tempfile

def test_redline():
    """Test redline with identical text + small changes to verify surgical edits."""
    try:
        import win32com.client
        from diff_match_patch import diff_match_patch
    except ImportError as e:
        print(f"Missing dependency: {e}")
        return False

    # Create a test document with known text
    word = win32com.client.GetActiveObject("Word.Application")
    if not word:
        word = win32com.client.Dispatch("Word.Application")
    word.Visible = True

    doc = word.Documents.Add()
    time.sleep(0.5)

    # Write test text with paragraphs
    original_text = (
        "EVALUATION OF LIABILITY\r"
        "Plaintiff's complaint alleges a single cause of action for "
        "Premises Liability against Fastrip. To prevail on a claim for "
        "premises liability, a plaintiff must prove: (1) the defendant "
        "owned, leased, occupied, or controlled the property; (2) the "
        "defendant was negligent in the use or maintenance of the property; "
        "(3) the plaintiff was harmed; and (4) the defendant's negligence "
        "was a substantial factor in causing the plaintiff's harm.\r"
        "A. Ownership or Control\r"
        "It is undisputed that Fastrip owned and operated the property "
        "where the incident took place. As such, Plaintiff will be able "
        "to establish this element.\r"
        "B. Negligent Use or Maintenance\r"
        "Plaintiff will allege that the property was negligently maintained "
        "because a wet floor was left unattended without warning signs.\r"
    )

    # Insert the text
    doc.Content.Text = original_text
    time.sleep(0.3)

    # Now create revised text - MOSTLY IDENTICAL with a few targeted changes
    revised_text = (
        "EVALUATION OF LIABILITY\r"
        "Plaintiff's complaint alleges a single cause of action for "
        "Premises Liability against Fastrip. To prevail on a claim for "
        "premises liability, a plaintiff must prove: (1) the defendant "
        "owned, leased, occupied, or controlled the property; (2) the "
        "defendant was negligent in the use or maintenance of the property; "
        "(3) the plaintiff was harmed; and (4) the defendant's negligence "
        "was a substantial factor in causing the plaintiff's harm.\r"
        "A. Ownership or Control\r"
        # CHANGED: Added a sentence
        "It is undisputed that Fastrip owned and operated the property "
        "where the incident took place. Recent discovery confirms this fact. "
        "As such, Plaintiff will be able to establish this element.\r"
        "B. Negligent Use or Maintenance\r"
        # CHANGED: "wet floor" -> "spill on the floor"
        "Plaintiff will allege that the property was negligently maintained "
        "because a spill on the floor was left unattended without warning signs.\r"
    )

    # Select all the text
    doc.Content.Select()
    selection = word.Selection
    range_start = selection.Range.Start
    range_end = selection.Range.End

    # Capture original text from selection (like the real code does)
    captured_text = selection.Text
    print(f"\n=== CAPTURED TEXT FROM WORD ===")
    print(f"Length: {len(captured_text)}")
    print(f"Repr (first 200): {repr(captured_text[:200])}")

    # Check if captured text matches what we wrote
    print(f"\n=== COMPARISON ===")
    print(f"Original text length: {len(original_text)}")
    print(f"Captured text length: {len(captured_text)}")
    print(f"Match: {original_text == captured_text}")

    # Character-by-character comparison for mismatches
    mismatches = []
    for i, (a, b) in enumerate(zip(original_text, captured_text)):
        if a != b:
            mismatches.append((i, a, b))
    if len(original_text) != len(captured_text):
        mismatches.append(("LENGTH", len(original_text), len(captured_text)))
    if mismatches:
        print(f"Mismatches found: {len(mismatches)}")
        for m in mismatches[:10]:
            if m[0] == "LENGTH":
                print(f"  LENGTH: original={m[1]}, captured={m[2]}")
            else:
                print(f"  Position {m[0]}: original={repr(m[1])} U+{ord(m[1]):04X}, "
                      f"captured={repr(m[2])} U+{ord(m[2]):04X}")
    else:
        print("PERFECT MATCH!")

    # Now test the diff
    print(f"\n=== DIFF TEST ===")
    # Normalize revised text to use \r (like the fix does)
    revised_clean = revised_text.replace('\r\n', '\n').replace('\n', '\r')

    dmp = diff_match_patch()
    diffs = dmp.diff_main(captured_text, revised_clean)
    dmp.diff_cleanupSemantic(diffs)

    equal_chars = sum(len(t) for op, t in diffs if op == 0)
    delete_chars = sum(len(t) for op, t in diffs if op == -1)
    insert_chars = sum(len(t) for op, t in diffs if op == 1)
    print(f"Equal: {equal_chars}, Deleted: {delete_chars}, Inserted: {insert_chars}")
    print(f"Segments: {len(diffs)}")
    print(f"Pct preserved: {equal_chars / len(captured_text) * 100:.1f}%")

    for op, text in diffs:
        label = {0: 'EQUAL', -1: 'DELETE', 1: 'INSERT'}[op]
        preview = repr(text[:60]) + ('...' if len(text) > 60 else '')
        print(f"  {label}: {preview}")

    # Now apply the redlines
    print(f"\n=== APPLYING REDLINES ===")

    # Enable Track Changes
    doc.TrackRevisions = True
    time.sleep(0.2)

    # Build the range
    range_obj = doc.Range(range_start, range_end)

    # Build operations
    operations = []
    current_pos = 0
    for op, text in diffs:
        if op == 0:
            current_pos += len(text)
        elif op == -1:
            operations.append(('delete', current_pos, current_pos + len(text)))
            current_pos += len(text)
        elif op == 1:
            operations.append(('insert', current_pos, text))

    print(f"Operations to apply: {len(operations)}")
    for op in operations:
        if op[0] == 'delete':
            print(f"  DELETE positions {op[1]}-{op[2]}")
        else:
            print(f"  INSERT at {op[1]}: {repr(op[2][:50])}")

    # Apply in reverse
    for operation in reversed(operations):
        if operation[0] == 'delete':
            _, start_offset, end_offset = operation
            delete_range = doc.Range(
                range_obj.Start + start_offset,
                range_obj.Start + end_offset
            )
            actual_text = delete_range.Text
            print(f"  Deleting [{start_offset}:{end_offset}]: {repr(actual_text[:50])}")
            delete_range.Delete()
        elif operation[0] == 'insert':
            _, offset, text = operation
            insert_range = doc.Range(
                range_obj.Start + offset,
                range_obj.Start + offset
            )
            print(f"  Inserting at [{offset}]: {repr(text[:50])}")
            insert_range.InsertAfter(text)

    print(f"\n=== DONE ===")
    print("Check the Word document for tracked changes.")
    print("Expected: Only 'wet floor' -> 'spill on the floor' and added sentence should be tracked.")
    print("If the ENTIRE section is struck out and rewritten, the position tracking is broken.")

    # Count tracked changes
    time.sleep(0.5)
    try:
        revisions = doc.Revisions
        print(f"\nTotal revisions: {revisions.Count}")
        for i in range(1, min(revisions.Count + 1, 20)):
            rev = revisions(i)
            rev_type = "INSERT" if rev.Type == 1 else "DELETE" if rev.Type == 2 else f"TYPE_{rev.Type}"
            rev_text = rev.Range.Text[:60] if rev.Range.Text else ""
            print(f"  Rev {i}: {rev_type} - {repr(rev_text)}")
    except Exception as e:
        print(f"Could not enumerate revisions: {e}")

    return True


if __name__ == "__main__":
    try:
        test_redline()
    except Exception as e:
        import traceback
        print(f"ERROR: {e}")
        traceback.print_exc()
