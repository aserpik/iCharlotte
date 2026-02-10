"""
Test paragraph-level redline engine against a real Word document.
Verifies that:
1. Unchanged paragraphs keep their formatting
2. Changed words within a paragraph get surgical tracked changes
3. Paragraph marks are never deleted (preserving styles)

Run: python tests/test_redline_paragraph.py
"""
import time


def test_paragraph_redline():
    import win32com.client
    from diff_match_patch import diff_match_patch
    from difflib import SequenceMatcher

    word = win32com.client.GetActiveObject("Word.Application")
    word.Visible = True

    doc = word.Documents.Add()
    time.sleep(0.5)

    # Build a document with formatted paragraphs
    sel = word.Selection

    # Heading - Bold, Underline, ALL CAPS
    sel.Font.Bold = True
    sel.Font.Underline = True
    sel.TypeText("EVALUATION OF LIABILITY")
    sel.TypeParagraph()

    # Body paragraph - Normal
    sel.Font.Bold = False
    sel.Font.Underline = False
    sel.TypeText(
        "Plaintiff's complaint alleges a single cause of action for "
        "Premises Liability against Fastrip. To prevail, a plaintiff must prove "
        "the required elements."
    )
    sel.TypeParagraph()

    # Subheading - Bold, Underline
    sel.Font.Bold = True
    sel.Font.Underline = True
    sel.TypeText("A. Ownership or Control")
    sel.TypeParagraph()

    # Body paragraph - Normal
    sel.Font.Bold = False
    sel.Font.Underline = False
    sel.TypeText(
        "It is undisputed that Fastrip owned and operated the property "
        "where the incident took place. Plaintiff will establish this element."
    )
    sel.TypeParagraph()

    # Another subheading
    sel.Font.Bold = True
    sel.Font.Underline = True
    sel.TypeText("B. Negligent Use or Maintenance")
    sel.TypeParagraph()

    # Body paragraph - Normal
    sel.Font.Bold = False
    sel.Font.Underline = False
    sel.TypeText(
        "Plaintiff will allege that the property was negligently maintained "
        "because a wet floor was left unattended without warning signs."
    )
    sel.TypeParagraph()

    time.sleep(0.3)

    # Now capture the text
    doc.Content.Select()
    sel = word.Selection
    original_text = sel.Text
    range_start = sel.Range.Start
    range_end = sel.Range.End

    print(f"Original: {len(original_text)} chars")
    print(f"Paragraphs (by \\r): {original_text.count(chr(13))}")

    # Simulate LLM output with \n and targeted changes:
    # - Heading: unchanged
    # - First body: unchanged
    # - Subheading A: unchanged
    # - Second body: CHANGED (added a sentence)
    # - Subheading B: unchanged
    # - Third body: CHANGED (wet floor -> spill)
    revised_text = (
        "EVALUATION OF LIABILITY\n"
        "Plaintiff\u2019s complaint alleges a single cause of action for "
        "Premises Liability against Fastrip. To prevail, a plaintiff must prove "
        "the required elements.\n"
        "A. Ownership or Control\n"
        "It is undisputed that Fastrip owned and operated the property "
        "where the incident took place. Recent discovery confirms this fact. "
        "Plaintiff will establish this element.\n"
        "B. Negligent Use or Maintenance\n"
        "Plaintiff will allege that the property was negligently maintained "
        "because a spill on the floor was left unattended without warning signs.\n"
    )

    # Split into paragraphs
    orig_paras = original_text.split('\r')
    if orig_paras and orig_paras[-1] == '':
        orig_paras.pop()

    revised_clean = revised_text.replace('\r\n', '\n').replace('\r', '\n')
    rev_paras = revised_clean.split('\n')
    if rev_paras and rev_paras[-1] == '':
        rev_paras.pop()

    print(f"\nOriginal paragraphs ({len(orig_paras)}):")
    for i, p in enumerate(orig_paras):
        print(f"  [{i}] {repr(p[:60])}")

    print(f"\nRevised paragraphs ({len(rev_paras)}):")
    for i, p in enumerate(rev_paras):
        print(f"  [{i}] {repr(p[:60])}")

    # Get Word paragraphs
    range_obj = doc.Range(range_start, range_end)
    word_paras = []
    para_collection = range_obj.Paragraphs
    for i in range(1, para_collection.Count + 1):
        wp = para_collection(i)
        word_paras.append({
            'start': wp.Range.Start,
            'end': wp.Range.End,
            'text': wp.Range.Text,
        })
        print(f"  Word para [{i-1}]: start={wp.Range.Start}, end={wp.Range.End}, "
              f"bold={wp.Range.Font.Bold}, text={repr(wp.Range.Text[:40])}")

    # Diff paragraphs
    matcher = SequenceMatcher(None, orig_paras, rev_paras)
    opcodes = matcher.get_opcodes()

    print(f"\nOpcodes:")
    for tag, i1, i2, j1, j2 in opcodes:
        print(f"  {tag}: orig[{i1}:{i2}] rev[{j1}:{j2}]")

    # Enable Track Changes and apply
    doc.TrackRevisions = True
    dmp = diff_match_patch()
    total_changes = 0

    for tag, i1, i2, j1, j2 in reversed(opcodes):
        if tag == 'equal':
            continue

        elif tag == 'replace':
            orig_slice = list(range(i1, i2))
            rev_slice = rev_paras[j1:j2]
            pair_count = min(len(orig_slice), len(rev_slice))

            for k in reversed(range(pair_count)):
                orig_idx = i1 + k
                new_text = rev_slice[k]

                if orig_idx >= len(word_paras):
                    continue

                wp = word_paras[orig_idx]
                para_text = wp['text']
                if para_text.endswith('\r'):
                    para_text = para_text[:-1]
                para_text = para_text.rstrip('\r\n')

                if para_text == new_text:
                    print(f"  SKIP para [{orig_idx}] - identical")
                    continue

                diffs = dmp.diff_main(para_text, new_text)
                dmp.diff_cleanupSemantic(diffs)

                para_ops = []
                pos = 0
                for op, text in diffs:
                    if op == 0:
                        pos += len(text)
                    elif op == -1:
                        para_ops.append(('delete', pos, pos + len(text)))
                        pos += len(text)
                    elif op == 1:
                        para_ops.append(('insert', pos, text))

                para_start = wp['start']
                for operation in reversed(para_ops):
                    if operation[0] == 'delete':
                        _, s, e = operation
                        dr = doc.Range(para_start + s, para_start + e)
                        print(f"  DEL in para [{orig_idx}]: {repr(dr.Text[:50])}")
                        dr.Delete()
                        total_changes += 1
                    elif operation[0] == 'insert':
                        _, off, text = operation
                        ir = doc.Range(para_start + off, para_start + off)
                        ir.InsertAfter(text)
                        print(f"  INS in para [{orig_idx}]: {repr(text[:50])}")
                        total_changes += 1

    print(f"\nTotal changes applied: {total_changes}")

    # Verify: check revisions
    time.sleep(0.3)
    revs = doc.Revisions
    print(f"Word revisions: {revs.Count}")
    for i in range(1, min(revs.Count + 1, 15)):
        rev = revs(i)
        rtype = 'INS' if rev.Type == 1 else 'DEL' if rev.Type == 2 else f'T{rev.Type}'
        print(f"  {rtype}: {repr(rev.Range.Text[:60])}")

    # Verify: check formatting is preserved
    print(f"\nFormatting check:")
    para_collection2 = doc.Content.Paragraphs
    for i in range(1, para_collection2.Count + 1):
        wp = para_collection2(i)
        text_preview = repr(wp.Range.Text[:40])
        bold = wp.Range.Font.Bold
        underline = wp.Range.Font.Underline
        # Bold values: 0=False, -1=True, 9999999=Mixed
        bold_str = "BOLD" if bold == -1 else "mixed" if bold not in (0, -1) else "normal"
        ul_str = "UNDERLINE" if underline else "normal"
        print(f"  Para {i}: {bold_str}/{ul_str} - {text_preview}")

    print("\n=== DONE ===")
    print("CHECK THE WORD DOCUMENT:")
    print("1. Heading 'EVALUATION OF LIABILITY' should still be BOLD+UNDERLINE")
    print("2. Subheadings 'A.' and 'B.' should still be BOLD+UNDERLINE")
    print("3. Body paragraphs should be normal formatting")
    print("4. Only 'wet floor'->'spill on the floor' and added sentence should be tracked")
    print("5. Inserted text should be in NORMAL formatting, not bold/underline")


if __name__ == "__main__":
    try:
        test_paragraph_redline()
    except Exception as e:
        import traceback
        print(f"ERROR: {e}")
        traceback.print_exc()
