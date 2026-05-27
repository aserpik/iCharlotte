"""
Test the paragraph-level redline engine on the actual Wehunt document.
Simulates the AI Assistant for Windows workflow:
1. Opens the document in Word
2. Selects the EVALUATION OF LIABILITY section
3. Calls Gemini with "update and improve this section"
4. Applies paragraph-level redlines
5. Reports results

Run: python tests/test_redline_real_doc.py
"""
import sys
import time
import shutil
import os

# Add project root to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

DOC_PATH = r"Z:\Shared\Current Clients\3800- NATIONWIDE\173 - Wehunt\STATUS\[Draft] Carrier005 (Pre-Trial Report).docx"
BACKUP_PATH = DOC_PATH.replace(".docx", "_BACKUP.docx")
OUTPUT_PATH = DOC_PATH.replace(".docx", " - REDLINE TEST.docx")


def find_section(word_app, doc, section_heading, next_sections):
    """Find a section range in the document by heading text."""
    find_range = doc.Content.Duplicate

    find_range.Find.ClearFormatting()
    find_range.Find.Execute(FindText=section_heading, Forward=True)

    if not find_range.Find.Found:
        print(f"ERROR: Could not find '{section_heading}' in document")
        return None, None

    section_start = find_range.Start
    print(f"Found '{section_heading}' at position {section_start}")

    section_end = doc.Content.End
    for next_heading in next_sections:
        search_range = doc.Range(section_start + len(section_heading) + 5, doc.Content.End)
        search_range.Find.ClearFormatting()
        search_range.Find.Execute(FindText=next_heading, Forward=True)
        if search_range.Find.Found:
            section_end = search_range.Start
            print(f"Found next section '{next_heading}' at position {section_end}")
            break

    return section_start, section_end


def call_gemini(prompt: str, text: str) -> str:
    """Call Gemini with the prompt and text."""
    from icharlotte_core.llm import LLMHandler
    from icharlotte_core.config import API_KEYS

    provider = "Gemini"
    model_id = "gemini-3.1-pro-preview"

    api_key = API_KEYS.get(provider)
    if not api_key:
        raise ValueError(f"No API key for {provider}. Available: {list(API_KEYS.keys())}")

    settings = {
        'temperature': 0.7,
        'top_p': 0.95,
        'max_tokens': -1,
        'stream': False,
        'thinking_level': 'None'
    }

    # Use the redline-aware system prompt
    system_prompt = (
        "You are a helpful writing assistant operating in REDLINE MODE. "
        "Your output will be diffed against the original text to produce "
        "Track Changes in Microsoft Word. You MUST preserve unchanged text "
        "EXACTLY as written — same words, same order, same punctuation. "
        "Only modify the specific parts that need changing per the user's "
        "instructions. Output only the requested text without any preamble, "
        "explanation, or markdown formatting unless specifically asked."
    )

    # Build the prompt like the AI Assistant does (no full document context)
    redline_prefix = (
        "IMPORTANT — REDLINE MODE IS ACTIVE: Your output will be compared "
        "word-by-word against the original text to generate Track Changes. "
        "You MUST follow these rules:\n"
        "1. Preserve any text that does not need to change EXACTLY as-is — "
        "same wording, same sentence structure, same word order.\n"
        "2. PRESERVE THE EXACT PARAGRAPH STRUCTURE — keep the same number of "
        "paragraphs and blank lines between them. Each paragraph must start "
        "and end at the same boundaries as the original. Do NOT merge paragraphs, "
        "split paragraphs, or remove blank lines.\n"
        "3. Do NOT rephrase, reorganize, or rewrite portions that are already correct.\n"
        "4. Only modify the specific words or sentences that need to change.\n"
        "5. If a sentence is fine as-is, copy it verbatim.\n\n"
    )

    full_prompt = redline_prefix + prompt + "\n\nText to process:\n" + text

    print(f"Calling Gemini ({model_id})...")
    print(f"Prompt length: {len(full_prompt)} chars")
    print(f"Selected text length: {len(text)} chars")

    result = LLMHandler.generate(
        provider=provider,
        model=model_id,
        system_prompt=system_prompt,
        user_prompt=full_prompt,
        file_contents="",
        settings=settings
    )

    if not result:
        raise ValueError("LLM returned empty response")

    return result.strip()


def apply_paragraph_redlines(doc, range_start, range_end, original_text, revised_text):
    """Apply redlines using flat diff mapped back to Word paragraphs.

    Strategy: flatten both texts (strip paragraph breaks), diff them
    character-by-character, then map each change back to the correct
    Word paragraph. This never touches paragraph marks (\r), so
    formatting is always preserved regardless of LLM paragraph behavior.
    """
    from diff_match_patch import diff_match_patch
    import bisect

    # --- 1. Build Word paragraph map ---
    range_obj = doc.Range(range_start, range_end)
    word_paras = []
    para_collection = range_obj.Paragraphs
    para_count = para_collection.Count
    for i in range(1, para_count + 1):
        wp = para_collection(i)
        text = wp.Range.Text
        # Strip trailing \r — that's the paragraph mark we must never touch
        content = text[:-1] if text.endswith('\r') else text
        word_paras.append({
            'start': wp.Range.Start,
            'end': wp.Range.End,
            'text': text,
            'content': content,
        })

    print(f"\nWord paragraphs: {len(word_paras)}")
    for i, wp in enumerate(word_paras):
        print(f"  [{i}] start={wp['start']} end={wp['end']} "
              f"({len(wp['content'])} chars) {repr(wp['content'][:70])}")

    # --- 2. Build flat original text and position map ---
    # Join paragraphs with a single space separator. This matches how LLMs
    # typically return text (spaces where paragraph breaks were).
    # We track which flat positions are "boundary spaces" so we can map
    # back to Word positions correctly (boundary spaces don't exist in Word).
    import re

    orig_flat = ''
    flat_para_starts = []  # flat offset where each paragraph's CONTENT begins
    boundary_positions = set()  # flat positions that are separator spaces

    for i, wp in enumerate(word_paras):
        if i > 0 and wp['content']:
            # Add separator space between paragraphs (only before non-empty ones)
            # Find last non-empty paragraph before this one
            prev_had_content = any(word_paras[j]['content'] for j in range(i))
            if prev_had_content:
                boundary_positions.add(len(orig_flat))
                orig_flat += ' '
        flat_para_starts.append(len(orig_flat))
        orig_flat += wp['content']
    # Sentinel for bisect
    flat_para_starts.append(len(orig_flat))

    print(f"\nOriginal flat text: {len(orig_flat)} chars")
    print(f"Boundary spaces: {len(boundary_positions)}")

    # --- 3. Flatten revised text ---
    # Replace all line breaks with spaces, collapse multiple spaces
    rev_flat = revised_text.replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')
    rev_flat = re.sub(r'  +', ' ', rev_flat).strip()

    print(f"Revised flat text: {len(rev_flat)} chars")

    # --- 4. Diff the flat texts ---
    dmp = diff_match_patch()
    diffs = dmp.diff_main(orig_flat, rev_flat)
    dmp.diff_cleanupSemantic(diffs)

    eq = sum(len(t) for op, t in diffs if op == 0)
    de = sum(len(t) for op, t in diffs if op == -1)
    ins = sum(len(t) for op, t in diffs if op == 1)
    print(f"\nFlat diff: {eq} equal, {de} deleted, {ins} inserted")
    pct = eq / len(orig_flat) * 100 if orig_flat else 0
    print(f"Preserved: {pct:.1f}%")

    # Show diff segments
    for op, text in diffs:
        label = {0: 'EQUAL', -1: 'DELETE', 1: 'INSERT'}[op]
        preview = repr(text[:80]) + ('...' if len(text) > 80 else '')
        print(f"  {label}: {preview}")

    # --- 5. Build operations list with flat positions ---
    # Filter out operations that only affect boundary spaces (artifacts of flattening)
    operations = []
    pos = 0  # position in orig_flat
    for op, text in diffs:
        if op == 0:
            pos += len(text)
        elif op == -1:
            # Check if the entire delete is just boundary space(s)
            all_boundary = all(pos + i in boundary_positions for i in range(len(text)))
            if not all_boundary:
                # Filter out boundary characters from the delete
                real_start = pos
                real_end = pos + len(text)
                # Find non-boundary ranges within this delete
                i = real_start
                while i < real_end:
                    if i in boundary_positions:
                        i += 1
                        continue
                    # Start of a real delete segment
                    seg_start = i
                    while i < real_end and i not in boundary_positions:
                        i += 1
                    operations.append(('delete', seg_start, i, orig_flat[seg_start:i]))
            else:
                print(f"  SKIP boundary delete: {repr(text)}")
            pos += len(text)
        elif op == 1:
            # Check if this is a space inserted right at a boundary position
            if text == ' ' and pos in boundary_positions:
                print(f"  SKIP boundary insert at {pos}")
            else:
                operations.append(('insert', pos, text))

    print(f"\nOperations: {len(operations)}")

    # --- 6. Map flat positions to Word positions and apply in reverse ---
    doc.TrackRevisions = True
    time.sleep(0.2)

    total_changes = 0

    def flat_to_word_pos(flat_pos):
        """Convert a flat string position to a Word document position."""
        para_idx = bisect.bisect_right(flat_para_starts, flat_pos) - 1
        if para_idx < 0:
            para_idx = 0
        if para_idx >= len(word_paras):
            para_idx = len(word_paras) - 1
        offset_in_para = flat_pos - flat_para_starts[para_idx]
        return word_paras[para_idx]['start'] + offset_in_para

    print("\nApplying redlines...")
    for operation in reversed(operations):
        if operation[0] == 'delete':
            _, flat_start, flat_end, del_text = operation
            start_para = bisect.bisect_right(flat_para_starts, flat_start) - 1
            end_para = bisect.bisect_right(flat_para_starts, flat_end - 1) - 1

            if start_para == end_para:
                word_start = flat_to_word_pos(flat_start)
                word_end = flat_to_word_pos(flat_end)
                # Never delete past the paragraph mark (\r)
                para_content_end = word_paras[start_para]['end'] - 1
                if word_end > para_content_end:
                    word_end = para_content_end
                if word_start >= word_end:
                    print(f"  SKIP empty delete in para [{start_para}]")
                    continue
                # Skip if this only deletes trailing whitespace at paragraph end
                # (deleting chars adjacent to \r risks capturing the \r mark)
                if word_end == para_content_end and del_text.strip() == '':
                    print(f"  SKIP trailing whitespace in para [{start_para}]")
                    continue
                dr = doc.Range(word_start, word_end)
                print(f"  DEL in para [{start_para}]: {repr(dr.Text[:60])}")
                dr.Delete()
                total_changes += 1
            else:
                for pidx in reversed(range(start_para, end_para + 1)):
                    p_flat_start = flat_para_starts[pidx]
                    p_flat_end = flat_para_starts[pidx + 1]
                    clip_start = max(flat_start, p_flat_start)
                    clip_end = min(flat_end, p_flat_end)
                    if clip_start >= clip_end:
                        continue
                    word_start = word_paras[pidx]['start'] + (clip_start - p_flat_start)
                    word_end = word_paras[pidx]['start'] + (clip_end - p_flat_start)
                    # Never delete past the paragraph mark
                    para_content_end = word_paras[pidx]['end'] - 1
                    if word_end > para_content_end:
                        word_end = para_content_end
                    if word_start >= word_end:
                        continue
                    dr = doc.Range(word_start, word_end)
                    print(f"  DEL in para [{pidx}]: {repr(dr.Text[:60])}")
                    dr.Delete()
                    total_changes += 1

        elif operation[0] == 'insert':
            _, flat_pos, ins_text = operation
            word_pos = flat_to_word_pos(flat_pos)
            ir = doc.Range(word_pos, word_pos)
            ir.InsertAfter(ins_text)
            para_idx = bisect.bisect_right(flat_para_starts, flat_pos) - 1
            print(f"  INS in para [{para_idx}]: {repr(ins_text[:60])}")
            total_changes += 1

    print(f"\nTotal changes applied: {total_changes}")

    # ================================================================
    # POST-APPLICATION VALIDATION (shared module)
    # ================================================================
    time.sleep(0.5)
    pre_edit_content_paras = sum(1 for wp in word_paras if wp['content'].strip())

    try:
        from icharlotte_core.word_validator import validate_redline
        val_result = validate_redline(
            doc_com=doc,
            range_start=range_start,
            range_end=range_end,
            pre_content_para_count=pre_edit_content_paras,
            orig_flat_length=len(orig_flat),
            deleted_chars=de,
            inserted_chars=ins,
        )
        val_result.print_summary()
    except Exception as e:
        print(f"Could not run validation: {e}")
        import traceback
        traceback.print_exc()

    return not val_result.has_errors if 'val_result' in dir() else False


def main():
    import win32com.client

    # Make backup
    if os.path.exists(DOC_PATH) and not os.path.exists(BACKUP_PATH):
        shutil.copy2(DOC_PATH, BACKUP_PATH)
        print(f"Backup saved to: {BACKUP_PATH}")

    # Connect to Word
    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True

    # Make a working copy so we never touch the original
    shutil.copy2(DOC_PATH, OUTPUT_PATH)
    print(f"Working copy: {OUTPUT_PATH}")

    # Close any stale test copies first
    for i in range(word.Documents.Count, 0, -1):
        d = word.Documents(i)
        if "REDLINE TEST" in d.Name:
            d.Close(SaveChanges=False)
            time.sleep(0.5)

    # Open the copy
    doc = word.Documents.Open(OUTPUT_PATH)
    time.sleep(2)
    doc.Activate()
    time.sleep(0.5)
    print(f"Opened: {doc.Name}")

    # Find the EVALUATION OF EXPOSURE section
    section_heading = "EVALUATION OF EXPOSURE"
    next_sections = ["SETTLEMENT", "FURTHER CASE HANDLING"]
    section_start, section_end = find_section(word, doc, section_heading, next_sections)
    if section_start is None:
        return

    # Select the section
    sel_range = doc.Range(section_start, section_end)
    original_text = sel_range.Text
    print(f"\nSelected {section_heading}: {len(original_text)} chars")
    print(f"Preview: {repr(original_text[:200])}")
    print(f"Line endings: \\r={original_text.count(chr(13))}, \\n={original_text.count(chr(10))}")

    # Call Gemini
    try:
        revised_text = call_gemini("update and improve this section.", original_text)
    except Exception as e:
        print(f"LLM call failed: {e}")
        import traceback
        traceback.print_exc()
        return

    print(f"\nGemini returned: {len(revised_text)} chars")
    print(f"Preview: {repr(revised_text[:200])}")

    # Re-establish COM connection (can go stale during long LLM calls)
    try:
        word = win32com.client.GetActiveObject("Word.Application")
        for i in range(1, word.Documents.Count + 1):
            d = word.Documents(i)
            if "REDLINE TEST" in d.Name:
                doc = d
                break
        print(f"Re-connected to: {doc.Name}")
    except Exception as e:
        print(f"COM reconnect failed: {e}")
        return

    # Apply redlines
    passed = apply_paragraph_redlines(doc, section_start, section_end, original_text, revised_text)

    # Save the redlined document
    doc.Save()
    print(f"\nSaved redlined document to: {OUTPUT_PATH}")

    if passed:
        print("\n=== TEST PASSED ===")
    else:
        print("\n=== TEST FAILED — SEE ERRORS ABOVE ===")
    print(f"Redlined version: {OUTPUT_PATH}")
    print(f"Original untouched: {DOC_PATH}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print(f"ERROR: {e}")
        traceback.print_exc()
