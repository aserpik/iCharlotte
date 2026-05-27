"""
AI Assistant for Word — Evaluation Test Script
================================================
Tests 4 combinations of settings with the same prompt and model on a specific
section of a legal document.

Runs:
  1. Use All Doc Text = ON,  Redline = ON
  2. Use All Doc Text = ON,  Redline = OFF
  3. Use All Doc Text = OFF, Redline = ON
  4. Use All Doc Text = OFF, Redline = OFF

Each run starts from the original (unmodified) document.
"""

import sys
import os
import time
import shutil
import traceback

# Add project root
sys.path.insert(0, r'C:\geminiterminal2')

import win32com.client
import pythoncom

from icharlotte_core.llm import LLMHandler
from icharlotte_core.word_hotkey import apply_flat_diff_redline

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
DOC_PATH = r'C:\geminiterminal2\tests\AI Assistant\[OLD} SS iso MTC Further Rsps RPD.docx'
SECTION_HEADING = "REASON WHY FURTHER RESPONSE IS NOT WARRANTED AS TO REQUESTS FOR PRODUCTION NO. 6"
PROMPT = "convert this to a persuasive argument in narrative format"
MODEL = "gemini-3.5-flash"
PROVIDER = "Gemini"

LLM_SETTINGS = {
    'temperature': 0.7,
    'top_p': 0.95,
    'max_tokens': -1,
    'stream': False,
    'thinking_level': 'None',
}

REDLINE_SYSTEM_PROMPT = (
    "You are a helpful writing assistant operating in REDLINE MODE. "
    "Your output will be diffed against the original text to produce "
    "Track Changes in Microsoft Word. You MUST preserve unchanged text "
    "EXACTLY as written — same words, same order, same punctuation. "
    "Only modify the specific parts that need changing per the user's "
    "instructions. Output only the requested text without any preamble, "
    "explanation, or markdown formatting unless specifically asked."
)

NORMAL_SYSTEM_PROMPT = (
    "You are a helpful writing assistant. Follow the user's instructions "
    "precisely. Output only the requested text without any preamble, "
    "explanation, or markdown formatting unless specifically asked."
)

REDLINE_PREFIX = (
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

# Runs definition: (label, use_all_text, redline)
RUNS = [
    ("Run 1: All Doc Text ON  + Redline ON",  True,  True),
    ("Run 2: All Doc Text ON  + Redline OFF", True,  False),
    ("Run 3: All Doc Text OFF + Redline ON",  False, True),
    ("Run 4: All Doc Text OFF + Redline OFF", False, False),
]

OUTPUT_DIR = r'C:\geminiterminal2\tests\AI Assistant\eval_outputs'


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_word_app():
    """Get or create Word COM application."""
    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except Exception:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
    return word


def open_document(word, path):
    """Open a document, closing the previous instance if open."""
    # Check if already open
    for i in range(1, word.Documents.Count + 1):
        doc = word.Documents(i)
        if os.path.normcase(doc.FullName) == os.path.normcase(path):
            doc.Close(SaveChanges=0)  # Close without saving
            break
    time.sleep(0.5)
    doc = word.Documents.Open(path)
    time.sleep(0.5)
    return doc


def find_section_range(doc, heading_text):
    """
    Find the range of a section by its heading text.
    Returns (range_start, range_end, section_text).

    Searches for the heading, then captures everything from the first
    paragraph AFTER the heading until the next same-level heading or end of doc.
    """
    # Search for the heading paragraph
    total_paras = doc.Paragraphs.Count
    heading_para_idx = None

    for i in range(1, total_paras + 1):
        para = doc.Paragraphs(i)
        para_text = para.Range.Text.strip()
        if heading_text.lower() in para_text.lower():
            heading_para_idx = i
            print(f"  Found heading at paragraph {i}: '{para_text[:80]}...'")
            break

    if heading_para_idx is None:
        raise ValueError(f"Could not find heading: {heading_text}")

    # The content starts at the paragraph AFTER the heading
    content_start_idx = heading_para_idx + 1
    if content_start_idx > total_paras:
        raise ValueError("Heading is the last paragraph — no content after it")

    content_start = doc.Paragraphs(content_start_idx).Range.Start

    # Find the end: look for the next heading at the same or higher level,
    # or another "REASON WHY" heading, or end of document
    content_end = None
    for i in range(content_start_idx, total_paras + 1):
        para = doc.Paragraphs(i)
        para_text = para.Range.Text.strip()

        # Check if this is another major heading (all caps, bold, or starts with "REASON WHY")
        if i > content_start_idx:
            is_heading = False
            # Check for next "REASON WHY" section
            if para_text.upper().startswith("REASON WHY"):
                is_heading = True
            # Check for all-caps heading (>10 chars, all uppercase)
            elif len(para_text) > 10 and para_text == para_text.upper() and para.Range.Font.Bold:
                is_heading = True

            if is_heading:
                content_end = doc.Paragraphs(i).Range.Start
                print(f"  Section ends before paragraph {i}: '{para_text[:60]}...'")
                break

    if content_end is None:
        # Go to end of last paragraph
        content_end = doc.Paragraphs(total_paras).Range.End
        print(f"  Section extends to end of document")

    # Get the text
    section_range = doc.Range(content_start, content_end)
    section_text = section_range.Text

    return content_start, content_end, section_text


def get_all_document_text(doc):
    """Get the full document text."""
    return doc.Content.Text


def build_prompt(base_prompt, selected_text, all_doc_text, use_all_text, redline):
    """Build the full prompt matching word_hotkey.py logic."""
    prompt = base_prompt

    # Prepend redline prefix if in redline mode
    if redline:
        prompt = REDLINE_PREFIX + prompt

    # Build full prompt based on use_all_text setting
    if use_all_text:
        full_prompt = (
            f"{prompt}\n\n"
            f"=== FULL DOCUMENT (for context) ===\n{all_doc_text}\n\n"
            f"=== SELECTED TEXT TO PROCESS ===\n{selected_text}\n\n"
            f"Process only the SELECTED TEXT above according to the instructions. "
            f"Use the full document for context but output only the processed version of the selected text."
        )
    else:
        full_prompt = f"{prompt}\n\nText to process:\n{selected_text}"

    return full_prompt


def call_llm(full_prompt, redline):
    """Call the LLM and return the result."""
    system_prompt = REDLINE_SYSTEM_PROMPT if redline else NORMAL_SYSTEM_PROMPT

    result = LLMHandler.generate(
        provider=PROVIDER,
        model=MODEL,
        system_prompt=system_prompt,
        user_prompt=full_prompt,
        file_contents="",
        settings=LLM_SETTINGS,
    )
    return result.strip() if result else ""


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Make a backup of the original document
    backup_path = os.path.join(OUTPUT_DIR, "_original_backup.docx")
    shutil.copy2(DOC_PATH, backup_path)
    print(f"Backed up original to: {backup_path}")

    word = get_word_app()
    results = {}

    for run_label, use_all_text, redline in RUNS:
        print(f"\n{'='*70}")
        print(f"  {run_label}")
        print(f"  use_all_text={use_all_text}, redline={redline}")
        print(f"{'='*70}")

        # Restore original document for each run
        shutil.copy2(backup_path, DOC_PATH)
        time.sleep(0.3)

        # Open the document fresh
        doc = open_document(word, DOC_PATH)
        print(f"  Opened document: {doc.Name}")

        try:
            # Find the section
            range_start, range_end, section_text = find_section_range(doc, SECTION_HEADING)
            print(f"  Section range: {range_start} - {range_end}")
            print(f"  Section text length: {len(section_text)} chars")
            print(f"  Section preview: {section_text[:200]}...")

            # Get full doc text if needed
            all_doc_text = get_all_document_text(doc) if use_all_text else ""

            # Build the prompt
            full_prompt = build_prompt(PROMPT, section_text, all_doc_text, use_all_text, redline)

            # Call the LLM
            print(f"  Calling LLM ({PROVIDER}/{MODEL})...")
            start_time = time.time()
            llm_output = call_llm(full_prompt, redline)
            elapsed = time.time() - start_time
            print(f"  LLM responded in {elapsed:.1f}s, output length: {len(llm_output)} chars")

            # Store the raw LLM output
            results[run_label] = {
                'llm_output': llm_output,
                'elapsed': elapsed,
                'original_text': section_text,
                'redline': redline,
                'use_all_text': use_all_text,
            }

            # Apply redlines if in redline mode
            if redline:
                print(f"  Applying redlines to document...")
                try:
                    redline_result = apply_flat_diff_redline(
                        doc_com=doc,
                        range_start=range_start,
                        range_end=range_end,
                        original_text=section_text,
                        revised_text=llm_output,
                        auto_enable_track_changes=True,
                    )
                    results[run_label]['redline_result'] = redline_result
                    if redline_result.get('success'):
                        print(f"  Redlines applied: {redline_result.get('total_changes', 0)} changes")
                        print(f"    Equal: {redline_result.get('equal_chars', 0)}, "
                              f"Deleted: {redline_result.get('deleted_chars', 0)}, "
                              f"Inserted: {redline_result.get('inserted_chars', 0)}")
                    else:
                        print(f"  Redline application FAILED: {redline_result.get('error', 'unknown')}")
                except Exception as e:
                    print(f"  Redline application ERROR: {e}")
                    traceback.print_exc()
                    results[run_label]['redline_error'] = str(e)

                # Save redlined document
                save_name = run_label.replace(":", "").replace(" ", "_").replace("+", "").strip("_")
                save_path = os.path.join(OUTPUT_DIR, f"{save_name}_redlined.docx")
                doc.SaveAs(save_path)
                print(f"  Saved redlined doc: {save_path}")
                results[run_label]['saved_doc'] = save_path
            else:
                # For non-redline mode, we can also apply the text to show the result
                # But the user mainly wants the output text
                # Save the output as a text file for review
                save_name = run_label.replace(":", "").replace(" ", "_").replace("+", "").strip("_")
                txt_path = os.path.join(OUTPUT_DIR, f"{save_name}_output.txt")
                with open(txt_path, 'w', encoding='utf-8') as f:
                    f.write(llm_output)
                print(f"  Saved LLM output: {txt_path}")
                results[run_label]['saved_txt'] = txt_path

            # Close doc without saving (we already saved copies if needed)
            doc.Close(SaveChanges=0)
            time.sleep(0.3)

        except Exception as e:
            print(f"  ERROR in {run_label}: {e}")
            traceback.print_exc()
            results[run_label] = {'error': str(e)}
            try:
                doc.Close(SaveChanges=0)
            except:
                pass

    # Restore original document
    shutil.copy2(backup_path, DOC_PATH)
    print(f"\nRestored original document from backup.")

    # ---------------------------------------------------------------------------
    # Print Summary
    # ---------------------------------------------------------------------------
    print(f"\n\n{'#'*70}")
    print(f"  EVALUATION RESULTS SUMMARY")
    print(f"{'#'*70}")

    for run_label, use_all_text, redline in RUNS:
        print(f"\n{'='*70}")
        print(f"  {run_label}")
        print(f"{'='*70}")

        data = results.get(run_label, {})

        if 'error' in data and 'llm_output' not in data:
            print(f"  FAILED: {data['error']}")
            continue

        print(f"  Time: {data.get('elapsed', 0):.1f}s")
        print(f"  Output length: {len(data.get('llm_output', ''))} chars")

        if redline and 'redline_result' in data:
            rr = data['redline_result']
            print(f"  Redline stats: {rr.get('total_changes', 0)} changes, "
                  f"equal={rr.get('equal_chars', 0)}, "
                  f"deleted={rr.get('deleted_chars', 0)}, "
                  f"inserted={rr.get('inserted_chars', 0)}")

        if redline and 'redline_error' in data:
            print(f"  Redline ERROR: {data['redline_error']}")

        print(f"\n  --- LLM OUTPUT ---")
        print(data.get('llm_output', '(no output)'))
        print(f"  --- END OUTPUT ---")

    # Save full report
    report_path = os.path.join(OUTPUT_DIR, "evaluation_report.txt")
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write("AI Assistant for Word — Evaluation Report\n")
        f.write(f"Model: {PROVIDER}/{MODEL}\n")
        f.write(f"Prompt: {PROMPT}\n")
        f.write(f"Section: {SECTION_HEADING}\n")
        f.write(f"Date: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("="*70 + "\n\n")

        for run_label, use_all_text, redline in RUNS:
            f.write(f"\n{'='*70}\n")
            f.write(f"{run_label}\n")
            f.write(f"use_all_text={use_all_text}, redline={redline}\n")
            f.write(f"{'='*70}\n\n")

            data = results.get(run_label, {})

            if 'error' in data and 'llm_output' not in data:
                f.write(f"FAILED: {data['error']}\n")
                continue

            f.write(f"Time: {data.get('elapsed', 0):.1f}s\n")
            f.write(f"Output length: {len(data.get('llm_output', ''))} chars\n")

            if redline and 'redline_result' in data:
                rr = data['redline_result']
                f.write(f"Redline stats: {rr.get('total_changes', 0)} changes, "
                        f"equal={rr.get('equal_chars', 0)}, "
                        f"deleted={rr.get('deleted_chars', 0)}, "
                        f"inserted={rr.get('inserted_chars', 0)}\n")

            f.write(f"\n--- ORIGINAL TEXT ---\n")
            f.write(data.get('original_text', '(not captured)'))
            f.write(f"\n--- LLM OUTPUT ---\n")
            f.write(data.get('llm_output', '(no output)'))
            f.write(f"\n--- END ---\n")

    print(f"\nFull report saved to: {report_path}")
    print("Done!")


if __name__ == "__main__":
    main()
