import sys
import os
import argparse
import traceback
import win32com.client as win32
import re

# Add the root directory to path so we can import gemini_utils
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from Scripts.gemini_utils import call_gemini_api, log_event

def get_word_app():
    """Tries to get the active Word instance, otherwise creates a new one."""
    try:
        word = win32.GetActiveObject("Word.Application")
        return word, False # False means we didn't create it
    except:
        try:
            word = win32.DispatchEx("Word.Application")
            return word, True # True means we created it
        except:
            word = win32.Dispatch("Word.Application")
            return word, True

def get_or_open_document(word, doc_path, read_only=False):
    """Checks if the document is already open in Word; if not, opens it."""
    abs_path = os.path.abspath(doc_path).lower()
    for doc in word.Documents:
        try:
            if os.path.abspath(doc.FullName).lower() == abs_path:
                return doc, False # False means it was already open
        except: continue
    doc = word.Documents.Open(abs_path, ReadOnly=read_only, AddToRecentFiles=False)
    return doc, True # True means we opened it

def ai_rewrite(doc_path, original_text, mode="narrative", custom_prompt=None, model=None, context_text=None):
    """
    Finds original_text in doc_path, rewrites it via Gemini, and replaces it.
    Uses context_text (entire doc) for better AI understanding.
    """
    word = None
    doc = None
    we_created_word = False
    we_opened_doc = False
    success = False
    
    print(f"--- AI Rewrite Started ---")
    print(f"Doc: {doc_path}")
    print(f"Mode: {mode}")
    if model:
        print(f"Model: {model}")
    
    if not original_text or len(original_text.strip()) < 5:
        print("Error: Selection too short or empty.")
        return False

    try:
        # 1. Query Gemini
        print("Querying AI Model...")
        
        system_instruction = "You are a professional legal report assistant."
        
        # Build the background context part of the prompt
        context_block = ""
        if context_text:
            context_block = f"BACKGROUND CONTEXT (Entire Document):\n{context_text}\n\n"

        if custom_prompt:
            prompt = f"{context_block}INSTRUCTION: {custom_prompt}\n\nTARGET TEXT TO REWRITE:\n{original_text}\n\nOutput ONLY the rewritten version of the target text."
        elif mode == "narrative":
            prompt = (
                f"{context_block}INSTRUCTION: Rewrite the following target text to be a cohesive, professional narrative. "
                "Use the provided background context to ensure factual accuracy and consistent tone. "
                "Maintain all factual details and use a formal, objective tone. Output ONLY the rewritten text.\n\n"
                f"TARGET TEXT:\n{original_text}"
            )
        else:
            prompt = (
                f"{context_block}INSTRUCTION: Improve the following target text for clarity, professionalism, and flow. "
                "Ensure it fits perfectly within the provided background context. Output ONLY the improved text.\n\n"
                f"TARGET TEXT:\n{original_text}"
            )
            
        target_models = [model] if model else ["models/gemini-3-flash-preview", "models/gemini-3-pro-preview", "models/gemini-2.5-flash"]
        new_text = call_gemini_api(prompt, context_text=None, models=target_models) # context_text is already in prompt
        
        if not new_text:
            print("Error: AI failed to generate a response.")
            return False

        # ... (rest of the function stays the same)
        new_text = new_text.strip().replace('\n', '\r')
        print("AI Response received and normalized.")


        # 2. Open Word and Replace
        print("Opening Word document...")
        word, we_created_word = get_word_app()
        word.DisplayAlerts = 0 # wdAlertsNone
        
        doc, we_opened_doc = get_or_open_document(word, doc_path, read_only=False)
        
        # We need to be careful with multi-paragraph replacement.
        # Word's Find tool works best for strings within a single paragraph or 
        # short strings. If original_text is multi-paragraph, we might need a 
        # more robust search.
        
        # Normalize text for searching (Word uses \r for paragraph breaks)
        search_text = original_text.replace('\n', '\r').strip()
        
        # Try finding the exact text first
        rng = doc.Content
        fnd = rng.Find
        fnd.ClearFormatting()
        
        # Normalize search text for Word
        search_text = original_text.replace('\n', '\r').strip()
        
        success = False
        if len(search_text) < 250:
            success = fnd.Execute(FindText=search_text)
        
        if success:
            print("Exact match found. Replacing...")
            # Capture start position before replacement
            start_pos = rng.Start
            rng.Text = new_text
            
            # Apply formatting for narrative mode
            if mode == "narrative":
                try:
                    # 1. Ensure blank line above
                    try:
                        p1 = doc.Range(start_pos, start_pos + 1).Paragraphs(1)
                        if p1.Index > 1:
                            prev_p = doc.Paragraphs(p1.Index - 1)
                            if prev_p.Range.Text.strip():
                                # Previous paragraph is not empty, insert a blank line
                                p1.Range.InsertBefore("\r")
                                # Shift start pos for the actual text formatting
                                start_pos += 1
                    except: pass

                    # 2. Apply formatting to each paragraph in the range
                    final_rng = doc.Range(start_pos, start_pos + len(new_text))
                    for p in final_rng.Paragraphs:
                        if p.Range.Text.strip():
                            p.Format.LeftIndent = 0
                            p.Format.FirstLineIndent = word.InchesToPoints(0.5)
                        else:
                            p.Format.LeftIndent = 0
                            p.Format.FirstLineIndent = 0
                except Exception as fe:
                    print(f"Warning: Failed to apply paragraph formatting: {fe}")
            
            success = True
        else:
            print("Exact match not found (or text too long). Attempting fuzzy paragraph match...")
            
            def clean_text(t):
                # Remove common non-matching characters and normalize whitespace
                t = re.sub(r'[\r\n\t\s]+', ' ', t)
                return t.strip()

            def robust_match(anchor, target_text, is_start=True):
                if not anchor: return False
                if anchor in target_text:
                    return True
                # Try stripping common list markers and symbols from anchor
                # e.g. "- Text" -> "Text"
                if is_start:
                    clean_anchor = re.sub(r'^[^a-zA-Z0-9]+', '', anchor).strip()
                else:
                    clean_anchor = re.sub(r'[^a-zA-Z0-9]+$', '', anchor).strip()
                
                if clean_anchor and len(clean_anchor) > 4:
                    return clean_anchor in target_text
                return False

            paras = [p for p in doc.Paragraphs]
            start_para = -1
            end_para = -1
            
            lines = [line.strip() for line in original_text.split('\n') if line.strip()]
            if not lines:
                return False
                
            start_anchor = clean_text(lines[0])[:60] 
            end_anchor = clean_text(lines[-1])[-60:]
            
            print(f"Searching for Start Anchor: '{start_anchor}'")
            print(f"Searching for End Anchor: '{end_anchor}'")

            for i, p in enumerate(paras):
                p_text_clean = clean_text(p.Range.Text)
                
                if start_para == -1 and robust_match(start_anchor, p_text_clean, is_start=True):
                    start_para = i
                    print(f"Found Start Anchor in Paragraph {i+1}")
                
                if start_para != -1 and robust_match(end_anchor, p_text_clean, is_start=False):
                    end_para = i
                    print(f"Found End Anchor in Paragraph {i+1}")
                    break
            
            if start_para != -1 and end_para != -1:
                print(f"Found match range: Paragraphs {start_para+1} to {end_para+1}")
                range_start = paras[start_para].Range.Start
                range_end = paras[end_para].Range.End
                
                final_range = doc.Range(range_start, range_end)
                final_range.Text = new_text + "\r"
                
                # Apply formatting for narrative mode
                if mode == "narrative":
                    try:
                        # 1. Ensure blank line above
                        try:
                            p1 = doc.Range(range_start, range_start + 1).Paragraphs(1)
                            if p1.Index > 1:
                                prev_p = doc.Paragraphs(p1.Index - 1)
                                if prev_p.Range.Text.strip():
                                    p1.Range.InsertBefore("\r")
                                    range_start += 1
                        except: pass

                        # 2. Apply formatting to each paragraph in the range
                        new_content_range = doc.Range(range_start, range_start + len(new_text))
                        for p in new_content_range.Paragraphs:
                            if p.Range.Text.strip():
                                p.Format.LeftIndent = 0
                                p.Format.FirstLineIndent = word.InchesToPoints(0.5)
                            else:
                                p.Format.LeftIndent = 0
                                p.Format.FirstLineIndent = 0
                    except Exception as fe:
                        print(f"Warning: Failed to apply paragraph formatting: {fe}")
                
                success = True
            else:
                # Last ditch: search for start anchor only and replace that paragraph?
                # Probably too dangerous.
                print("Failed to locate text in document.")
                success = False

        if success:
            print("Saving output...")
            doc.Save()
            print("Completed successfully.")
            return True
        else:
            return False

    except Exception as e:
        print(f"CRITICAL ERROR in AI Rewrite: {e}")
        traceback.print_exc()
        return False
    finally:
        if doc and we_opened_doc:
            try:
                doc.Close(SaveChanges=True if success else False)
            except:
                pass
        if word and we_created_word:
            try:
                if word.Documents.Count == 0:
                    word.Quit()
            except:
                pass

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("doc_path", help="Path to Word document")
    parser.add_argument("text_file", help="Path to a temp file containing the original text")
    parser.add_argument("--mode", default="narrative", help="Rewrite mode")
    parser.add_argument("--model", help="Gemini model to use")
    parser.add_argument("--prompt_file", help="Path to a temp file containing a custom prompt")
    parser.add_argument("--context_file", help="Path to a temp file containing the full document context")
    
    args = parser.parse_args()
    
    if not os.path.exists(args.text_file):
        print(f"Error: Text file not found at {args.text_file}")
        sys.exit(1)
        
    with open(args.text_file, 'r', encoding='utf-8') as f:
        original_text = f.read()
        
    custom_prompt = None
    if args.prompt_file and os.path.exists(args.prompt_file):
        with open(args.prompt_file, 'r', encoding='utf-8') as f:
            custom_prompt = f.read()
            
    context_text = None
    if args.context_file and os.path.exists(args.context_file):
        with open(args.context_file, 'r', encoding='utf-8') as f:
            context_text = f.read()
        
    success = ai_rewrite(os.path.abspath(args.doc_path), original_text, args.mode, custom_prompt, args.model, context_text)
    if not success:
        sys.exit(1)

