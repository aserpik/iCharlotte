import sys
import os
import json
import re
import argparse
import traceback
import win32com.client as win32
import time
import shutil

def to_points(value, unit="in"):
    """Converts inches to points (1 inch = 72 points)."""
    try:
        if unit == "in":
            return float(value) * 72
        return float(value)
    except:
        return 0

def check_dynamic_properties(target, properties):
    """
    Checks if the target COM object matches all properties in the dictionary.
    Returns True if ALL match, False otherwise.
    """
    for path, expected_value in properties.items():
        try:
            parts = path.split('.')
            obj = target
            # Navigate
            for part in parts[:-1]:
                obj = getattr(obj, part)
            
            # Get Value
            final_prop = parts[-1]
            actual_value = getattr(obj, final_prop)
            
            # Comparison
            # Handle float precision
            if isinstance(expected_value, (float, int)) and isinstance(actual_value, (float, int)):
                if abs(actual_value - expected_value) > 0.1:
                    return False
            # Handle string case-insensitivity? optional, but strict for now
            elif str(actual_value) != str(expected_value):
                 # Try typical boolean mismatch (0 vs False, -1 vs True)
                 if isinstance(expected_value, bool):
                     if expected_value and actual_value != 0: continue # True match
                     if not expected_value and actual_value == 0: continue # False match
                 return False
                 
        except Exception:
            return False
            
    return True

def apply_dynamic_properties(target, properties):
    """
    Applies a dictionary of "Path.To.Property": Value to the target COM object.
    path is relative to target (e.g. "Range.Font.Bold").
    """
    changed = False
    for path, value in properties.items():
        try:
            parts = path.split('.')
            obj = target
            # Navigate to the parent of the final property
            for part in parts[:-1]:
                obj = getattr(obj, part)
            
            # Set the final property
            final_prop = parts[-1]
            current_val = getattr(obj, final_prop)
            
            # Simple equality check (with loose typing for some COM objects)
            if current_val != value:
                setattr(obj, final_prop, value)
                changed = True
        except Exception as e:
            print(f"  [Warning] Failed to set dynamic property '{path}': {e}")
            
    return changed

def apply_formatting(target, formatting):
    """
    Applies formatting dictionary to a Word object (Paragraph or Range).
    """
    if not formatting:
        return False
        
    changed = False

    # 4. Dynamic Properties (New System)
    if "dynamic_properties" in formatting:
        if apply_dynamic_properties(target, formatting["dynamic_properties"]):
            changed = True
    
    # 1. Style
    if "style" in formatting:
        try:
            target.Style = formatting["style"]
            changed = True
        except: pass
        
    # 2. Paragraph Properties
    props = {
        "left_indent": "LeftIndent",
        "right_indent": "RightIndent",
        "first_line_indent": "FirstLineIndent",
        "space_before": "SpaceBefore",
        "space_after": "SpaceAfter",
        "line_spacing": "LineSpacing"
    }
    
    for key, com_prop in props.items():
        if key in formatting:
            val = formatting[key]
            if "indent" in key:
                val = to_points(val, "in")
            
            current = getattr(target, com_prop)
            if abs(current - val) > 0.1:
                setattr(target, com_prop, val)
                changed = True
            
            if key in ["space_before", "space_after"]:
                try:
                    if target.NoSpaceBetweenParagraphsOfSameStyle:
                        target.NoSpaceBetweenParagraphsOfSameStyle = False
                        changed = True
                except: pass
            
    if "alignment" in formatting:
        align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
        val = align_map.get(formatting["alignment"].lower(), 0)
        if target.Alignment != val:
            target.Alignment = val
            changed = True

    # 3. Font Properties
    font_props = {
        "font_name": "Name",
        "font_size": "Size",
        "font_bold": "Bold",
        "font_italic": "Italic",
        "font_color": "Color"
    }
    
    font_obj = target.Range.Font if hasattr(target, "Range") else target.Font
    for key, com_prop in font_props.items():
        if key in formatting:
            val = formatting[key]
            current = getattr(font_obj, com_prop)
            if current != val:
                setattr(font_obj, com_prop, val)
                changed = True
            
    return changed

def get_word_app():
    """
    Tries to get the active Word instance, otherwise creates a new one.
    """
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
    """
    Checks if the document is already open in Word; if not, opens it.
    """
    abs_path = os.path.abspath(doc_path).lower()
    try:
        for doc in word.Documents:
            try:
                if os.path.abspath(doc.FullName).lower() == abs_path:
                    return doc, False # False means it was already open
            except: continue
    except:
        # If Word is busy (dialog open), word.Documents might fail.
        # We'll fall through and let word.Documents.Open try (it will also fail if busy, but it's consistent)
        pass
        
    doc = word.Documents.Open(abs_path, ReadOnly=read_only, AddToRecentFiles=False)
    return doc, True # True means we opened it

def apply_rules(doc_path, rules_path):
    word = None
    doc = None
    we_created_word = False
    we_opened_doc = False
    temp_save_path = doc_path + ".tmp.docx"
    count_changes = 0
    
    try:
        print(f"--- Rule Engine Started ---")
        if not os.path.exists(doc_path):
            print(f"Error: Doc not found: {doc_path}")
            return

        with open(rules_path, 'r') as f:
            rules = json.load(f)
            
        print(f"Loaded {len(rules)} rules.")

        word, we_created_word = get_word_app()
        word.DisplayAlerts = 0 # wdAlertsNone
        
        doc, we_opened_doc = get_or_open_document(word, doc_path, read_only=False)
        print(f"Working on: {doc.FullName}")

        # ... (rest of the logic)
        # Separate rules
        global_replace_rules = [r for r in rules if r.get('enabled', True) and r.get('trigger', {}).get('scope') == 'all_text' and r.get('action', {}).get('type') == 'replace']
        para_rules = [r for r in rules if r.get('enabled', True) and r.get('trigger', {}).get('scope', 'paragraph') == 'paragraph']

        # 1. Global Replace (Fast)
        for rule in global_replace_rules:
            trigger = rule.get('trigger', {})
            action = rule.get('action', {})
            pattern = trigger.get('pattern', '')
            replacement = action.get('replacement', '')
            
            print(f"Applying Global Rule: {rule.get('name')}")
            rng = doc.Content
            fnd = rng.Find
            fnd.ClearFormatting()
            fnd.Replacement.ClearFormatting()
            
            if fnd.Execute(FindText=pattern, ReplaceWith=replacement, Replace=2, 
                           MatchCase=trigger.get('case_sensitive', False),
                           MatchWholeWord=trigger.get('whole_word', False)):
                print(f"  [SUCCESS] Global replacement applied.")
                count_changes += 1
                doc.Saved = False

        # 2. Paragraph Rules (Iterative)
        if para_rules:
            paras = [p for p in doc.Paragraphs]
            for i, para in enumerate(paras):
                txt = para.Range.Text.rstrip('\r')
                
                # Get List String if any
                list_str = ""
                try:
                    if para.Range.ListFormat.ListType != 0:
                        list_str = para.Range.ListFormat.ListString
                except: pass
                
                full_txt = (list_str + " " + txt).strip() if list_str else txt

                if not txt.strip() and not any(r.get('trigger', {}).get('pattern') == '.*' for r in para_rules):
                    continue

                for rule in para_rules:
                    trigger = rule.get('trigger', {})
                    action = rule.get('action', {})
                    pattern = trigger.get('pattern', '')
                    match_type = trigger.get('match_type', 'contains')
                    
                    match_found = False
                    check_pat = pattern if trigger.get('case_sensitive') else pattern.lower()
                    
                    # Try matching against raw text AND full text (with numbering)
                    for text_to_check in [txt, full_txt]:
                        check_txt = text_to_check if trigger.get('case_sensitive') else text_to_check.lower()
                        
                        if match_type == 'contains':
                            if trigger.get('whole_word'):
                                if re.search(r'\b' + re.escape(check_pat) + r'\b', check_txt): match_found = True
                            else:
                                if check_pat in check_txt: match_found = True
                        elif match_type == 'starts_with' and check_txt.lstrip().startswith(check_pat):
                            match_found = True
                        elif match_type == 'regex':
                            try:
                                if re.search(pattern, text_to_check, re.IGNORECASE if not trigger.get('case_sensitive') else 0):
                                    match_found = True
                            except: pass
                        
                        if match_found: break # Found it

                    if not match_found:
                        continue

                    # Check List Format Trigger (if specified)
                    is_list_trigger = trigger.get('is_list')
                    if is_list_trigger is not None:
                        try:
                            # 0 = wdListNoNumbering
                            is_list_actual = (para.Range.ListFormat.ListType != 0)
                            if is_list_trigger != is_list_actual:
                                match_found = False
                        except:
                            match_found = False
                    
                    # Check List String Regex (if specified)
                    list_string_regex = trigger.get('list_string_regex')
                    if list_string_regex and match_found:
                        try:
                            if para.Range.ListFormat.ListType == 0:
                                match_found = False
                            else:
                                l_str = para.Range.ListFormat.ListString
                                if not re.search(list_string_regex, l_str):
                                    match_found = False
                        except:
                            match_found = False

                    # Check Property Match (New Hybrid Trigger)
                    prop_match = trigger.get('property_match')
                    if prop_match and match_found:
                        if not check_dynamic_properties(para, prop_match):
                            match_found = False

                    if match_found:
                        if action.get('type') in ['format', 'format_advanced']:
                            if apply_formatting(para, action.get('formatting', {})):
                                print(f"  [SUCCESS] Para {i+1} formatted (Rule: {rule.get('name')})")
                                count_changes += 1
                                doc.Saved = False
                        
                        # Only run replacement if specifically requested
                        if action.get('type') in ['replace', 'cycle'] or 'replacement' in action:
                            replacement = None
                            if action.get('type') == 'cycle':
                                variations = action.get('variations', [])
                                if variations:
                                    replacement = variations[0] # Fallback
                                    for idx, v in enumerate(variations):
                                        if v in txt:
                                            replacement = variations[(idx + 1) % len(variations)]
                                            break
                            else:
                                replacement = action.get('replacement', None)
                            
                            if replacement is not None:
                                if match_type == 'regex':
                                    print(f"  [Info] Skipping text replacement for Regex rule '{rule.get('name')}' to avoid Word crash.")
                                else:
                                    text_before = para.Range.Text
                                    fnd = para.Range.Find
                                    if fnd.Execute(FindText=pattern, ReplaceWith=replacement, Replace=2,
                                                MatchCase=trigger.get('case_sensitive', False),
                                                MatchWholeWord=trigger.get('whole_word', False)):
                                        if para.Range.Text != text_before:
                                            print(f"  [SUCCESS] Para {i+1} replaced (Rule: {rule.get('name')})")
                                            count_changes += 1
                                            doc.Saved = False
                                            txt = para.Range.Text.rstrip('\r')

        print(f"Total changes recorded: {count_changes}")
        
        if doc.Saved == False:
            if we_opened_doc:
                if os.path.exists(temp_save_path): os.remove(temp_save_path)
                doc.SaveAs2(FileName=temp_save_path, FileFormat=16)
                doc.Close(SaveChanges=False)
                if we_created_word: word.Quit()
                time.sleep(0.5)
                shutil.copy2(temp_save_path, doc_path)
                os.remove(temp_save_path)
                print("Document updated successfully.")
            else:
                # If it was already open by user, we just Save it in place
                doc.Save()
                print("Document saved in active Word instance.")
        else:
            print("No changes needed.")
            if we_opened_doc: doc.Close(SaveChanges=False)
            if we_created_word: word.Quit()

    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        if doc and we_opened_doc: doc.Close(SaveChanges=False)
        if word and we_created_word: word.Quit()

def convert_to_html(doc_path, html_path):
    word = None
    doc = None
    temp_doc_path = html_path + f".{int(time.time())}.temp.docx"
    try:
        # 1. Create a copy to avoid UI interference
        if os.path.exists(temp_doc_path): os.remove(temp_doc_path)
        shutil.copy2(doc_path, temp_doc_path)

        # 2. Use a dedicated, isolated instance for previews
        try:
            word = win32.DispatchEx("Word.Application")
        except:
            word = win32.Dispatch("Word.Application")
            
        word.Visible = False
        word.DisplayAlerts = 0
        
        doc = word.Documents.Open(os.path.abspath(temp_doc_path), ReadOnly=True, AddToRecentFiles=False)
        
        # Format 8 = wdFormatHTML, more script-friendly than MHTML
        try:
            doc.SaveAs2(os.path.abspath(html_path), FileFormat=8)
        except:
            doc.SaveAs(os.path.abspath(html_path), FileFormat=8)
            
        print(f"Preview generated for: {doc_path}")
    except Exception as e:
        print(f"Preview Error: {e}")
        traceback.print_exc()
    finally:
        if doc: 
            try: doc.Close(SaveChanges=False)
            except: pass
        if word:
            try: word.Quit()
            except: pass
        if os.path.exists(temp_doc_path):
            try: os.remove(temp_doc_path)
            except: pass
        
        # Cleanup associated Word HTML folder if it exists
        folder_path = html_path.replace(".html", ".files").replace(".mhtml", ".files")
        if os.path.exists(folder_path):
            try: shutil.rmtree(folder_path)
            except: pass

def save_preview_to_docx(preview_path, doc_path):
    word = None
    doc = None
    we_created_word = False
    we_opened_doc = False
    try:
        word, we_created_word = get_word_app()
        word.DisplayAlerts = 0 # wdAlertsNone
        
        # Open the target document
        doc, we_opened_doc = get_or_open_document(word, doc_path, read_only=False)
        
        # Clear existing content
        doc.Content.Delete()
        
        # Insert the MHTML/HTML file content
        doc.Content.InsertFile(os.path.abspath(preview_path))
        
        doc.Save()
        print(f"Document updated successfully from preview.")
        return True
    except Exception as e:
        print(f"Save Error: {e}")
        traceback.print_exc()
        return False
    finally:
        if doc and we_opened_doc: doc.Close()
        if word and we_created_word: word.Quit()

def get_selection_formatting():
    """
    Retrieves formatting from the active selection in Word.
    """
    word = None
    try:
        # Try to get existing instance
        word = win32.GetActiveObject("Word.Application")
        if not word or word.Documents.Count == 0:
            return {"error": "No active Word document found."}
        
        sel = word.Selection
        para = sel.Paragraphs(1) if sel.Paragraphs.Count > 0 else None
        
        fmt = {}
        
        # Paragraph Formatting
        if para:
            fmt["left_indent"] = round(para.LeftIndent / 72.0, 2) # points to inches
            fmt["space_after"] = para.SpaceAfter
            align_map = {0: "left", 1: "center", 2: "right", 3: "justify"}
            fmt["alignment"] = align_map.get(para.Alignment, "left")
            try: fmt["style"] = str(para.Style)
            except: pass

        # Font Formatting
        font = sel.Font
        fmt["font_name"] = font.Name
        fmt["font_size"] = font.Size
        fmt["font_bold"] = bool(font.Bold)
        fmt["font_italic"] = bool(font.Italic)
        
        return fmt
    except Exception as e:
        return {"error": str(e)}

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--preview", action="store_true")
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--save-preview", action="store_true")
    parser.add_argument("--get-selection", action="store_true")
    parser.add_argument("doc_path", nargs='?')
    parser.add_argument("extra_arg", nargs='?')
    args = parser.parse_args()
    
    if args.get_selection:
        print(json.dumps(get_selection_formatting(), indent=2))
    elif args.preview:
        convert_to_html(os.path.abspath(args.doc_path), os.path.abspath(args.extra_arg))
    elif args.apply:
        apply_rules(os.path.abspath(args.doc_path), os.path.abspath(args.extra_arg))
    elif args.save_preview:
        # In this mode, doc_path is the preview source and extra_arg is the target DOCX
        save_preview_to_docx(os.path.abspath(args.doc_path), os.path.abspath(args.extra_arg))