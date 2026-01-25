import os
import sys
import json
import win32com.client as win32
import time

def create_test_doc(path):
    word = win32.Dispatch("Word.Application")
    doc = word.Documents.Add()
    doc.Range().Text = "Paragraph 1\rParagraph 2\rParagraph 3"
    doc.SaveAs(path)
    doc.Close()
    word.Quit()
    print(f"Test document created at {path}")

def check_formatting(path):
    word = win32.Dispatch("Word.Application")
    doc = word.Documents.Open(path)
    results = []
    for para in doc.Paragraphs:
        results.append(para.SpaceAfter)
    doc.Close()
    word.Quit()
    return results

if __name__ == "__main__":
    test_doc = os.path.abspath("formatting_test.docx")
    rules_file = os.path.abspath("formatting_rules.json")
    
    # 1. Create Rules
    rules = [{
        "name": "Add Double Spacing",
        "enabled": True,
        "trigger": {"scope": "paragraph", "match_type": "regex", "pattern": ".*"},
        "action": {"type": "format", "formatting": {"space_after": 24}}
    }]
    with open(rules_file, "w") as f:
        json.dump(rules, f)
        
    # 2. Create Doc
    if os.path.exists(test_doc): os.remove(test_doc)
    create_test_doc(test_doc)
    
    # 3. Run Rule Engine
    print("\n--- Running Rule Engine ---")
    script_path = os.path.abspath("Scripts/rule_engine.py")
    os.system(f'python "{script_path}" --apply "{test_doc}" "{rules_file}"')
    
    # 4. Verify
    print("\n--- Verification ---")
    spaces = check_formatting(test_doc)
    print(f"Actual SpaceAfter values: {spaces}")
    
    if all(s == 24 for s in spaces):
        print("TEST PASSED: Formatting applied correctly.")
    else:
        print("TEST FAILED: Formatting not applied.")
