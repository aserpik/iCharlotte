import os
import sys
import json
import argparse
import datetime
import logging
from typing import List, Dict

# Add Scripts to path to import gemini_utils
sys.path.append(os.path.dirname(__file__))
import gemini_utils

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from docx import Document
except ImportError:
    Document = None

# Constants
RULES_FILE = os.path.join(os.path.dirname(__file__), "..", ".gemini", "org_rules.json")
LEARNED_FILE = os.path.join(os.path.dirname(__file__), "..", ".gemini", "learned_patterns.json")

def get_file_content(path: str, max_pages: int = 3) -> str:
    """Extracts text content from a file for analysis."""
    ext = os.path.splitext(path)[1].lower()
    content = ""
    try:
        if ext == ".pdf" and fitz:
            doc = fitz.open(path)
            num_pages = min(len(doc), max_pages)
            for i in range(num_pages):
                content += doc[i].get_text()
            doc.close()
        elif ext == ".docx" and Document:
            doc = Document(path)
            paragraphs = [p.text for p in doc.paragraphs]
            content = "\n".join(paragraphs[:100]) # First 100 paragraphs
        elif ext in [".txt", ".log", ".csv"]:
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read(10000) # First 10k chars
    except Exception as e:
        print(f"Error reading {path}: {e}", file=sys.stderr)
    return content

def load_json(path: str, default):
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return default
    return default

def dry_run(file_paths: List[str]):
    print("Starting organization analysis...", flush=True)
    rules = load_json(RULES_FILE, {})
    learned = load_json(LEARNED_FILE, [])
    
    # Take last 10 learned patterns as few-shot examples
    examples = learned[-10:] if learned else []
    
    prompt = f"""
You are a legal document organizer. Your task is to analyze the content of a document and suggest a standardized filename and target folder based on the provided rules and examples.

[RULES]
Naming Conventions: {json.dumps(rules.get('naming_conventions', {}), indent=2)}
Target Folders: {json.dumps(rules.get('target_folders', {}), indent=2)}
Default Folder: {rules.get('default_folder', 'MISC')}

[LEARNED PATTERNS (EXAMPLES)]
{json.dumps(examples, indent=2)}

[INSTRUCTIONS]
1. Identify the Document Type (e.g., Correspondence, Pleading, Discovery, Medical, Invoice, Report).
2. Extract relevant metadata (Date, Subject, Pleading Type, Vendor, Provider, etc.).
3. Suggest a new filename following the naming convention for that type.
4. Suggest a target folder based on the rules.
5. Provide a brief reasoning for your choice.

Return ONLY a JSON object with the following fields:
{{
  "doc_type": "Type",
  "suggested_name": "New Filename.pdf",
  "suggested_folder": "FOLDER NAME",
  "reasoning": "Brief explanation"
}}
"""

    results = []
    for path in file_paths:
        if not os.path.exists(path):
            continue
            
        filename = os.path.basename(path)
        print(f"Processing {filename}...", flush=True)
        content = get_file_content(path)
        
        if not content.strip():
            # Fallback to filename analysis if no content
            content = f"[No content extracted. Filename: {filename}]"
            
        response = gemini_utils.call_gemini_api(prompt, context_text=f"Filename: {filename}\nContent:\n{content}")
        
        if response:
            try:
                cleaned = gemini_utils.clean_json_string(response)
                suggestion = json.loads(cleaned)
                suggestion["original_path"] = path
                results.append(suggestion)
            except Exception as e:
                results.append({
                    "original_path": path,
                    "suggested_name": filename,
                    "suggested_folder": "ERROR",
                    "reasoning": f"Failed to parse AI response: {e}",
                    "error": True
                })
        else:
            results.append({
                "original_path": path,
                "suggested_name": filename,
                "suggested_folder": "FAILED",
                "reasoning": "AI Model failed to respond.",
                "error": True
            })
            
    print("<<<JSON_START>>>")
    print(json.dumps(results, indent=2))
    print("<<<JSON_END>>>")

def apply_changes(changes_json: str):
    try:
        changes = json.loads(changes_json)
    except Exception as e:
        print(f"Error parsing changes JSON: {e}", file=sys.stderr)
        return

    learned = load_json(LEARNED_FILE, [])
    
    for change in changes:
        orig_path = change.get("original_path")
        new_name = change.get("suggested_name")
        new_folder = change.get("suggested_folder")
        
        if not orig_path or not os.path.exists(orig_path):
            continue
            
        case_dir = os.path.dirname(orig_path)
        # We assume the target folder is relative to the case root.
        # However, iCharlotte tree shows the whole case structure.
        # Let's find the case root by looking for common folders or just use current dir logic.
        # Better: iCharlotte should provide the base_path.
        
        # For now, let's assume we move relative to the original file's parent or a known root.
        # Actually, let's just use the absolute path if suggested_folder is absolute, 
        # or relative to original parent if not.
        
        # To be safe, the UI should probably pass the target_dir.
        target_dir = change.get("target_dir", os.path.dirname(orig_path))
        if new_folder:
            # If new_folder is provided, we might want to create it inside target_dir
            final_dir = os.path.join(target_dir, new_folder)
        else:
            final_dir = target_dir
            
        if not os.path.exists(final_dir):
            os.makedirs(final_dir, exist_ok=True)
            
        new_path = os.path.join(final_dir, new_name)
        
        try:
            os.rename(orig_path, new_path)
            print(f"Moved: {orig_path} -> {new_path}", flush=True)
            
            # Save to learned patterns (excluding reasoning to keep it clean)
            learned.append({
                "original_filename": os.path.basename(orig_path),
                "doc_type": change.get("doc_type"),
                "suggested_name": new_name,
                "suggested_folder": new_folder
            })
        except Exception as e:
            print(f"Failed to move {orig_path}: {e}", file=sys.stderr)

    # Save learned patterns
    try:
        with open(LEARNED_FILE, 'w', encoding='utf-8') as f:
            json.dump(learned, f, indent=2)
    except Exception as e:
        print(f"Failed to save learned patterns: {e}", file=sys.stderr)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="AI File Organizer")
    parser.add_argument("--dry-run", action="store_true", help="Analyze files and suggest changes")
    parser.add_argument("--apply", help="Apply changes from JSON string")
    parser.add_argument("files", nargs="*", help="Files to process")
    
    args = parser.parse_args()
    
    if args.dry_run:
        if not args.files:
            print("No files provided for dry-run", file=sys.stderr)
            sys.exit(1)
        dry_run(args.files)
    elif args.apply:
        apply_changes(args.apply)
    else:
        parser.print_help()
