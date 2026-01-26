import os
import sys
import logging
import datetime
import re
import subprocess
import time
import gc
from google import genai
from docx import Document
from docx.shared import Pt, Inches
from pypdf import PdfReader

# Import Case Data Manager
try:
    from case_data_manager import CaseDataManager
except ImportError:
    sys.path.append(os.path.join(os.getcwd(), 'Scripts'))
    from case_data_manager import CaseDataManager

try:
    import pytesseract
    from pdf2image import convert_from_path
    OCR_AVAILABLE = True
    
    # Windows-specific path configuration
    if os.name == 'nt':
        # Tesseract Path
        tesseract_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if os.path.exists(tesseract_path):
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
        
        # Poppler Path
        # We found it at C:\Program Files\poppler\Library\bin\pdftoppm.exe
        # pdf2image needs the folder containing the binaries
        poppler_path = r"C:\Program Files\poppler\Library\bin"
        if not os.path.exists(poppler_path):
             # Fallback attempt or standard location
             poppler_path = None
        
        # We will use this variable in extract_text
        POPPLER_PATH = poppler_path
    else:
        POPPLER_PATH = None

except ImportError:
    OCR_AVAILABLE = False
    POPPLER_PATH = None


# --- Configuration ---
# Hardcoded log path as per instructions
LOG_FILE = r"C:\GeminiTerminal\Summarize_Discovery_activity.log"
PROMPT_FILE = r"C:\GeminiTerminal\Scripts\SUMMARIZE_DISCOVERY_PROMPT.txt"
CONSOLIDATE_PROMPT_FILE = r"C:\GeminiTerminal\Scripts\CONSOLIDATE_DISCOVERY_PROMPT.txt"
EXTRACTION_PROMPT_FILE = r"C:\GeminiTerminal\Scripts\EXTRACTION_PASS_PROMPT.txt"
CROSS_CHECK_PROMPT_FILE = r"C:\GeminiTerminal\Scripts\CROSS_CHECK_PROMPT.txt"

# Set up logging
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

def log_event(message, level="info"):
    try:
        print(message, flush=True)  # Also print to stdout for debugging if run manually
    except UnicodeEncodeError:
        try:
            print(message.encode(sys.stdout.encoding or 'utf-8', errors='replace').decode(sys.stdout.encoding or 'utf-8'), flush=True)
        except Exception:
             print(message.encode('ascii', errors='replace').decode('ascii'), flush=True)
    
    sys.stdout.flush() # Force flush
    if level == "info":
        logging.info(message)
    elif level == "error":
        logging.error(message)
    elif level == "warning":
        logging.warning(message)

def is_verification_page(text):
    """Checks if the text content strongly indicates a verification page."""
    if not text:
        return False
    text_upper = text.upper()
    
    # Strong indicator: Title "VERIFICATION" on its own line or prominent
    lines = [l.strip() for l in text_upper.split('\n') if l.strip()]
    
    # Check first few lines for "VERIFICATION"
    # It might be preceded by attorney info, but usually the title is early.
    # If the document is JUST a verification, "VERIFICATION" is likely the main title.
    
    has_verification_title = False
    for i in range(min(20, len(lines))):
        if lines[i] == "VERIFICATION":
            has_verification_title = True
            break
            
    # Check for perjury language
    has_perjury_lang = "UNDER PENALTY OF PERJURY" in text_upper and ("TRUE AND CORRECT" in text_upper or "EXECUTED ON" in text_upper)
    
    # If we have the title "VERIFICATION" AND the perjury language, it is definitely a verification page.
    # If we just have "VERIFICATION" title, it might be a pleading title, but if it's the FIRST page, 
    # and the user wants to skip files that are "only verifications", then a file starting with "VERIFICATION" title
    # is extremely likely to be just that.
    
    if has_verification_title:
        return True
        
    # If no title, but strong perjury language and short text?
    # A full response will have "RESPONSE TO INTERROGATORIES" etc.
    # If it lacks "RESPONSE TO" but has perjury language, likely a verification.
    
    has_response_title = "RESPONSE TO" in text_upper or "RESPONSES TO" in text_upper
    
    if has_perjury_lang and not has_response_title:
        return True
        
    return False

def extract_text(file_path):
    """Extracts text from PDF, DOCX, or plain text files. Falls back to OCR for specific PDF pages if text is insufficient."""
    log_event(f"Extracting text from: {file_path}")
    ext = os.path.splitext(file_path)[1].lower()
    text = ""

    try:
        if ext == ".pdf":
            reader = PdfReader(file_path)
            total_pages = len(reader.pages)
            log_event(f"PDF has {total_pages} pages.")
            
            # Helper to convert specific page index to image if needed
            def get_page_image(page_index):
                try:
                    # convert_from_path returns a list of images. We need to fetch just the specific page.
                    # 'first_page' and 'last_page' are 1-based indices in pdf2image
                    # Pass poppler_path if explicitly set (mainly for Windows)
                    images = convert_from_path(file_path, first_page=page_index+1, last_page=page_index+1, poppler_path=POPPLER_PATH)
                    return images[0] if images else None
                except Exception as e:
                    log_event(f"Error converting page {page_index+1} to image: {e}", level="warning")
                    return None

            for i, page in enumerate(reader.pages):
                if i % 10 == 0:
                    gc.collect()

                page_text = page.extract_text() or ""
                
                # Check if this specific page has meaningful text (threshold: 50 chars)
                if len(page_text.strip()) < 50:
                    log_event(f"Page {i+1} has insufficient text ({len(page_text.strip())} chars). Attempting OCR...", level="warning")
                    
                    if OCR_AVAILABLE:
                        image = get_page_image(i)
                        if image:
                            try:
                                ocr_page_text = pytesseract.image_to_string(image)
                                if len(ocr_page_text.strip()) > len(page_text.strip()):
                                    log_event(f"OCR successful for page {i+1}. Extracted {len(ocr_page_text)} chars.")
                                    page_text = ocr_page_text # Use OCR text for checking
                                    text += ocr_page_text + "\n"
                                else:
                                    log_event(f"OCR for page {i+1} did not yield better results. Using original.", level="warning")
                                    text += page_text + "\n"
                            except Exception as ocr_e:
                                log_event(f"OCR failed for page {i+1}: {ocr_e}. Using original.", level="error")
                                text += page_text + "\n"
                        else:
                             log_event(f"Could not render image for page {i+1}. Using original.", level="warning")
                             text += page_text + "\n"
                    else:
                        log_event(f"OCR not available. Skipping OCR for page {i+1}.", level="warning")
                        text += page_text + "\n"
                else:
                    # Page text is sufficient
                    text += page_text + "\n"
                
                # Check the FIRST PAGE for Verification status
                if i == 0:
                    # Use the page_text we resolved (original or OCR)
                    if is_verification_page(page_text):
                        log_event(f"Skipping {file_path}: Identified as Verification only (Page 1 detection).")
                        return None

        elif ext == ".docx":
            doc = Document(file_path)
            # Read first few paragraphs for verification check
            first_page_text = ""
            for i, paragraph in enumerate(doc.paragraphs):
                text += paragraph.text + "\n"
                if i < 20: # Rough approximation of first page content
                    first_page_text += paragraph.text + "\n"
            
            if is_verification_page(first_page_text):
                log_event(f"Skipping {file_path}: Identified as Verification only.")
                return None
                
        else:
            # Assume plain text for other extensions
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                text = f.read()
            
            # Check first 2000 chars roughly
            if is_verification_page(text[:2000]):
                log_event(f"Skipping {file_path}: Identified as Verification only.")
                return None
        
        if not text.strip():
            log_event(f"Warning: Extracted text is empty for {file_path}", level="warning")
            return None
        
        log_event(f"Successfully extracted {len(text)} characters total.")
        log_event(f"Snippet: {text[:100].replace(chr(10), ' ')}...") # Log snippet
        return text

    except Exception as e:
        log_event(f"Error extracting text from {file_path}: {e}", level="error")
        return None

def call_gemini(prompt, text):
    """Calls Gemini API with fallback logic."""
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        log_event("Error: GEMINI_API_KEY environment variable not set.", level="error")
        return None

    client = genai.Client(api_key=api_key)

    full_prompt = f"{prompt}\n\nDOCUMENT CONTENT:\n{text}"
    
    models_to_try = ["gemini-1.5-pro", "gemini-2.0-flash-exp", "gemini-1.5-flash"]
    
    model_sequence = [
        "gemini-3-pro-preview", 
        "gemini-3-flash-preview", 
        "gemini-2.5-pro", 
        "gemini-2.5-flash"
    ]

    for model_name in model_sequence:
        log_event(f"Attempting to use model: {model_name}")
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=full_prompt
            )
            if response and response.text:
                log_event(f"Success with model: {model_name}")
                return response.text
        except BaseException as e: # Catch ALL exceptions including system exits
            log_event(f"Failed with model {model_name}: {e}", level="warning")
            continue
    
    log_event("Error: All model attempts failed.", level="error")
    return None

def call_gemini_cross_check(prompt, extraction_text, summary_text):
    """Calls Gemini API for cross-checking two text sections."""
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        log_event("Error: GEMINI_API_KEY environment variable not set.", level="error")
        return None

    client = genai.Client(api_key=api_key)

    full_prompt = f"""{prompt}

=== PASS 1 (EXTRACTED RESPONSES) ===
{extraction_text}

=== PASS 2 (NARRATIVE SUMMARY) ===
{summary_text}
"""

    model_sequence = [
        "gemini-3-pro-preview",
        "gemini-3-flash-preview",
        "gemini-2.5-pro",
        "gemini-2.5-flash"
    ]

    for model_name in model_sequence:
        log_event(f"Cross-check: Attempting model: {model_name}")
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=full_prompt
            )
            if response and response.text:
                log_event(f"Cross-check: Success with model: {model_name}")
                return response.text
        except BaseException as e:
            log_event(f"Cross-check: Failed with model {model_name}: {e}", level="warning")
            continue

    log_event("Cross-check: All model attempts failed.", level="error")
    return None

def add_section_divider(doc, section_title):
    """Adds a visual divider with section title to the document."""
    # Add some spacing
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.0

    # Add horizontal line (using underscores for compatibility)
    divider_line = "─" * 80
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(divider_line)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10)

    # Add section title (bold, centered)
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.0
    p.alignment = 1  # Center alignment
    run = p.add_run(section_title)
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)

    # Add another divider line
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(divider_line)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10)

    # Add spacing after
    doc.add_paragraph()

def add_markdown_to_doc(doc, content):
    """Parses basic Markdown and adds it to the docx Document."""
    lines = content.split('\n')
    
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        
        # Headings: Convert to bold text at start of paragraph, ending with period
        if stripped.startswith('#'):
            text = stripped.lstrip('#').strip()
            if not text.endswith('.'):
                text += "."
            
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.first_line_indent = Inches(0.5)
            run = p.add_run(text)
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            continue
        
        # List items
        if stripped.startswith('* ') or stripped.startswith('- '):
            text = stripped[2:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.left_indent = Inches(0.5)
            p.paragraph_format.first_line_indent = Inches(-0.25)
            
            # Use a real bullet character and tab for portability
            run = p.add_run("•\t")
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            
            # Support bold parsing within list items
            parts = re.split(r'(\*\*.*?\*\*)', text)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    r = p.add_run(part[2:-2])
                    r.bold = True
                else:
                    r = p.add_run(part)
                r.font.name = 'Times New Roman'
                r.font.size = Pt(12)
            continue
        
        # Normal text (with bold support)
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.first_line_indent = Inches(0.5)
        
        # Simple bold parsing: **text**
        # We assume standard markdown usage where ** is not nested
        parts = stripped.split('**')
        for i, part in enumerate(parts):
            run = p.add_run(part)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            # If the index is odd (1, 3, 5...), it was inside **...**
            if i % 2 == 1:
                run.bold = True

def overwrite_docx(content, output_path):
    """Overwrites the DOCX file with new content."""
    try:
        doc = Document()
        log_event(f"Overwriting file with consolidated content: {output_path}")

        # Apply Styles (Times New Roman, Size 12, Single Spaced)
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(12)
        style.paragraph_format.line_spacing = 1.0

        add_markdown_to_doc(doc, content)
        doc.save(output_path)
        log_event(f"Successfully consolidated and saved: {output_path}")
        return True
    except Exception as e:
        log_event(f"Error overwriting DOCX: {e}", level="error")
        return False

def consolidate_file(file_path):
    """Reads the current file content, calls Gemini to consolidate, and overwrites."""
    log_event(f"Starting consolidation for: {file_path}")
    
    # 1. Read existing content
    text = extract_text(file_path)
    if not text:
        log_event("Failed to extract text for consolidation.", level="error")
        return False

    # 2. Load Consolidation Prompt
    try:
        with open(CONSOLIDATE_PROMPT_FILE, "r", encoding="utf-8") as f:
            prompt_instruction = f.read()
    except Exception as e:
        log_event(f"Error reading consolidation prompt file {CONSOLIDATE_PROMPT_FILE}: {e}", level="error")
        return False

    # 3. Call LLM
    consolidated_summary = call_gemini(prompt_instruction, text)
    if not consolidated_summary:
        log_event("Gemini failed to generate consolidated summary.", level="error")
        return False

    # 4. Overwrite file
    return overwrite_docx(consolidated_summary, file_path)

def save_to_docx(extraction_content, summary_content, output_path, title_text, discovery_type="Written Discovery"):
    """Saves both extraction and summary content to a DOCX file with section dividers. Appends if exists. Handles locking."""
    base_name, ext = os.path.splitext(output_path)
    counter = 1
    current_output_path = output_path

    while True:
        try:
            if os.path.exists(current_output_path):
                try:
                    doc = Document(current_output_path)
                    log_event(f"Appending to existing file: {current_output_path}")
                except Exception:
                    # If opening fails, assume locked/corrupt and force next version
                    raise PermissionError("Cannot open existing file.")
            else:
                doc = Document()
                log_event(f"Creating new file: {current_output_path}")

            # Apply Styles (Times New Roman, Size 12, Single Spaced)
            style = doc.styles['Normal']
            style.font.name = 'Times New Roman'
            style.font.size = Pt(12)
            style.paragraph_format.line_spacing = 1.0

            # Count existing subheadings to determine the next number
            # We look for paragraphs that match our subheading pattern
            next_num = 1
            for para in doc.paragraphs:
                if re.match(r'^\d+\.\t.*Responses to', para.text):
                    next_num += 1

            # Determine "Responding Party" from title_text (filename)
            # Remove extension and potentially "discovery" or "responses" keywords
            party_name = title_text
            for skip in [".pdf", ".docx", "discovery", "responses", "response"]:
                party_name = party_name.replace(skip, "").replace(skip.upper(), "")
            party_name = party_name.strip(" _-")

            # Subheading: [Number]. [Party Name]'s Responses to [Discovery Type]
            # Indentation: Number at 1.0", text at 1.5"
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.0
            # Left indent 1.5", First line indent -0.5" (relative to 1.5") = 1.0" for the number
            p.paragraph_format.left_indent = Inches(1.5)
            p.paragraph_format.first_line_indent = Inches(-0.5)

            # Add tab stop at 1.5" for the text following the number
            tab_stops = p.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(1.5))

            # Number and Title
            run_num = p.add_run(f"{next_num}.\t")
            run_num.bold = True
            run_num.font.name = 'Times New Roman'
            run_num.font.size = Pt(12)

            run_title = p.add_run(f"{party_name}'s Responses to {discovery_type}")
            run_title.bold = True
            run_title.underline = True
            run_title.font.name = 'Times New Roman'
            run_title.font.size = Pt(12)

            # Add Section A: Extracted Responses
            add_section_divider(doc, "SECTION A: EXTRACTED RESPONSES")
            add_markdown_to_doc(doc, extraction_content)

            # Add Section B: Narrative Summary
            add_section_divider(doc, "SECTION B: NARRATIVE SUMMARY")
            add_markdown_to_doc(doc, summary_content)

            doc.save(current_output_path)
            log_event(f"Saved summary to: {current_output_path}")
            return True

        except (PermissionError, IOError) as e:
            log_event(f"File locked or inaccessible: {current_output_path}. Trying next version. Error: {e}", level="warning")
            counter += 1
            # Format: Original v.2.docx, Original v.3.docx
            current_output_path = f"{base_name} v.{counter}{ext}"

            if counter > 10:
                log_event("Failed to save after 10 attempts.", level="error")
                return False
        except Exception as e:
            log_event(f"Error saving to DOCX: {e}", level="error")
            return False

def run_concurrently(file_paths, extra_args=None, batch_size=3):
    """
    Runs subprocesses for the given file paths, limiting concurrency to batch_size.
    Waits for all processes to complete.
    """
    log_event(f"Starting concurrent processing for {len(file_paths)} files with batch size {batch_size}...")
    processes = []
    
    for path in file_paths:
        # Manage concurrency
        while len(processes) >= batch_size:
            # Remove finished processes
            processes = [p for p in processes if p.poll() is None]
            if len(processes) >= batch_size:
                time.sleep(1) # Wait before checking again
        
        # Prepare command
        cmd = [sys.executable, sys.argv[0], path]
        if extra_args:
            cmd.extend(extra_args)
            
        creation_flags = 0x08000000 if os.name == 'nt' else 0
        
        try:
            log_event(f"Spawning subprocess for: {path}")
            p = subprocess.Popen(cmd, creationflags=creation_flags)
            processes.append(p)
        except Exception as e:
            log_event(f"Failed to spawn subprocess for {path}: {e}", level="error")
            
    # Wait for remaining processes to finish
    while processes:
        processes = [p for p in processes if p.poll() is None]
        if processes:
            time.sleep(1)
            
    log_event("All concurrent tasks completed.")

def main():
    # Parse and strip custom flags manually to avoid argparse conflict/overhead with loose path args
    args = sys.argv[1:]
    report_dir = None
    clean_args = []
    
    i = 0
    while i < len(args):
        arg = args[i]
        if arg == "--report-dir":
            if i + 1 < len(args):
                report_dir = args[i+1]
                i += 2
            else:
                i += 1
        else:
            if arg.strip():
                clean_args.append(arg)
            i += 1

    if not clean_args:
        # If no paths provided, check if we just had flags (unlikely valid usage, but safety check)
        log_event("Error: No file paths provided.", level="error")
        sys.exit(1)

    # Heuristic to handle split paths: 
    # If multiple args are provided but the first doesn't exist, 
    # check if joining them all back together forms a valid path.
    if len(clean_args) > 1:
        combined_path = " ".join(clean_args).strip().strip('"').strip("'")
        log_event(f"Checking heuristic path: {combined_path}")
        if os.path.exists(combined_path) and not os.path.exists(clean_args[0]):
            log_event(f"Reconstructed split path: {combined_path}")
            file_paths = [combined_path]
        else:
            file_paths = clean_args
    else:
        file_paths = clean_args
    
    # --- Dispatcher Mode ---
    if len(file_paths) > 1:
        log_event(f"Detected multiple file paths: {len(file_paths)} items. Launching separate agents...")
        
        import tempfile
        import shutil
        
        # If report_dir was somehow passed to dispatcher, use it, else create one
        # (Dispatcher usually creates it)
        use_temp_dir = False
        if not report_dir:
            report_dir = tempfile.mkdtemp()
            use_temp_dir = True
            log_event(f"Created temporary report directory: {report_dir}")

        # Use concurrent runner with the report dir flag
        run_concurrently(file_paths, extra_args=["--report-dir", report_dir], batch_size=3)
        
        log_event("Processing consolidation reports...")
        unique_outputs = set()
        
        if os.path.exists(report_dir):
            for f in os.listdir(report_dir):
                if f.startswith("worker_"):
                    try:
                        with open(os.path.join(report_dir, f), 'r') as rf:
                            content = rf.read().strip()
                            if content:
                                unique_outputs.add(content)
                    except Exception as e:
                        log_event(f"Error reading report file {f}: {e}", level="warning")
        
        log_event(f"Found {len(unique_outputs)} unique output files to consolidate.")
        
        # Consolidate each unique file
        for doc_path in unique_outputs:
             if os.path.exists(doc_path):
                 consolidate_file(doc_path)
             else:
                 log_event(f"Output file not found for consolidation: {doc_path}", level="warning")
                 
        if use_temp_dir:
             try:
                 shutil.rmtree(report_dir)
                 log_event("Cleaned up temporary report directory.")
             except Exception as e:
                 log_event(f"Error cleaning up temp dir: {e}", level="warning")

        print(f"Finished processing {len(file_paths)} items.")
        sys.exit(0)

    # --- Worker Mode ---
    # At this point, we assume we are processing a SINGLE path
    
    input_path = file_paths[0]
    
    # Remove surrounding quotes if present (common in CLI args)
    # Use strip to handle unbalanced quotes if they occur
    input_path = input_path.strip().strip('"').strip("'")
    
    # Handle absolute path conversion
    input_path = os.path.abspath(input_path)

    if not os.path.exists(input_path):
        log_event(f"Error: File not found: {input_path}", level="error")
        sys.exit(1)

    log_event(f"--- Starting Agent for input: {input_path} ---")

    # --- Directory Handling ---
    if os.path.isdir(input_path):
        log_event(f"Input is a directory. Scanning for files in: {input_path}")
        files_to_process = []
        for root, _, files in os.walk(input_path):
            for file in files:
                if file.lower().endswith((".pdf", ".docx")):
                    # Exclude the output file itself if it happens to be in the scan path
                    if "Discovery_Response_Summaries" in file or "Discovery_Responses_" in file:
                        continue
                    
                    # Filter out documents with "RFP", "prod", "production", "RFA", "admission" in the filename
                    lower_file = file.lower()
                    if any(keyword in lower_file for keyword in ["rfp", "prod", "production", "rfa", "admission"]):
                        log_event(f"Skipping filtered file: {file}")
                        continue

                    files_to_process.append(os.path.join(root, file))
        
        if not files_to_process:
            log_event("No suitable files (.pdf, .docx) found in directory.")
            sys.exit(0)
        
        log_event(f"Found {len(files_to_process)} files to process.")
        
        # Directory logic also needs to handle the report_dir propagation if it was passed
        # OR if this is the root call for a directory, it acts as a dispatcher.
        
        import tempfile
        import shutil
        
        # If we are here, we might be a "worker" that was given a directory path, 
        # OR the initial call was for a directory.
        # If report_dir IS passed, we are a sub-worker? Unlikely logic flow but safe to handle.
        # Generally, if 'main' is called with a directory, it becomes a dispatcher.
        
        use_temp_dir = False
        if not report_dir:
            report_dir = tempfile.mkdtemp()
            use_temp_dir = True
            
        run_concurrently(files_to_process, extra_args=["--report-dir", report_dir], batch_size=3)
        
        # Only consolidate if we are the owner of the report_dir (the Dispatcher)
        if use_temp_dir:
            log_event("Processing consolidation reports (Directory Mode)...")
            unique_outputs = set()
            if os.path.exists(report_dir):
                for f in os.listdir(report_dir):
                    if f.startswith("worker_"):
                        try:
                            with open(os.path.join(report_dir, f), 'r') as rf:
                                content = rf.read().strip()
                                if content:
                                    unique_outputs.add(content)
                        except Exception:
                            pass
            
            for doc_path in unique_outputs:
                if os.path.exists(doc_path):
                    consolidate_file(doc_path)
            
            try:
                shutil.rmtree(report_dir)
            except:
                pass
                
        # If we were passed a report_dir (we are a sub-process given a directory? Weird but okay),
        # we should propagate the results up. But typically run_concurrently handles the children.
        # The children of THIS process will write to report_dir.
        # WE (this process) don't need to write to report_dir, because we didn't generate a summary ourselves,
        # we just delegated.
        
        log_event("--- Directory Processing Complete ---")
        sys.exit(0)
    
    # --- Single File Processing ---

    # Filter check for single file
    input_filename_lower = os.path.basename(input_path).lower()
    if any(keyword in input_filename_lower for keyword in ["rfp", "prod", "production", "rfa", "admission"]):
        log_event(f"Skipping filtered file: {input_path}")
        sys.exit(0)

    # 1. Extract Text
    text = extract_text(input_path)
    if not text:
        sys.exit(1)

    # ========== TWO-PASS EXTRACTION AND SUMMARIZATION ==========

    # --- PASS 1: Extraction ---
    log_event("Starting Pass 1: Extraction...")
    try:
        with open(EXTRACTION_PROMPT_FILE, "r", encoding="utf-8") as f:
            extraction_prompt = f.read()
    except Exception as e:
        log_event(f"Error reading extraction prompt file {EXTRACTION_PROMPT_FILE}: {e}", level="error")
        sys.exit(1)

    extraction_result = call_gemini(extraction_prompt, text)
    if not extraction_result:
        log_event("Pass 1 (Extraction) failed.", level="error")
        sys.exit(1)
    log_event("Pass 1 (Extraction) completed successfully.")

    # --- PASS 2: Narrative Summary ---
    log_event("Starting Pass 2: Narrative Summary...")
    try:
        with open(PROMPT_FILE, "r", encoding="utf-8") as f:
            summary_prompt = f.read()
    except Exception as e:
        log_event(f"Error reading summary prompt file {PROMPT_FILE}: {e}", level="error")
        sys.exit(1)

    summary_result = call_gemini(summary_prompt, text)
    if not summary_result:
        log_event("Pass 2 (Summary) failed.", level="error")
        sys.exit(1)
    log_event("Pass 2 (Narrative Summary) completed successfully.")

    # Strip extraction layer from summary if present (legacy behavior)
    summary_result = re.sub(r'<extraction_layer>.*?</extraction_layer>', '', summary_result, flags=re.DOTALL).strip()

    # --- PASS 3: Cross-Check ---
    log_event("Starting Pass 3: Cross-Check...")
    try:
        with open(CROSS_CHECK_PROMPT_FILE, "r", encoding="utf-8") as f:
            cross_check_prompt = f.read()
    except Exception as e:
        log_event(f"Error reading cross-check prompt file {CROSS_CHECK_PROMPT_FILE}: {e}", level="error")
        # If cross-check prompt fails, proceed with unchecked summary
        log_event("Proceeding without cross-check.", level="warning")
        final_summary = summary_result
    else:
        cross_check_result = call_gemini_cross_check(cross_check_prompt, extraction_result, summary_result)
        if cross_check_result:
            log_event("Pass 3 (Cross-Check) completed successfully.")
            final_summary = cross_check_result
        else:
            log_event("Pass 3 (Cross-Check) failed. Using unchecked summary.", level="warning")
            final_summary = summary_result

    # Parse Responding Party and Discovery Type from extraction result (more reliable source)
    responding_party = "Unknown_Party"
    discovery_type = "Written Discovery"

    # Check extraction result for metadata
    extraction_lines = extraction_result.split('\n')
    extraction_lines_to_remove = []
    for i in range(min(10, len(extraction_lines))):
        line = extraction_lines[i].strip()
        if line.startswith("RESPONDING_PARTY:"):
            responding_party = line.replace("RESPONDING_PARTY:", "").strip()
            extraction_lines_to_remove.append(i)
        elif line.startswith("DISCOVERY_TYPE:"):
            discovery_type = line.replace("DISCOVERY_TYPE:", "").strip()
            extraction_lines_to_remove.append(i)

    # Remove metadata lines from extraction content
    for i in sorted(extraction_lines_to_remove, reverse=True):
        extraction_lines.pop(i)
    extraction_content = "\n".join(extraction_lines).strip()

    # Also check and clean summary for metadata (in case cross-check included it)
    summary_lines = final_summary.split('\n')
    summary_lines_to_remove = []
    for i in range(min(10, len(summary_lines))):
        line = summary_lines[i].strip()
        if line.startswith("RESPONDING_PARTY:"):
            if responding_party == "Unknown_Party":
                responding_party = line.replace("RESPONDING_PARTY:", "").strip()
            summary_lines_to_remove.append(i)
        elif line.startswith("DISCOVERY_TYPE:"):
            if discovery_type == "Written Discovery":
                discovery_type = line.replace("DISCOVERY_TYPE:", "").strip()
            summary_lines_to_remove.append(i)

    # Remove metadata lines from summary content
    for i in sorted(summary_lines_to_remove, reverse=True):
        summary_lines.pop(i)
    summary_content = "\n".join(summary_lines).strip()

    if responding_party == "Unknown_Party":
        log_event("Warning: RESPONDING_PARTY tag not found in LLM response.", level="warning")

    # Save to Case Data
    data_manager = CaseDataManager()
    
    # Try standard pattern first: 1234.567
    file_num_match = re.search(r"(\d{4}\.\d{3})", input_path)
    file_num = None
    
    if file_num_match:
        file_num = file_num_match.group(1)
    else:
        # Try to parse from directory structure: .../Current Clients/5800 - Client/056 - Matter/...
        # Look for "Current Clients" and then grab the numbers from the next two folders
        parts = os.path.normpath(input_path).split(os.sep)
        try:
            # Find index of "Current Clients" (case-insensitive)
            cc_index = -1
            for i, part in enumerate(parts):
                if part.lower() == "current clients":
                    cc_index = i
                    break
            
            if cc_index != -1 and cc_index + 2 < len(parts):
                client_folder = parts[cc_index + 1]
                matter_folder = parts[cc_index + 2]
                
                # Extract leading digits
                client_code = re.match(r"^\d+", client_folder)
                matter_code = re.match(r"^\d+", matter_folder)
                
                if client_code and matter_code:
                    file_num = f"{client_code.group(0)}.{matter_code.group(0)}"
                    log_event(f"Extracted file number from path structure: {file_num}")
        except Exception as e:
            log_event(f"Error parsing file number from path: {e}", level="warning")

    if file_num:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        clean_var_name = re.sub(r"[^a-zA-Z0-9_]", "_", base_name.lower())
        var_key = f"discovery_summary_{clean_var_name}"
        
        log_event(f"Saving discovery summary to case variable: {var_key} for {file_num}")
        data_manager.save_variable(file_num, var_key, summary_content, source="discovery_agent", extra_tags=["Discovery"])
    else:
        log_event("Could not extract file number from path. Skipping variable save.", level="warning")

    # 4. Save Output
    parts = input_path.split(os.sep)
    output_dir = None
    case_root_parts = None

    # Priority 1: Find folder starting with exactly 3 digits (Case Folder)
    for i in range(len(parts) - 1, -1, -1):
        if re.match(r'^\d{3}(\D|$)', parts[i]):
            log_event(f"Identified Case Folder by 3-digit pattern: {parts[i]}")
            case_root_parts = parts[:i+1]
            break

    # Priority 2: Standard "Current Clients" structure (Client/Case)
    if not case_root_parts:
        for i, part in enumerate(parts):
            if part.lower() == "current clients":
                if i + 2 < len(parts):
                    log_event("Identified Case Folder by Standard Structure (Current Clients + 2)")
                    case_root_parts = parts[:i+3]
                break
    
    if case_root_parts:
        output_dir = os.sep.join(case_root_parts + ["NOTES", "AI OUTPUT"])

    if not output_dir:
        # Fallback 1: If "NOTES" is already in the path, use it.
        for i in range(len(parts) - 1, -1, -1):
            if parts[i].upper() == "NOTES":
                output_dir = os.path.join(os.sep.join(parts[:i+1]), "AI OUTPUT")
                break
    
    if not output_dir:
        # Fallback 2: Sibling NOTES to the file's parent folder (Original fallback)
        input_dir = os.path.dirname(input_path)
        parent_dir = os.path.dirname(input_dir)
        output_dir = os.path.join(parent_dir, "NOTES", "AI OUTPUT")

    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            log_event(f"Created directory: {output_dir}")
        except Exception as e:
            log_event(f"Error creating directory {output_dir}: {e}", level="error")
            sys.exit(1)

    # --- Party Name Matching Logic ---
    def normalize_party_name(name):
        """Normalizes party name for comparison (remove generic titles, special chars)."""
        # Remove common legal prefixes/titles
        titles = ["DEFENDANT", "PLAINTIFF", "RESPONDENT", "CROSS-DEFENDANT", "CROSS-COMPLAINANT", "THE", "AND"]
        cleaned = name.upper()
        for title in titles:
            cleaned = cleaned.replace(title, "")
        # Remove non-alphanumeric (keep spaces for splitting)
        cleaned = re.sub(r"[^A-Z0-9\s]", "", cleaned)
        # Collapse whitespace
        return " ".join(cleaned.split())

    def find_existing_party_file(output_dir, new_party_name):
        """Checks if a file already exists for this party (fuzzy match)."""
        normalized_new = normalize_party_name(new_party_name)
        if len(normalized_new) < 3: # Too short to reliably match
            return None
            
        try:
            existing_files = [f for f in os.listdir(output_dir) if f.startswith("Discovery_Responses_") and f.endswith(".docx")]
        except OSError:
            return None

        best_match_file = None
        # Heuristic: Check for substring inclusion
        # e.g. "Jose Manuel Landeros" (new) matches "Discovery_Responses_Defendant_Jose_Manuel_Landeros_Loera.docx" (existing)
        
        for file in existing_files:
            # Extract name part from filename: Discovery_Responses_{NAME}.docx
            name_part = file[len("Discovery_Responses_"):-5] # remove prefix and .docx
            # Undo the underscores used in filename
            existing_party_raw = name_part.replace("_", " ")
            normalized_existing = normalize_party_name(existing_party_raw)
            
            # Check for bidirectional inclusion
            if normalized_new in normalized_existing or normalized_existing in normalized_new:
                log_event(f"Matched party '{new_party_name}' to existing file '{file}'")
                return file
                
        return None

    # Check for existing match
    existing_file = find_existing_party_file(output_dir, responding_party)
    
    if existing_file:
        output_filename = existing_file
        # We also want to keep the title consistent within the doc if possible, 
        # but the prompt logic uses 'responding_party' for the subheading.
        # We will stick to the new responding_party name for the subheading 
        # as it might be more accurate for this specific doc, but the file is shared.
    else:
        # Sanitize responding_party for filename
        safe_party_name = re.sub(r'[\\/*?:"<>|]', "", responding_party).replace(" ", "_")
        output_filename = f"Discovery_Responses_{safe_party_name}.docx"

    output_file = os.path.join(output_dir, output_filename)

    save_success = save_to_docx(extraction_content, summary_content, output_file, responding_party, discovery_type)
    
    if save_success:
        if report_dir:
             # Write to report file for batched consolidation later
             import random 
             fname = f"worker_{os.getpid()}_{random.randint(0,1000)}.txt"
             try:
                 with open(os.path.join(report_dir, fname), "w") as f:
                     f.write(output_file)
                 log_event(f"Reported output file to batch: {output_file}")
             except Exception as e:
                 log_event(f"Failed to write report file: {e}", level="error")
        else:
             # Standard behavior: Consolidate immediately
             consolidate_file(output_file)
    
    log_event("--- Agent Finished Successfully ---")

if __name__ == '__main__':
    main()