import os
import sys
import logging
import datetime
import re
import subprocess
from google import genai
from docx import Document
from docx.shared import Pt, Inches
import concurrent.futures
import io

# --- Dependency Setup ---
try:
    import fitz # PyMuPDF
except ImportError:
    print("Error: 'pymupdf' is required for this optimized version.")
    print("Please run: pip install pymupdf")
    sys.exit(1)

try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        pass # Not strictly needed if fitz works, but good for fallback logic if kept

OCR_AVAILABLE = False
# ... (rest of imports remain the same) ...
# ... (logging setup remains the same) ...

def log_event(message, level="info"):
    print(message)
    sys.stdout.flush()
    if level == "info":
        logging.info(message)
    elif level == "error":
        logging.error(message)
    elif level == "warning":
        logging.warning(message)

# ... (text extraction functions remain same) ...

def get_page_text_fast(pdf_path, page_num):
    """
    Extracts text from a single page using PyMuPDF (fitz).
    If text is sparse, it renders the page image and uses OCR (Tesseract).
    page_num is 1-based.
    """
    try:
        # Open PDF (thread-safe if opened locally)
        doc = fitz.open(pdf_path)
        page = doc.load_page(page_num - 1)
        
        # 1. Try direct text extraction (VERY FAST)
        text = page.get_text("text")
        
        # 2. If sparse, fall back to OCR
        if len(text.strip()) < 50 and OCR_AVAILABLE:
            # Render page to image (pixmap)
            # matrix=2.0 (144 dpi) -> 300 dpi approx
            mat = fitz.Matrix(2, 2) 
            pix = page.get_pixmap(matrix=mat)
            
            # Convert to PIL Image
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            
            # Run OCR
            ocr_text = pytesseract.image_to_string(img)
            
            if len(ocr_text.strip()) > len(text.strip()):
                text = ocr_text

        doc.close()
        return (page_num, text)
        
    except Exception as e:
        log_event(f"Error processing page {page_num}: {e}", level="warning")
        return (page_num, "")

def extract_text_parallel(file_path):
    """Extracts text from PDF using parallel processing."""
    log_event(f"Extracting text from: {file_path} using parallel processing...")
    
    try:
        doc = fitz.open(file_path)
        num_pages = len(doc)
        doc.close()
    except Exception as e:
        log_event(f"Error reading PDF structure: {e}", level="error")
        return None

    log_event(f"Total pages: {num_pages}")
    
    # Use ThreadPoolExecutor
    max_workers = os.cpu_count() or 4
    if max_workers > 16: max_workers = 16
    
    page_texts = {}
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_page = {
            executor.submit(get_page_text_fast, file_path, i + 1): i + 1 
            for i in range(num_pages)
        }
        
        processed_count = 0
        for future in concurrent.futures.as_completed(future_to_page):
            try:
                p_num, text = future.result()
                page_texts[p_num] = text
            except Exception as exc:
                log_event(f"Page generated an exception: {exc}", level="error")
            
            processed_count += 1
            if processed_count % 50 == 0 or processed_count == num_pages:
                log_event(f"Processed {processed_count}/{num_pages} pages")

    # Reassemble in order
    full_text = ""
    for i in range(1, num_pages + 1):
        if i in page_texts:
            full_text += page_texts[i] + "\n"
            
    return full_text

def extract_text(file_path):
    """Wrapper to handle different file types."""
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext == ".pdf":
        return extract_text_parallel(file_path)
        
    elif ext == ".docx":
        try:
            doc = Document(file_path)
            text = ""
            for paragraph in doc.paragraphs:
                text += paragraph.text + "\n"
            return text
        except Exception as e:
            log_event(f"Error reading DOCX: {e}", level="error")
            return None
            
    else:
        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
        except Exception as e:
            log_event(f"Error reading text file: {e}", level="error")
            return None

def chunk_text(text, chunk_size=500000):
    """Splits text into chunks of approximately chunk_size characters."""
    chunks = []
    for i in range(0, len(text), chunk_size):
        chunks.append(text[i:i + chunk_size])
    return chunks

def clean_response(text):
    """
    Extracts content between <rewritten_chronology> tags.
    If tags are missing but a scratchpad is present, it attempts to take everything after the scratchpad.
    """
    # 1. Try to find content within <rewritten_chronology> tags
    chron_pattern = re.compile(r'<rewritten_chronology>(.*?)(?:</rewritten_chronology>|$)', re.DOTALL | re.IGNORECASE)
    chron_match = chron_pattern.search(text)
    if chron_match:
        content = chron_match.group(1).strip()
        if content:
            return content

    # 2. Fallback: If <scratchpad> is present, take everything after </scratchpad>
    if "<scratchpad>" in text.lower() or "</scratchpad>" in text.lower():
        scratch_end_pattern = re.compile(r'</scratchpad>(.*)', re.DOTALL | re.IGNORECASE)
        scratch_match = scratch_end_pattern.search(text)
        if scratch_match:
            return scratch_match.group(1).strip()
            
    # 3. Last resort: Return the text but strip common AI conversational markers if they exist
    # (Though usually the above two cover 99% of cases with the current prompt)
    return text.strip()

def call_gemini(prompt, text):
    """Calls Gemini API, handling large text via chunking."""
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        log_event("Error: GEMINI_API_KEY environment variable not set.", level="error")
        return None

    client = genai.Client(api_key=api_key)

    total_chars = len(text)
    log_event(f"Total document length: {total_chars} characters.")

    # Chunk if larger than ~800k characters (safety margin for 1M token limit models)
    chunks = chunk_text(text, chunk_size=500000)
    full_response = ""
    
    model_sequence = [
        "gemini-3-flash-preview", 
        "gemini-2.5-flash"
    ]

    for i, chunk in enumerate(chunks):
        log_event(f"Processing chunk {i+1}/{len(chunks)} ({len(chunk)} chars)...")
        
        # Add context for chunks
        chunk_prompt = prompt
        if len(chunks) > 1:
            chunk_prompt += f"\n\n(Note: This is part {i+1} of {len(chunks)} of the document. Summarize/extract relevant information from this section.)"

        full_prompt = f"{chunk_prompt}\n\nDOCUMENT CONTENT:\n{chunk}"
        
        chunk_success = False
        for model_name in model_sequence:
            log_event(f"Attempting to use model: {model_name}")
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=full_prompt
                )
                if response and response.text:
                    log_event(f"Success with model: {model_name}")
                    cleaned = clean_response(response.text)
                    full_response += cleaned + "\n\n"
                    chunk_success = True
                    break # Move to next chunk
            except Exception as e:
                log_event(f"Failed with model {model_name}: {e}", level="warning")
                if "429" in str(e):
                    import time
                    log_event("Rate limit hit. Waiting 30 seconds...")
                    time.sleep(30)
                continue
        
        if not chunk_success:
            log_event(f"Error: Failed to process chunk {i+1}. Skipping.", level="error")
            full_response += f"\n[Error: Could not process part {i+1} of the document]\n"

    return full_response.strip()

def add_markdown_to_doc(doc, content):
    """Parses basic Markdown and applies formatting."""
    lines = content.split('\n')
    active_paragraph = None
    
    # Regex to identify dates at the beginning of a paragraph/sentence
    # Matches: (Optional "On ") (Date String)
    # Date String: Month DD, YYYY
    date_pattern = re.compile(r"^(On )?((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]* \d{1,2},? \d{4})")

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        
        if stripped.startswith('#'):
            text = stripped.lstrip('#').strip()
            if not text.endswith('.'):
                text += "."
            active_paragraph = doc.add_paragraph()
            run = active_paragraph.add_run(text + " ")
            run.bold = True
            continue
        
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
            active_paragraph = None
            continue
        
        # Regular paragraph
        if active_paragraph:
            p = active_paragraph
        else:
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Inches(0.5)
        
        # Check for date at start of line
        match = date_pattern.match(stripped)
        if match:
            on_prefix = match.group(1) # "On " or None
            date_str = match.group(2)  # "January 1, 2024"
            
            total_match_len = len(match.group(0))
            remaining_text = stripped[total_match_len:]
            
            if on_prefix:
                p.add_run(on_prefix)
            
            run = p.add_run(date_str)
            run.underline = True
            
            # Process remaining text for bold markers
            parts = re.split(r'(\**.*?\**)', remaining_text)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    run = p.add_run(part[2:-2])
                    run.bold = True
                else:
                    p.add_run(part)
        else:
            # Standard markdown parsing
            parts = re.split(r'(\**.*?\**)', stripped)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    run = p.add_run(part[2:-2])
                    run.bold = True
                else:
                    p.add_run(part)
        
        active_paragraph = None

def extract_provider_from_filename(filename):
    """Extracts provider name from filename based on patterns."""
    # Pattern 1: 12345-001_ PROVIDER NAME (1).pdf
    # Split by underscore
    parts = filename.split('_')
    if len(parts) > 1:
        # Take the part after the first underscore
        potential_name = parts[1]
        # Remove extension
        potential_name = os.path.splitext(potential_name)[0]
        # Remove trailing parentheses like (1)
        potential_name = re.sub(r'\s*\(\d+\)$', '', potential_name)
        return potential_name.strip()
    
    # Fallback: Use filename without extension and sanitize
    name = os.path.splitext(filename)[0]
    return name.strip()

def sanitize_filename(name):
    """Sanitizes a string to be safe for filenames."""
    # Remove invalid characters
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    # Trim whitespace
    name = name.strip()
    return name

def save_to_docx(content, output_dir, provider_name, original_filename):
    """Saves to DOCX with dynamic filename."""
    
    safe_provider = sanitize_filename(provider_name)
    if not safe_provider:
        safe_provider = "Unknown_Provider"
    
    # Filename format: Med_Record_[name of treatment provider].docx
    filename = f"Med_Record_{safe_provider}.docx"
    output_path = os.path.join(output_dir, filename)

    try:
        if os.path.exists(output_path):
             # If exists, we append, similar to summarize.py
             doc = Document(output_path)
             doc.add_page_break()
             log_event(f"Appending to existing file: {output_path}")
        else:
             doc = Document()
             log_event(f"Creating new file: {output_path}")

        # Styles
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(12)
        style.paragraph_format.line_spacing = 1.0
        style.paragraph_format.space_after = Pt(12)
        
        for i in range(1, 10):
            if f'Heading {i}' in doc.styles:
                h = doc.styles[f'Heading {i}']
                h.font.name = 'Times New Roman'
                h.font.size = Pt(12)
                h.paragraph_format.line_spacing = 1.0
                h.paragraph_format.space_after = Pt(12)

        for s in ['List Bullet', 'List Number']:
            if s in doc.styles:
                l = doc.styles[s]
                l.font.name = 'Times New Roman'
                l.font.size = Pt(12)
                l.paragraph_format.line_spacing = 1.0
                l.paragraph_format.space_after = Pt(12)

        # Title
        p = doc.add_paragraph()
        run = p.add_run(f"Medical Record Chronology - {provider_name}")
        run.bold = True
        run.underline = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)

        doc.add_paragraph(f"Source: {original_filename}")
        doc.add_paragraph(f"Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        add_markdown_to_doc(doc, content)
            
        doc.save(output_path)
        log_event(f"Saved to: {output_path}")
        return True

    except Exception as e:
        log_event(f"Error saving DOCX: {e}", level="error")
        return False

def main():
    if len(sys.argv) < 2:
        log_event("Error: No file path provided.", level="error")
        sys.exit(1)

    # Use sys.argv[1:] directly to allow space-separated file paths.
    # Users must quote paths containing spaces.
    raw_args = sys.argv[1:]
    
    if not raw_args:
        log_event("No valid file paths found.", level="error")
        sys.exit(1)

    # Heuristic: Check if arguments form a single split path (common with unquoted paths having spaces)
    combined_arg = " ".join(raw_args)
    # Aggressively strip quotes that might have caused the split or be lingering
    clean_combined = combined_arg.strip().strip('"').strip("'")
    
    log_event(f"Checking heuristic path: {clean_combined}")
    
    if os.path.exists(clean_combined):
        file_paths = [clean_combined]
        log_event(f"Reconstructed split path: {clean_combined}")
    else:
        file_paths = raw_args
    
    # --- Dispatcher Mode ---
    if len(file_paths) > 1:
        log_event(f"Detected multiple file paths: {len(file_paths)} items. Launching separate agents...")
        for path in file_paths:
            # We must quote paths with spaces when re-launching
            quoted_path = f'"{path}"' if ' ' in path and not path.startswith('"') else path
            log_event(f"Spawning agent for: {path}")
            try:
                # Spawn independent process
                if os.name == 'nt':
                    # Use CREATE_NO_WINDOW (0x08000000) for headless execution
                    subprocess.Popen([sys.executable, sys.argv[0], path], creationflags=0x08000000)
                else:
                    subprocess.Popen([sys.executable, sys.argv[0], path])
            except Exception as e:
                log_event(f"Failed to spawn agent for {path}: {e}", level="error")
        
        print(f"Launched {len(file_paths)} medical record agents.")
        sys.exit(0)

    # --- Worker Mode ---
    input_path = file_paths[0]
    
    # Remove surrounding quotes if present
    input_path = input_path.strip().strip('"').strip("'")
    
    # Handle absolute path conversion
    input_path = os.path.abspath(input_path)

    if not os.path.exists(input_path):
        log_event(f"Error: File not found: {input_path}", level="error")
        sys.exit(1)

    log_event(f"--- Starting Med Record Agent for: {input_path} ---")

    if os.path.isdir(input_path):
        log_event(f"Input is a directory. Scanning: {input_path}")
        files_to_process = []
        for root, _, files in os.walk(input_path):
            for file in files:
                if file.lower().endswith(('.pdf', '.docx')):
                    if "Med_Record_" in file: continue
                    files_to_process.append(os.path.join(root, file))
        
        if not files_to_process:
            log_event("No suitable files found.")
            sys.exit(0)
        
        for file_path in files_to_process:
            try:
                subprocess.run([sys.executable, sys.argv[0], file_path], check=True)
            except subprocess.CalledProcessError as e:
                log_event(f"Subprocess failed for {file_path}: {e}", level="error")
        sys.exit(0)
    
    # Single File
    text = extract_text(input_path)
    if not text:
        sys.exit(1)

    try:
        with open(PROMPT_FILE, "r", encoding="utf-8") as f:
            prompt_instruction = f.read()
    except Exception as e:
        log_event(f"Error reading prompt file: {e}", level="error")
        sys.exit(1)

    # Determine Provider Name from Filename
    filename = os.path.basename(input_path)
    provider_name = extract_provider_from_filename(filename)
    log_event(f"Identified Provider from Filename: {provider_name}")

    final_content = call_gemini(prompt_instruction, text)
    if not final_content:
        sys.exit(1)

    # Determine Output Directory
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
        for i in range(len(parts) - 1, -1, -1):
            if parts[i].upper() == "NOTES":
                output_dir = os.path.join(os.sep.join(parts[:i+1]), "AI OUTPUT")
                break
    
    if not output_dir:
        input_dir = os.path.dirname(input_path)
        parent_dir = os.path.dirname(input_dir)
        output_dir = os.path.join(parent_dir, "NOTES", "AI OUTPUT")

    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
        except Exception:
            pass

    save_to_docx(final_content, output_dir, provider_name, filename)
    log_event("--- Agent Finished ---")

if __name__ == '__main__':
    main()