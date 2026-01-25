import argparse
import os
import sys
import json
import subprocess
import logging
import platform
import shutil
import tempfile
import time
import google.generativeai as genai

# --- Dependency Setup ---
try:
    import pypdf # Prefer pypdf
    PyPDF2 = pypdf
except ImportError:
    try:
        import PyPDF2
    except ImportError:
        print("Error: pypdf or PyPDF2 is required.")
        sys.exit(1)

OCR_AVAILABLE = False
POPPLER_PATH = None

# OCR Path Logic (Adapted from complaint.py)
if os.name == 'nt':
    # Tesseract Path
    try:
        import pytesseract
        from pdf2image import convert_from_path
        
        tesseract_paths = [
            r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
            os.path.expanduser(r"~\\AppData\\Local\\Programs\\Tesseract-OCR\\tesseract.exe")
        ]
        for p in tesseract_paths:
            if os.path.exists(p):
                pytesseract.pytesseract.tesseract_cmd = p
                OCR_AVAILABLE = True
                break
        
        # Poppler Path
        poppler_paths = [
            r"C:\\Program Files\\poppler\\Library\\bin",
            r"C:\\Program Files\\poppler-0.68.0\\bin",
            r"C:\\Program Files (x86)\\poppler\\Library\\bin",
            r"C:\\Program Files\\poppler\\bin" # Added common path
        ]
        for p in poppler_paths:
            if os.path.exists(p):
                POPPLER_PATH = p
                break
    except ImportError:
        pass
else:
    if shutil.which("tesseract") and shutil.which("pdftoppm"): # pdftoppm comes with poppler
        try:
            import pytesseract
            from pdf2image import convert_from_path
            OCR_AVAILABLE = True
        except ImportError:
            pass

# --- Logging Setup ---
# Log to project root
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG_FILE = os.path.join(PROJECT_ROOT, "separator_activity.log")

logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("Separator")

# --- Constants ---
PRIMARY_MODEL = "gemini-3-flash-preview"
FALLBACK_MODEL = "gemini-2.5-flash"

# --- Functions ---

def ocr_page_header(pdf_path, page_num):
    """OCRs the top 50% of a page."""
    if not OCR_AVAILABLE:
        return "[Image Only - OCR Not Available]"
    try:
        # Convert first 50% of page (page_num is 1-based)
        images = convert_from_path(
            pdf_path, 
            first_page=page_num, 
            last_page=page_num, 
            poppler_path=POPPLER_PATH
        )
        if not images: 
            return "[Image Conversion Failed]"
        
        img = images[0]
        width, height = img.size
        # Crop to top half
        img = img.crop((0, 0, width, height // 2))
        
        text = pytesseract.image_to_string(img)
        return text
    except Exception as e:
        logger.error(f"OCR Error on page {page_num}: {e}")
        return f"[OCR Error: {e}]"

def extract_headers(pdf_path):
    """Extracts headers from all pages."""
    headers = []
    try:
        reader = PyPDF2.PdfReader(pdf_path)
    except Exception as e:
        logger.error(f"Failed to read PDF: {e}")
        sys.exit(1)
        
    num_pages = len(reader.pages)
    logger.info(f"Extracting headers from {num_pages} pages in {pdf_path}...")
    
    for i in range(num_pages):
        page_num = i + 1
        text = ""
        try:
            page = reader.pages[i]
            text = page.extract_text() or ""
        except Exception:
            pass
            
        # Logic: If text is sparse (< 50 chars), try OCR
        if len(text.strip()) < 50: 
             ocr_text = ocr_page_header(pdf_path, page_num)
             # Use OCR text if it provided more info
             if len(ocr_text.strip()) > len(text.strip()):
                 text = ocr_text
                 # logger.info(f"Page {page_num} used OCR.")
             
        # Take top ~500 chars clean
        clean_text = " ".join(text[:1000].split())[:500]
        headers.append(f"Page {page_num}: {clean_text}")
        
        if page_num % 20 == 0:
            logger.info(f"Processed {page_num}/{num_pages} pages")
            
    return headers

def call_gemini(prompt, model):
    """Calls Gemini API directly."""
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise Exception("GEMINI_API_KEY environment variable not set.")
        
    genai.configure(api_key=api_key)
    
    try:
        # Run API call
        m = genai.GenerativeModel(model)
        response = m.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        raise Exception(f"Gemini API error: {e}")

def analyze_headers(headers):
    """Sends headers to AI to find boundaries."""
    prompt = (
        "I have a single PDF containing multiple distinct legal or administrative documents. "
        "Below is the text from the top of each page. "
        "Identify the start and end page of each distinct document. "
        "For each document, also extract the document's date (e.g., '2023-05-15') if it is visible on the first page of that document. "
        "Return the result ONLY as a list of lines, one per document, using the pipe character '|' as a separator.\n"
        "Format: ID|Title|Date|StartPage|EndPage\n"
        "Provide a detailed and descriptive title for each document (e.g., 'Plaintiff's Motion to Compel Discovery' instead of just 'Motion', or 'Exhibit A - Medical Report' instead of 'Exhibit').\n"
        "Example:\n"
        "1|Plaintiff's Complaint for Damages|2023-05-15|1|5\n"
        "2|Proof of Service of Summons||6|7\n"
        "3|Exhibit A - Medical Report from Dr. Smith|2022-12-01|8|10\n"
        "\n\nPDF HEADERS:\n" + "\n".join(headers)
    )
    
    response = ""
    try:
        logger.info(f"Analyzing structure with {PRIMARY_MODEL}...")
        response = call_gemini(prompt, PRIMARY_MODEL)
    except Exception as e:
        logger.warning(f"Primary model failed: {e}. Trying fallback {FALLBACK_MODEL}...")
        try:
            response = call_gemini(prompt, FALLBACK_MODEL)
        except Exception as e2:
            logger.error(f"Fallback model failed: {e2}")
            sys.exit(1)
            
    # Parse Response
    docs = []
    lines = response.split('\n')
    for line in lines:
        line = line.strip()
        if not line or '|' not in line: continue
        if line.startswith("```"): continue # Skip markdown
        
        parts = line.split('|')
        if len(parts) >= 5:
            try:
                docs.append({
                    "id": parts[0].strip(),
                    "title": parts[1].strip(),
                    "date": parts[2].strip(),
                    "start": int(parts[3].strip()),
                    "end": int(parts[4].strip())
                })
            except ValueError:
                pass
    
    if not docs:
        logger.error("No documents identified by AI.")
        logger.debug(f"Raw AI response: {response}")
        sys.exit(1)
        
    return docs

def sanitize_filename(name):
    return "".join(c for c in name if c.isalnum() or c in (' ', '-', '_')).strip()

def split_pdf_pages(pdf_path, output_folder, selection):
    """Splits the PDF based on selection."""
    try:
        reader = PyPDF2.PdfReader(pdf_path)
    except Exception as e:
        print(f"Error reading source PDF: {e}")
        return

    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    
    print(f"\nProcessing {len(selection)} items...")
    
    for doc in selection:
        try:
            writer = PyPDF2.PdfWriter()
            start = doc['start'] - 1 # 1-based to 0-based
            end = doc['end']     # exclusive in python slice, inclusive in prompt? Prompt said "1-5", usually means 1,2,3,4,5. 
                                 # Range(start, end) where end is 5 would be 0,1,2,3,4. 
                                 # If page 5 is included, we need index 4. Range(0, 5) covers 0,1,2,3,4.
                                 # So if doc says end=5, we want loop range(0, 5). 
            
            # Verify bounds
            if start < 0: start = 0
            if end > len(reader.pages): end = len(reader.pages)
            
            for i in range(start, end):
                writer.add_page(reader.pages[i])
            
            safe_title = sanitize_filename(doc['title'])
            if len(safe_title) > 100:
                safe_title = safe_title[:100].strip()
            # Format: MyPacket- 01 - Complaint.pdf
            # User example: "MyPacket- 01 - Complaint.pdf" (original + id + title)
            # Assuming ID is numeric.
            try:
                 id_num = int(doc['id'])
                 id_str = f"{id_num:02d}"
            except:
                 id_str = str(doc['id'])
                 
            out_filename = f"{base_name} - {id_str} - {safe_title}.pdf"
            out_path = os.path.join(output_folder, out_filename)
            
            with open(out_path, "wb") as f:
                writer.write(f)
                
            print(f"  [OK] Saved: {out_filename}")
            logger.info(f"Saved {out_filename}")
            
        except Exception as e:
            print(f"  [ERROR] Failed to save {doc['title']}: {e}")
            logger.error(f"Failed to save {doc['title']}: {e}")

# --- Modes ---

def run_analysis(pdf_path):
    if not os.path.exists(pdf_path):
        logger.error(f"File not found: {pdf_path}")
        return

    logger.info(f"Starting analysis for: {pdf_path}")
    
    # 1. Extract
    headers = extract_headers(pdf_path)
    
    # 2. Analyze
    docs = analyze_headers(headers)
    logger.info(f"identified {len(docs)} documents.")
    
    # 3. Save Temp Map
    fd, temp_path = tempfile.mkstemp(suffix=".json", text=True)
    with os.fdopen(fd, 'w') as f:
        json.dump(docs, f)
    
    logger.info(f"Map saved to {temp_path}. Launching interactive window...")
    
    # 4. Launch Interactive Window
    # We recursively call this script with different arguments in a new console
    current_script = os.path.abspath(__file__)
    
    cmd = [sys.executable, current_script, "--interactive", temp_path, "--original-pdf", pdf_path]
    
    if os.name == 'nt':
        subprocess.Popen(cmd, creationflags=subprocess.CREATE_NEW_CONSOLE)
    else:
        # Linux fallback (not primary target per prompt, but good practice)
        subprocess.Popen(cmd)

def run_interactive(json_path, pdf_path):
    # Load Docs
    docs = []
    try:
        with open(json_path, 'r') as f:
            docs = json.load(f)
    except Exception as e:
        print(f"Error loading document map: {e}")
        input("Press Enter to exit...")
        return

    # Display Loop
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print(f"--- Document Separator: {os.path.basename(pdf_path)} ---\n")
        print(f"{ 'ID':<4} | { 'Date':<12} | { 'Pages':<8} | {'Title'}")
        print("-" * 60)
        for doc in docs:
            print(f"{doc['id']:<4} | {doc['date']:<12} | {doc['start']}-{doc['end']:<6} | {doc['title']}")
        print("-" * 60)
        
        print("\nEnter IDs to keep (e.g., '1, 3', '1-5', 'all') or 'q' to quit.")
        choice = input("> ").strip().lower()
        
        if choice == 'q':
            break
            
        selection = []
        if choice == 'all':
            selection = docs
        else:
            try:
                # Parse range/list
                ids_to_keep = set()
                parts = choice.split(',')
                for part in parts:
                    part = part.strip()
                    if '-' in part:
                        s, e = map(int, part.split('-'))
                        ids_to_keep.update(range(s, e + 1))
                    else:
                        ids_to_keep.add(str(part)) # Try string first
                        ids_to_keep.add(str(int(part))) # Also numeric string
                
                selection = [d for d in docs if str(d['id']) in ids_to_keep]
            except Exception as e:
                print(f"Invalid input: {e}")
                time.sleep(1)
                continue
        
        if not selection:
            print("No documents selected.")
            time.sleep(1)
            continue
            
        # Process Selection
        base_dir = os.path.dirname(os.path.abspath(pdf_path))
        source_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_folder = os.path.join(base_dir, f"PULLED-{source_name}")
        
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            
        split_pdf_pages(pdf_path, output_folder, selection)
        
        print("\nDone! Files saved to:")
        print(output_folder)
        print("\nPress Enter to close this window.")
        input()
        
        # Cleanup temp file
        try:
            os.remove(json_path)
        except:
            pass
        break

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf_path", help="PDF File Path", nargs='?')
    parser.add_argument("--interactive", help="Path to temp json for interactive mode")
    parser.add_argument("--original-pdf", help="Original PDF for interactive mode")
    args = parser.parse_args()

    if args.interactive and args.original_pdf:
        run_interactive(args.interactive, args.original_pdf)
    elif args.pdf_path:
        run_analysis(args.pdf_path)
    else:
        # Fallback if no args provided (shouldn't happen with correct usage)
        print("Usage: separate.py <pdf_path>")

if __name__ == "__main__":
    main()
