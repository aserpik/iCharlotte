import sys
import os
import fitz  # PyMuPDF
import pytesseract
from PIL import Image, ImageEnhance
Image.MAX_IMAGE_PIXELS = None
import io
from datetime import datetime
import threading
import multiprocessing
import queue

# --- Configuration ---
DPI = 300
JPEG_QUALITY = 30  # Aggressive compression for mostly-text documents

# --- Logging Setup ---
def agent_log(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_file_path = os.path.join(os.getcwd(), "ocr_activity.log")
    try:
        with open(log_file_path, "a") as f:
            f.write(f"[{timestamp}] {message}\n")
    except Exception:
        pass

# --- Worker Function ---
def process_page(args):
    page_num, pdf_path, target_dpi, jpeg_qual = args
    
    try:
        doc = fitz.open(pdf_path)
        page = doc.load_page(page_num)
        
        # --- Normalization Step ---
        rect = page.rect
        width_pts = rect.width
        height_pts = rect.height
        target_w = 612
        target_h = 792
        
        if width_pts > 1440 or height_pts > 1440:
             agent_log(f"OCR INFO: Normalizing massive metadata ({width_pts:.0f}x{height_pts:.0f} pts) to standard page size.")
             scale = min(target_w / width_pts, target_h / height_pts)
             mat = fitz.Matrix(scale, scale)
        else:
             mat = fitz.Matrix(1, 1)
        
        current_dpi = target_dpi
        pix = None
        
        while current_dpi >= 150:
            try:
                agent_log(f"OCR TRACE: Rendering Page {page_num + 1} at {current_dpi} DPI...")
                pix = page.get_pixmap(matrix=mat.prescale(current_dpi/72, current_dpi/72))
                break 
            except Exception as e:
                agent_log(f"OCR WARNING: Page {page_num + 1} render failed: {e}. Retrying...")
                current_dpi = int(current_dpi / 1.5)
        
        if not pix:
             agent_log(f"OCR ERROR: Could not render Page {page_num + 1}.")
             doc.close()
             return (page_num, None)
        
        # Convert to PIL Image
        mode = "RGB" if pix.alpha == 0 else "RGBA"
        img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        if mode == "RGBA":
            img = img.convert("RGB")
            
        # --- Background Leveling Optimization ---
        enhancer_c = ImageEnhance.Contrast(img)
        img = enhancer_c.enhance(1.15)
        enhancer_b = ImageEnhance.Brightness(img)
        img = enhancer_b.enhance(1.05)
            
        img_buffer = io.BytesIO()
        img.save(img_buffer, format="JPEG", quality=jpeg_qual, optimize=True)
        img_buffer.seek(0)
        
        compressed_img = Image.open(img_buffer)
        
        agent_log(f"OCR TRACE: Running Tesseract on Page {page_num + 1}...")
        pdf_bytes = pytesseract.image_to_pdf_or_hocr(compressed_img, extension='pdf')
        
        compressed_img.close()
        img.close()
        doc.close()
        
        return (page_num, pdf_bytes)
        
    except Exception as e:
        agent_log(f"OCR ERROR: Page {page_num + 1} failed: {e}")
        return (page_num, None)

# --- Main Processing Logic ---
def process_pdf(file_path):
    agent_log(f"OCR INFO: Starting processing for {os.path.basename(file_path)}")
    try:
        doc = fitz.open(file_path)
        total_pages = len(doc)
        doc.close()
    except Exception as e:
        agent_log(f"OCR ERROR: Could not open file: {e}")
        return

    agent_log(f"OCR INFO: Found {total_pages} pages.")

    results = {}
    agent_log(f"OCR INFO: Processing {total_pages} pages sequentially for memory safety...")
    
    for i in range(total_pages):
        task = (i, file_path, DPI, JPEG_QUALITY)
        p_num, p_bytes = process_page(task)
        if p_bytes:
            results[p_num] = p_bytes
            agent_log(f"OCR TRACE: Page {p_num + 1} completed.")
        else:
            agent_log(f"OCR ERROR: Page {p_num + 1} returned no data.")
        
        # Report progress for UI
        percentage = int(((i + 1) / total_pages) * 100)
        print(f"PROGRESS: {percentage}")
        sys.stdout.flush()

    # Reassemble PDF
    agent_log("OCR INFO: Merging pages into final document...")
    final_doc = fitz.open()
    
    success_count = 0
    for i in range(total_pages):
        if i in results:
            try:
                page_doc = fitz.open("pdf", results[i])
                final_doc.insert_pdf(page_doc)
                success_count += 1
            except Exception as e:
                agent_log(f"OCR ERROR: Failed to merge page {i+1}: {e}")

    if success_count > 0:
        temp_output = file_path + ".temp"
        final_doc.save(temp_output, garbage=4, deflate=True)
        final_doc.close()
        
        import time
        max_retries = 5
        success_save = False
        final_path = file_path
        
        for attempt in range(max_retries):
            try:
                os.replace(temp_output, file_path)
                agent_log("OCR SUCCESS: File processed and saved with optimization.")
                success_save = True
                break
            except OSError as e:
                if attempt < max_retries - 1:
                    agent_log(f"OCR WARNING: Could not replace original file (attempt {attempt+1}). Retrying in 1s... Error: {e}")
                    time.sleep(1)
                else:
                    agent_log(f"OCR ERROR: Final attempt to replace original file failed: {e}")
        
        if not success_save:
            # Fallback: Save to a new file
            base, ext = os.path.splitext(file_path)
            fallback_path = f"{base}_OCR{ext}"
            try:
                os.replace(temp_output, fallback_path)
                agent_log(f"OCR SUCCESS: Saved to fallback path: {fallback_path}")
                final_path = fallback_path
                success_save = True
            except Exception as fe:
                agent_log(f"OCR CRITICAL ERROR: Could not save to fallback path: {fe}")
                sys.exit(1)
        
        # Print final path to stdout so caller can capture it
        print(f"FINAL_PATH: {final_path}")
    else:
        agent_log("OCR ERROR: No pages were processed successfully.")
        final_doc.close()
        sys.exit(1)

def main():
    if os.name == 'nt':
        tesseract_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe")
        ]
        for p in tesseract_paths:
            if os.path.exists(p):
                pytesseract.pytesseract.tesseract_cmd = p
                break

    if len(sys.argv) < 2:
        print("Usage: python ocr.py <file_or_dir_path>")
        sys.exit(1)

    target_path = " ".join(sys.argv[1:]).strip("'").strip("'").strip()
    
    if not os.path.exists(target_path) and os.path.exists(sys.argv[1].strip("'").strip("'").strip()):
        target_path = sys.argv[1].strip("'").strip("'").strip()

    agent_log(f"OCR AGENT: Started processing target: {target_path}")

    files_to_process = []
    if os.path.isfile(target_path):
        files_to_process = [target_path]
    elif os.path.isdir(target_path):
        import glob
        files_to_process = glob.glob(os.path.join(target_path, "**", "*.pdf"), recursive=True)

    for fp in files_to_process:
        try:
            process_pdf(fp)
        except Exception as e:
             agent_log(f"OCR CRITICAL ERROR: {e}")

    agent_log("OCR AGENT: All tasks completed.")

if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()