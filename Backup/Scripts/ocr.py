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
DPI = 600
JPEG_QUALITY = 30  # Aggressive compression for mostly-text documents

# --- Logging Setup ---
log_queue = None

def log_listener(queue, log_file_path):
    with open(log_file_path, "a") as f:
        while True:
            try:
                message = queue.get()
                if message == "STOP":
                    break
                f.write(message + "\n")
                f.flush()
            except Exception:
                break

def agent_log(queue, message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        queue.put(f"[{timestamp}] {message}")
    except Exception:
        pass

# --- Worker Function ---
def process_page(args):
    page_num, pdf_path, target_dpi, jpeg_qual, q = args
    
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
             agent_log(q, f"OCR INFO: Normalizing massive metadata ({width_pts:.0f}x{height_pts:.0f} pts) to standard page size.")
             scale = min(target_w / width_pts, target_h / height_pts)
             mat = fitz.Matrix(scale, scale)
        else:
             mat = fitz.Matrix(1, 1)
        
        current_dpi = target_dpi
        pix = None
        
        while current_dpi >= 150:
            try:
                agent_log(q, f"OCR TRACE: Rendering Page {page_num + 1} at {current_dpi} DPI...")
                pix = page.get_pixmap(matrix=mat.prescale(current_dpi/72, current_dpi/72))
                break 
            except Exception as e:
                agent_log(q, f"OCR WARNING: Page {page_num + 1} render failed: {e}. Retrying...")
                current_dpi = int(current_dpi / 1.5)
        
        if not pix:
             agent_log(q, f"OCR ERROR: Could not render Page {page_num + 1}.")
             doc.close()
             return (page_num, None)
        
        # Convert to PIL Image
        mode = "RGB" if pix.alpha == 0 else "RGBA"
        img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        if mode == "RGBA":
            img = img.convert("RGB")
            
        # --- Background Leveling Optimization ---
        # Cleans off-white noise to pure white, improving JPEG compression drastically
        # while keeping stamps and letterheads intact.
        enhancer_c = ImageEnhance.Contrast(img)
        img = enhancer_c.enhance(1.15)
        enhancer_b = ImageEnhance.Brightness(img)
        img = enhancer_b.enhance(1.05)
            
        img_buffer = io.BytesIO()
        img.save(img_buffer, format="JPEG", quality=jpeg_qual, optimize=True)
        img_buffer.seek(0)
        
        compressed_img = Image.open(img_buffer)
        
        agent_log(q, f"OCR TRACE: Running Tesseract on Page {page_num + 1}...")
        pdf_bytes = pytesseract.image_to_pdf_or_hocr(compressed_img, extension='pdf')
        
        compressed_img.close()
        img.close()
        doc.close()
        
        return (page_num, pdf_bytes)
        
    except Exception as e:
        agent_log(q, f"OCR ERROR: Page {page_num + 1} failed: {e}")
        return (page_num, None)

# --- Main Processing Logic ---
def process_pdf(file_path, queue):
    agent_log(queue, f"OCR INFO: Starting processing for {os.path.basename(file_path)}")
    try:
        doc = fitz.open(file_path)
        total_pages = len(doc)
        doc.close()
    except Exception as e:
        agent_log(queue, f"OCR ERROR: Could not open file: {e}")
        return

    agent_log(queue, f"OCR INFO: Found {total_pages} pages.")

    # Force sequential processing to save memory
    cpu_count = 1 
    results = {}
    
    agent_log(queue, f"OCR INFO: Processing {total_pages} pages sequentially for memory safety...")
    
    for i in range(total_pages):
        task = (i, file_path, DPI, JPEG_QUALITY, queue)
        p_num, p_bytes = process_page(task)
        if p_bytes:
            results[p_num] = p_bytes
            agent_log(queue, f"OCR TRACE: Page {p_num + 1} completed.")
        else:
            agent_log(queue, f"OCR ERROR: Page {p_num + 1} returned no data.")

    # Reassemble PDF
    agent_log(queue, "OCR INFO: Merging pages into final document...")
    final_doc = fitz.open()
    
    success_count = 0
    for i in range(total_pages):
        if i in results:
            try:
                page_doc = fitz.open("pdf", results[i])
                final_doc.insert_pdf(page_doc)
                success_count += 1
            except Exception as e:
                agent_log(queue, f"OCR ERROR: Failed to merge page {i+1}: {e}")

    if success_count > 0:
        temp_output = file_path + ".temp"
        # Using garbage=4 and deflate=True for maximum compression
        final_doc.save(temp_output, garbage=4, deflate=True)
        final_doc.close()
        
        try:
            os.replace(temp_output, file_path)
            agent_log(queue, "OCR SUCCESS: File processed and saved with optimization.")
        except OSError as e:
             agent_log(queue, f"OCR ERROR: Could not replace original file: {e}")
    else:
        agent_log(queue, "OCR ERROR: No pages were processed successfully.")
        final_doc.close()

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

    log_file = os.path.join(os.getcwd(), "ocr_activity.log")
    manager = multiprocessing.Manager()
    queue = manager.Queue()
    
    listener = threading.Thread(target=log_listener, args=(queue, log_file))
    listener.start()

    if len(sys.argv) < 2:
        print("Usage: python ocr.py <file_or_dir_path>")
        queue.put("STOP")
        listener.join()
        sys.exit(1)

    target_path = " ".join(sys.argv[1:]).strip("'").strip("'").strip()
    
    if not os.path.exists(target_path) and os.path.exists(sys.argv[1].strip("'").strip("'").strip()):
        target_path = sys.argv[1].strip("'").strip("'").strip()

    agent_log(queue, f"OCR AGENT: Started processing target: {target_path}")

    files_to_process = []
    if os.path.isfile(target_path):
        files_to_process = [target_path]
    elif os.path.isdir(target_path):
        import glob
        files_to_process = glob.glob(os.path.join(target_path, "**", "*.pdf"), recursive=True)

    for fp in files_to_process:
        try:
            process_pdf(fp, queue)
        except Exception as e:
             agent_log(queue, f"OCR CRITICAL ERROR: {e}")

    agent_log(queue, "OCR AGENT: All tasks completed.")
    queue.put("STOP")
    listener.join()

if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()