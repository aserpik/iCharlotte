import os
import sys
import logging
import json
import datetime
import glob
import re
import subprocess
import shutil

# --- Configuration (Kept from original) ---
BASE_PATH_WIN = r"Z:\\Shared\\Current Clients"
GEMINI_DATA_DIR = os.path.join(os.getcwd(), ".gemini", "case_data")
LOG_FILE = os.path.join(os.getcwd(), "Docket_activity.log")
SCRIPTS_DIR = os.path.join(os.getcwd(), "Scripts")
LA_SCRAPER = os.path.join(SCRIPTS_DIR, "docket_scraper.py")
RIVERSIDE_SCRAPER = os.path.join(SCRIPTS_DIR, "riverside_scraper.py")
COMPLAINT_SCRIPT = os.path.join(SCRIPTS_DIR, "complaint.py")

logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def log_event(message, level="info"):
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"[{timestamp}] {message}")
    if level == "info": logging.info(message)
    elif level == "error": logging.error(message)

def get_case_path(file_num):
    carrier_num, case_num = file_num.split('.')
    carrier_folders = glob.glob(os.path.join(BASE_PATH_WIN, f"{carrier_num} - *"))
    if not carrier_folders: return None
    case_folders = glob.glob(os.path.join(carrier_folders[0], f"{case_num} - *"))
    return os.path.abspath(case_folders[0]) if case_folders else None

def get_case_var(file_num, key):
    json_file = os.path.join(GEMINI_DATA_DIR, f"{file_num}.json")
    if not os.path.exists(json_file): return None
    with open(json_file, 'r') as f:
        return json.load(f).get(key)

def main():
    if len(sys.argv) < 2:
        sys.exit(1)

    file_num = sys.argv[1]
    log_event(f"--- Starting Docket Agent for File: {file_num} ---")

    case_number = get_case_var(file_num, "case_number")
    venue = str(get_case_var(file_num, "venue") or "").lower()

    # Determine which scraper to use
    if "riverside" in venue:
        active_scraper = RIVERSIDE_SCRAPER
        log_event(f"Venue: Riverside. Launching Riverside Scraper.")
    else:
        active_scraper = LA_SCRAPER
        log_event(f"Venue: LA. Launching LA Scraper.")

    try:
        # Run the selected scraper script
        result = subprocess.run([sys.executable, active_scraper, case_number], capture_output=True, text=True)
        if result.returncode != 0:
            log_event(f"Scraper Failed: {result.stderr}", level="error")
            sys.exit(1)
    except Exception as e:
        log_event(f"Execution Error: {e}", level="error")
        sys.exit(1)

    # Post-Processing: Move generated PDF
    date_str = datetime.datetime.now().strftime("%Y.%m.%d")
    expected_filename = f"docket_{date_str}.pdf"
    
    # Locate generated file (handles cases where filename varies slightly)
    candidates = glob.glob("docket_*.pdf")
    if not candidates:
        log_event("No PDF found after scraping.", level="error")
        sys.exit(1)
    
    latest_file = max(candidates, key=os.path.getmtime)
    case_path = get_case_path(file_num)
    
    if case_path:
        target_dir = os.path.join(case_path, "NOTES")
        if not os.path.exists(target_dir): os.makedirs(target_dir)
        target_path = os.path.join(target_dir, f"Docket_{date_str}.pdf")
        shutil.move(latest_file, target_path)
        log_event(f"COMPLETED: Saved to {target_path}")

if __name__ == "__main__":
    main()