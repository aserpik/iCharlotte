import os
import sys
import logging
import json
import datetime
import glob
import re
import subprocess
import shutil
import time

# --- Configuration ---
# Directories
BASE_PATH_WIN = r"Z:\\Shared\\Current Clients"
GEMINI_DATA_DIR = os.path.join(os.getcwd(), ".gemini", "case_data")
LOG_FILE = os.path.join(os.getcwd(), "Docket_activity.log")
SCRIPTS_DIR = os.path.join(os.getcwd(), "Scripts")
SCRAPER_SCRIPT = os.path.join(SCRIPTS_DIR, "docket_scraper.py")
COMPLAINT_SCRIPT = os.path.join(SCRIPTS_DIR, "complaint.py")

# Setup Logging
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

def log_event(message, level="info"):
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"[{timestamp}] {message}")
    sys.stdout.flush()
    if level == "info":
        logging.info(message)
    elif level == "error":
        logging.error(message)
    elif level == "warning":
        logging.warning(message)

def get_case_path(file_num: str):
    """Locates the physical directory for a file number (####.###)."""
    if not re.match(r'^\d{4}\.\d{3}$', file_num):
        log_event(f"Invalid file number format: {file_num}. Expected ####.###", level="error")
        return None

    carrier_num, case_num = file_num.split('.')
    base_path = BASE_PATH_WIN

    if not os.path.exists(base_path):
        log_event(f"Base path not found: {base_path}", level="error")
        return None

    # Find Carrier Folder
    carrier_folders = glob.glob(os.path.join(base_path, f"{carrier_num} - *"))
    if not carrier_folders:
        log_event(f"Carrier folder starting with {carrier_num} not found in {base_path}", level="error")
        return None
    carrier_path = carrier_folders[0]

    # Find Case Folder
    case_folders = glob.glob(os.path.join(carrier_path, f"{case_num} - *"))
    if not case_folders:
        log_event(f"Case folder starting with {case_num} not found in {carrier_path}", level="error")
        return None
    
    return os.path.abspath(case_folders[0])

def get_case_var(file_num: str, key: str):
    """Retrieves a variable from the case's JSON state file."""
    json_file = os.path.join(GEMINI_DATA_DIR, f"{file_num}.json")
    if not os.path.exists(json_file):
        return None
    
    try:
        with open(json_file, 'r') as f:
            data = json.load(f)
            return data.get(key)
    except Exception as e:
        log_event(f"Error reading JSON for {file_num}: {e}", level="error")
        return None

def run_complaint_agent(file_num: str):
    """Runs the complaint agent script."""
    log_event(f"Triggering Complaint Agent for {file_num}...")
    try:
        subprocess.run([sys.executable, COMPLAINT_SCRIPT, file_num], check=True)
        log_event("Complaint Agent finished.")
    except subprocess.CalledProcessError as e:
        log_event(f"Complaint Agent failed with exit code {e.returncode}.", level="error")
    except Exception as e:
        log_event(f"Failed to run Complaint Agent: {e}", level="error")

def main():
    if len(sys.argv) < 2:
        log_event("Usage: python docket.py <file_number>", level="error")
        sys.exit(1)

    file_num = sys.argv[1]
    log_event(f"--- Starting Docket Agent for File: {file_num} ---")

    # --- Phase 1: Preparation ---
    case_number = get_case_var(file_num, "case_number")
    
    if not case_number or str(case_number).lower() in ["null", "n/a", "none"]:
        log_event(f"Case number not found (or invalid) for {file_num}. Running Complaint Agent.")
        run_complaint_agent(file_num)
        
        # Check again
        case_number = get_case_var(file_num, "case_number")
        if not case_number or str(case_number).lower() in ["null", "n/a", "none"]:
            log_event("Error: Complaint Agent failed to retrieve case number. Terminating.", level="error")
            sys.exit(1)
            
    log_event(f"Found Case Number: {case_number}. Proceeding to Phase 2.")

    # --- Phase 2: Data Retrieval ---
    log_event("Launching hidden web browser to download docket...")
    
    # Run Scraper
    try:
        # Using sys.executable to ensure we use the same python env
        result = subprocess.run([sys.executable, SCRAPER_SCRIPT, case_number], capture_output=True, text=True)
        
        if result.returncode != 0:
            log_event(f"Docket Scraper failed.\nSTDOUT: {result.stdout}\nSTDERR: {result.stderr}", level="error")
            sys.exit(1)
        else:
            log_event("Docket Scraper completed successfully.")
            # log output for debug
            # log_event(f"Scraper Output: {result.stdout}")

    except Exception as e:
        log_event(f"Failed to execute Docket Scraper: {e}", level="error")
        sys.exit(1)

    # --- Post-Processing: Move File ---
    # Find the generated PDF (docket_YYYY.MM.DD.pdf)
    date_str = datetime.datetime.now().strftime("%Y.%m.%d")
    expected_filename = f"docket_{date_str}.pdf"
    
    if not os.path.exists(expected_filename):
        # Try looking for any docket_*.pdf if strict name match fails?
        # The scraper uses time.strftime, so it should match local time.
        # But let's be safe.
        candidates = glob.glob("docket_*.pdf")
        if candidates:
            # Sort by modification time, newest first
            candidates.sort(key=os.path.getmtime, reverse=True)
            expected_filename = candidates[0]
        else:
            log_event(f"Error: Generated PDF {expected_filename} not found in current directory.", level="error")
            sys.exit(1)

    case_path = get_case_path(file_num)
    if case_path:
        target_dir = os.path.join(case_path, "NOTES")
        if not os.path.exists(target_dir):
            os.makedirs(target_dir)
        
        new_filename = f"Docket_{date_str}.pdf"
        target_path = os.path.join(target_dir, new_filename)
        
        try:
            shutil.move(expected_filename, target_path)
            log_event(f"COMPLETED: Docket downloaded and saved to {target_path}")
            print(f"Success: Docket saved to {target_path}")
        except Exception as e:
            log_event(f"Error moving file to {target_path}: {e}", level="error")
    else:
        log_event("Could not resolve case path. Leaving file in current directory.", level="warning")

if __name__ == "__main__":
    main()
