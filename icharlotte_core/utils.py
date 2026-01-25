import os
import re
import glob
import logging
import io
import datetime
from .config import LOG_FILE, BASE_PATH_WIN

try:
    import fitz # PyMuPDF
except ImportError:
    fitz = None

try:
    from docx import Document
except ImportError:
    Document = None

# Try to import CaseDataManager
try:
    from Scripts.case_data_manager import CaseDataManager
except ImportError:
    CaseDataManager = None

# --- Logging Setup ---
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

def parse_hearing_data(hearings_str):
    """
    Parses a string of hearings (comma or newline separated) into a list of dicts:
    [{'description': 'CMC', 'date_obj': datetime, 'date_sort': '2026-02-05', 'display': 'CMC 2/5/26'}, ...]
    Sorted by date.
    """
    if not hearings_str or str(hearings_str).lower() == "none":
        return []

    # Split by newlines or commas
    # If distinct lines, use lines. If single line with commas, use commas.
    if "\n" in hearings_str:
        segments = hearings_str.split("\n")
    else:
        segments = hearings_str.split(",")

    parsed = []
    
    # Regex for various date formats
    # M/D/YY, M/D/YYYY, YYYY-MM-DD, MM-DD-YYYY
    date_patterns = [
        r"(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", # M/D/Y or M-D-Y
        r"(\d{4})-(\d{2})-(\d{2})" # Y-M-D
    ]
    
    for seg in segments:
        seg = seg.strip()
        if not seg: continue
        
        found_date = None
        date_obj = None
        
        # Try to find a date
        for pat in date_patterns:
            match = re.search(pat, seg)
            if match:
                g = match.groups()
                try:
                    if len(g) == 3:
                        if len(g[0]) == 4: # YYYY-MM-DD
                            date_obj = datetime.datetime(int(g[0]), int(g[1]), int(g[2]))
                        else: # M/D/Y
                            y = int(g[2])
                            if y < 100: y += 2000 # Assume 20xx
                            date_obj = datetime.datetime(y, int(g[0]), int(g[1]))
                        
                        found_date = match.group(0)
                        break
                except:
                    continue
        
        if date_obj:
            # Clean description
            desc = seg.replace(found_date, "").strip()
            # Remove " on ", " at "
            desc = re.sub(r"(?i)\s+(on|at)\s*$", "", desc)
            desc = re.sub(r"^\s*(on|at)\s+", "", desc)
            desc = desc.strip(" -:,")
            if not desc: desc = "Hearing"

            # Apply Abbreviations Mapping
            abbrev_map = [
                (r"(?i)Hearing on Motion to continue trial", "Cnt Trial"),
                (r"(?i)(Initial\s+)?Case Management Conference", "CMC"),
                (r"(?i)Post-Mediation Status Conference", "Med. SC"),
                (r"(?i)Mandatory Settlement Conference", "MSC"),
                (r"(?i)Trial Setting Conference", "TSC"),
                (r"(?i)Motion for Summary Judgment", "MSJ"),
                (r"(?i)Trial Readiness Conference", "TRC"),
                (r"(?i)Final Status Conference", "FSC"),
                (r"(?i)Order to Show Cause(?!.*Sanctions)", "OSC"),
                (r"(?i)OSC Re: Sanctions", "OSC Sanctions")
            ]
            for pattern, abbrev in abbrev_map:
                if re.search(pattern, desc):
                    desc = abbrev
                    break
            
            # Format user requested: "CMC 2/5/26"
            display = f"{desc} {date_obj.month}/{date_obj.day}/{date_obj.strftime('%y')}"
            
            parsed.append({
                'description': desc,
                'date_obj': date_obj,
                'date_sort': date_obj.strftime("%Y-%m-%d"),
                'display': display,
                'raw': seg
            })
        else:
            # No date found, append to end with far future or past? 
            # If no date, maybe ignore or put at bottom?
            # User wants "Next Hearing", so dateless ones are tricky. 
            # Let's include them with a null sort date.
            parsed.append({
                'description': seg,
                'date_obj': datetime.datetime.max, # Put at end
                'date_sort': "9999-99-99",
                'display': seg,
                'raw': seg
            })

    # Sort by date
    parsed.sort(key=lambda x: x['date_obj'])
    
    return parsed

def format_date_to_mm_dd_yyyy(date_val):
    """
    Attempts to format various date inputs (string or mtime) into mm-dd-yyyy.
    """
    if not date_val:
        return ""
        
    if isinstance(date_val, (int, float)):
        try:
            if date_val <= 0: return ""
            dt = datetime.datetime.fromtimestamp(date_val)
            return dt.strftime("%m-%d-%Y")
        except:
            return ""

    if isinstance(date_val, str):
        date_str = date_val.strip()
        if not date_str or date_str.lower() == "none":
            return ""
        # Try common formats
        formats = [
            "%Y-%m-%d %H:%M", "%Y-%m-%d", "%m/%d/%Y", 
            "%m-%d-%Y", "%d/%m/%Y", "%B %d, %Y", 
            "%b %d, %Y", "%Y.%m.%d"
        ]
        for fmt in formats:
            try:
                dt = datetime.datetime.strptime(date_str, fmt)
                return dt.strftime("%m-%d-%Y")
            except ValueError:
                continue
        return date_str # Return original if all else fails
        
    return str(date_val)

def log_event(message, level="info"):
    print(message)
    if level == "info":
        logging.info(message)
    elif level == "error":
        logging.error(message)
    elif level == "warning":
        logging.warning(message)

def sanitize_filename(name):
    return "".join(c for c in name if c.isalnum() or c in (' ', '-', '_')).strip()

def get_case_path(file_num):
    """Locates the physical directory for a file number (####.###)."""
    # --- HARD OVERRIDES (CRITICAL) ---
    if file_num == "1280.001":
        return os.path.abspath(r"Z:\Shared\Current Clients\1000 - PLAINTIFF\1280.001 - Natalia Finkovskaya")
    
    if file_num in ["1011.001", "1011.011"]:
        return os.path.abspath(r"Z:\Shared\Current Clients\1000 - PLAINTIFF\1011 - Brenna Cox Johnson\1011.011 Iwanaga\JANE DOE v. RALPH DILLON")

    if file_num == "4519.003":
        return os.path.abspath(r"Z:\Shared\Current Clients\4500 - JBW DEFENSE\4519 - L1 Technologies\003 - Bushnell")

    if file_num == "4537.001":
        return os.path.abspath(r"Z:\Shared\Current Clients\4500 - JBW DEFENSE\4537.001 - Nixon")

    if file_num == "4553.001":
        return os.path.abspath(r"Z:\Shared\Current Clients\4500 - JBW DEFENSE\4553.001 - Mark Egerman")

    if file_num == "4558.001":
        return os.path.abspath(r"Z:\Shared\Current Clients\4500 - JBW DEFENSE\4558.001 - Pinscreen")

    if file_num == "4573.001":
        return os.path.abspath(r"Z:\Shared\Current Clients\4500 - JBW DEFENSE\4573.001 - Gutierrez")

    # --- CHECK CASEDATAMANAGER FOR OVERRIDES ---
    if CaseDataManager:
        try:
            manager = CaseDataManager()
            stored_path = manager.get_value(file_num, "file_path")
            if stored_path and os.path.exists(stored_path):
                return os.path.abspath(stored_path)
        except Exception as e:
            # Fallback if CaseDataManager fails
            pass

    # --- HARD OVERRIDES ---
    if file_num == "1011.011":
        path = r"Z:\Shared\Current Clients\1000 - PLAINTIFF\1011 - Brenna Cox Johnson\1011.011 Iwanaga\JANE DOE v. RALPH DILLON"
        if os.path.exists(path):
            return os.path.abspath(path)

    if not re.match(r'^\d{4}\.\d{3}$', file_num):
        log_event(f"Invalid file number format: {file_num}", "error")
        return None

    carrier_num, case_num = file_num.split('.')
    base_path = BASE_PATH_WIN

    # --- SPECIAL RULE: 1000 - PLAINTIFF (Starts with 10 or 12) ---
    if file_num.startswith("10") or file_num.startswith("12"):
        plaintiff_root = os.path.join(base_path, "1000 - PLAINTIFF")
        if os.path.exists(plaintiff_root):
            # Strategy A: Direct Match (e.g. 1280.001 - Name)
            direct_matches = glob.glob(os.path.join(plaintiff_root, f"{file_num}*"))
            if direct_matches:
                return os.path.abspath(direct_matches[0])
            
            # Strategy B: Nested in Carrier (e.g. 1011 - Name / 1011.011 - Name)
            carrier_matches = glob.glob(os.path.join(plaintiff_root, f"{carrier_num}*"))
            for cm in carrier_matches:
                # Look for file number inside carrier folder
                nested_matches = glob.glob(os.path.join(cm, f"{file_num}*"))
                if nested_matches:
                    case_folder = nested_matches[0]
                    
                    # Special Heuristic for 1011.011 (or deep nesting):
                    # If the folder contains NO files but DOES contain exactly 1 subfolder, take the subfolder.
                    # (Matches user example: ...\1011.011 Iwanaga\JANE DOE v. RALPH DILLON)
                    try:
                        items = os.listdir(case_folder)
                        subfolders = [i for i in items if os.path.isdir(os.path.join(case_folder, i))]
                        files = [i for i in items if os.path.isfile(os.path.join(case_folder, i))]
                        
                        if len(subfolders) == 1 and len(files) < 3: # Allow loose files like .DS_Store or thumbs
                             return os.path.abspath(os.path.join(case_folder, subfolders[0]))
                    except:
                        pass
                        
                    return os.path.abspath(case_folder)

    # --- SPECIAL RULE: 45XXX.001 -> 4500 - JBW DEFENSE ---
    if file_num.startswith("45") and file_num.endswith(".001"):
        # The user said "file numbers that begin with 45XXX.001", but then gave "4500 - JBW DEFENSE" as the carrier folder.
        # This implies standard carrier/case structure might not apply strictly, OR we just force the carrier path.
        # The prompt says: "for all file numbers that begin with 45XXX.001, the case folders can all be found within the following carrier folder: Z:\Shared\Current Clients\4500 - JBW DEFENSE"
        # So we force the carrier path.
        special_carrier_path = os.path.join(base_path, "4500 - JBW DEFENSE")
        
        # Now look for the case folder inside. 
        # Usually case folder is "### - Name" (suffix) or "####.### - Name" (full) or "##### - Name" (something else)?
        # The standard logic below looks for `case_num` (suffix) - *.
        # For 45123.001, carrier=45123 (no, 45123 isn't 4 chars), case=001.
        # Wait, splitting `45123.001` by `.` gives `45123` and `001`? 
        # No, standard file numbers here are 4 digits dot 3 digits (e.g., 3200.284).
        # The prompt says "45XXX.001". This looks like 5 digits dot 3 digits? 
        # Or maybe carrier is `45`? `45XXX`?
        # Let's assume standard splitting `carrier_num, case_num = file_num.split('.')` works if it matches regex `^\d{4}\.\d{3}$`.
        # BUT `45XXX.001` might imply `45` then 3 variable digits then `.001`.
        # If the file number is `4512.001`, then carrier=4512, case=001.
        # If the file number is `45001.001` (5 digits), the regex `^\d{4}\.\d{3}$` will FAIL.
        # Let's adjust the regex to allow 4 or 5 digits for carrier if needed?
        # The user wrote "45XXX.001". If X is a digit, that is 5 digits total prefix.
        # Let's relax the regex first.
        
    # RELAXED REGEX to allow 4-5 digit prefixes
    if not re.match(r'^\d{4,5}\.\d{3}$', file_num):
        log_event(f"Invalid file number format: {file_num}", "error")
        return None
    
    carrier_num, case_num = file_num.split('.')

    if not os.path.exists(base_path):
        log_event(f"Base path not found: {base_path}", "error")
        return None

    # Find Carrier Folder
    carrier_path = None
    
    # SPECIAL RULE: 45XXX.001
    if file_num.startswith("45") and case_num == "001":
         carrier_path = os.path.join(base_path, "4500 - JBW DEFENSE")
    elif carrier_num == "3850":
        parent_3800_candidates = glob.glob(os.path.join(base_path, "3800*"))
        for p in parent_3800_candidates:
            possible_path = os.path.join(p, "3850")
            if os.path.exists(possible_path):
                carrier_path = possible_path
                break
    else:
        potential_folders = glob.glob(os.path.join(base_path, f"{carrier_num}*"))
        for p in potential_folders:
            folder_name = os.path.basename(p)
            if folder_name.startswith(f"{carrier_num}"):
                remainder = folder_name[len(carrier_num):]
                if not remainder or remainder[0] not in '0123456789':
                    carrier_path = p
                    break

    if not carrier_path or not os.path.exists(carrier_path):
        # Fallback logging or return
        if not carrier_path:
             log_event(f"Carrier folder not found for: {carrier_num}", "error")
        return None

    # Find Case Folder
    # Standard logic searches for "case_num - *". 
    # For 45XXX.001, case_num is 001. 
    # Does the folder look like "001 - Name"? Or does it use the FULL number?
    # Usually "3200.284" -> "3200" folder -> "284 - Name" folder.
    # If this special client uses "45XXX - Name" inside "4500", we might need different logic.
    # But sticking to standard pattern first: look for `case_num - *`.
    # However, if `45XXX.001` implies the `XXX` is the important part?
    # Let's assume standard behavior first.
    
    # Try finding folder starting with case_num (001)
    case_folders = glob.glob(os.path.join(carrier_path, f"{case_num} - *"))
    
    # If not found, and it's the special JBW case, maybe look for the prefix part?
    # If file is 45123.001, and folder is "45123 - Name"?
    if not case_folders and file_num.startswith("45") and case_num == "001":
         # Try looking for the full prefix inside the JBW folder?
         # "45123.001" -> carrier="45123".
         # Search for "45123 - *"
         case_folders = glob.glob(os.path.join(carrier_path, f"{carrier_num} - *"))

    if not case_folders:
        log_event(f"Case folder not found for: {case_num} in {carrier_path}", "error")
        return None
    
    return os.path.abspath(case_folders[0])

def extract_text_from_file(file_path):
    """
    Extracts text from PDF, DOCX, or TXT files.
    Returns the extracted text as a string, or None if failed.
    """
    if not os.path.exists(file_path):
        return None

    ext = os.path.splitext(file_path)[1].lower()
    
    try:
        if ext == '.pdf':
            if not fitz:
                return "[Error: PyMuPDF not installed]"
            doc = fitz.open(file_path)
            text = ""
            for page in doc:
                text += page.get_text() + "\n"
            return text
            
        elif ext == '.docx':
            if not Document:
                return "[Error: python-docx not installed]"
            doc = Document(file_path)
            text = "\n".join([p.text for p in doc.paragraphs])
            return text
            
        elif ext in ['.txt', '.md', '.py', '.json', '.xml', '.html', '.htm', '.msg']:
             # For .msg, we really should use Outlook or extract-msg, but here we'll stick to text-based or rely on the user to convert.
             # Wait, the prompt implies "drag and drop files (including PDFs, word documents, txt files, etc.)".
             # For .msg specifically, if we are in this app context, we might use win32com if needed, but for now treat as text/fail.
             with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                 return f.read()
        
        else:
            return f"[Unsupported file type: {ext}]"

    except Exception as e:
        log_event(f"Error extracting text from {file_path}: {e}", "error")
        return None