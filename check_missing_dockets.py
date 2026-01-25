
import os
import glob
import json
import sys

# Ensure project root is in path
sys.path.append(os.getcwd())

from icharlotte_core.master_db import MasterCaseDatabase
from icharlotte_core.config import GEMINI_DATA_DIR
from icharlotte_core.utils import get_case_path

def check_missing_dockets():
    db = MasterCaseDatabase()
    cases = db.get_all_cases()
    
    missing_dockets = []

    print(f"Checking {len(cases)} cases for missing dockets...")
    print("-" * 60)
    print(f"{'File Number':<15} | {'County':<30}")
    print("-" * 60)

    for case in cases:
        file_number = case['file_number']
        case_path = case['case_path']
        
        # If case_path in DB is empty, try to resolve it
        if not case_path or not os.path.exists(case_path):
             case_path = get_case_path(file_number)

        has_docket = False
        if case_path and os.path.exists(case_path):
            ai_output_dir = os.path.join(case_path, "NOTES", "AI OUTPUT")
            if os.path.exists(ai_output_dir):
                dockets = glob.glob(os.path.join(ai_output_dir, "Docket*.pdf"))
                if dockets:
                    has_docket = True
        
        if not has_docket:
            # Get County
            county = "Unknown"
            
            # Try to load from JSON variable
            json_path = os.path.join(GEMINI_DATA_DIR, f"{file_number}.json")
            if os.path.exists(json_path):
                try:
                    with open(json_path, 'r', encoding='utf-8') as f:
                        vars_data = json.load(f)
                        # Look for 'venue_county', 'county', or 'venue'
                        # Structure might be key: value or key: {value: ...}
                        
                        def get_val(key):
                            v = vars_data.get(key)
                            if isinstance(v, dict) and "value" in v:
                                return v["value"]
                            return v

                        c = get_val("venue_county") or get_val("county") or get_val("venue")
                        if c:
                            county = str(c).strip()
                except:
                    pass
            
            print(f"{file_number:<15} | {county:<30}")
            missing_dockets.append((file_number, county))

    print("-" * 60)
    print(f"Total Missing Dockets: {len(missing_dockets)}")

if __name__ == "__main__":
    check_missing_dockets()
