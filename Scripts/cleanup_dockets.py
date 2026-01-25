import os
import shutil
import sys
# Import from Scripts
sys.path.append(os.path.join(os.getcwd(), 'Scripts'))
from docket import get_case_path

# Map Case to File
case_map = {
    "21STCV01676": "3800.045",
    "CV-19-002704": "3800.082",
    "23VECV03479": "3800.089",
    "23STCV25278": "3800.133",
    "CIVSB2402748": "3800.156",
    "BCV-24-101066": "3800.160",
    "BC503381": "3800.168",
    "23SMCV04350": "3800.173",
    "37-2023-00008633-CU-CR-CTL": "3800.183",
    "CIVRS2400721": "3800.190"
}

gemini_dir = os.getcwd()
moved_count = 0

print("Starting cleanup...")

for filename in os.listdir(gemini_dir):
    if filename.startswith("docket_") and filename.endswith(".pdf"):
        # suffix _YYYY.MM.DD.pdf is 15 chars
        if len(filename) > 22: # minimal check
            case_num = filename[7:-15]
            
            # Check map
            file_num = case_map.get(case_num)
            if file_num:
                print(f"Matched {filename} to {file_num}")
                case_path = get_case_path(file_num)
                if case_path:
                    target_dir = os.path.join(case_path, "NOTES", "AI OUTPUT")
                    if not os.path.exists(target_dir):
                        os.makedirs(target_dir)
                    
                    src = os.path.join(gemini_dir, filename)
                    dst = os.path.join(target_dir, "Docket_2025.12.28.pdf") 
                    
                    try:
                        shutil.move(src, dst)
                        print(f"Moved to {dst}")
                        moved_count += 1
                    except Exception as e:
                        print(f"Error moving {filename}: {e}")
                else:
                    print(f"Could not resolve path for {file_num}")

print(f"Cleanup finished. Moved {moved_count} files.")