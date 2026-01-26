import os
import glob
import subprocess

def extract_text(file_path):
    try:
        # Run antiword
        result = subprocess.run(['antiword', '-f', file_path], capture_output=True, text=True, encoding='utf-8', errors='replace')
        if result.returncode != 0:
            return f"[Error running antiword: {result.stderr}]"
        return result.stdout
    except Exception as e:
        return f"[Exception: {e}]"

def main():
    target_dir = r"C:\geminiterminal2\LLM Resources\Calendaring\LLM Scan"
    files = glob.glob(os.path.join(target_dir, "*.doc"))
    files.sort()

    all_text = ""
    for f_path in files:
        if os.path.basename(f_path).startswith("~$"):
            continue
            
        print(f"Processing: {os.path.basename(f_path)}")
        text = extract_text(f_path)
        all_text += "\n\n==============================================================================\n"
        all_text += f"FILE: {os.path.basename(f_path)}\n"
        all_text += "==============================================================================\n\n"
        all_text += text
        
    with open("extracted_text_antiword.txt", "w", encoding="utf-8") as f:
        f.write(all_text)

if __name__ == "__main__":
    main()
