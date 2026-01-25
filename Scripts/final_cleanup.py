import os
import re

def final_examples_cleanup():
    path = os.path.join("Scripts", "status_update_examples.txt")
    if not os.path.exists(path):
        return
        
    with open(path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    cleaned = []
    # Remove lines starting with SUBJECT: or EXAMPLE SOURCE:
    for line in lines:
        s = line.strip()
        if s.startswith("SUBJECT:") or s.startswith("EXAMPLE SOURCE:"):
            continue
        cleaned.append(line)
        
    with open(path, 'w', encoding='utf-8') as f:
        f.writelines(cleaned)
    print("Cleaned SUBJECT and EXAMPLE SOURCE from examples.")

if __name__ == "__main__":
    final_examples_cleanup()
