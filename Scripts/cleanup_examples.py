import os

def cleanup_examples():
    input_path = os.path.join("Scripts", "status_update_examples.txt")
    output_path = os.path.join("Scripts", "status_update_examples_cleaned.txt")
    
    prefixes_to_remove = [
        "Factual Update:",
        "Factual Background:",
        "Procedural Status:"
    ]
    
    with open(input_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    cleaned_lines = []
    skip_next_blank = False
    
    for line in lines:
        stripped = line.strip()
        # Check if line starts with any of the prefixes
        should_remove = any(stripped.startswith(p) for p in prefixes_to_remove)
        
        if should_remove:
            # Skip this line
            skip_next_blank = True # Usually there's a blank line after these sections
            continue
        
        if skip_next_blank and stripped == "":
            skip_next_blank = False
            continue
            
        cleaned_lines.append(line)
        skip_next_blank = False

    with open(input_path, 'w', encoding='utf-8') as f:
        f.writelines(cleaned_lines)
    
    print(f"Cleaned {input_path}. Removed sections starting with: {', '.join(prefixes_to_remove)}")

if __name__ == "__main__":
    cleanup_examples()
