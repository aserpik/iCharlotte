import os
import re

def advanced_cleanup():
    input_path = os.path.join("Scripts", "status_update_examples.txt")
    # We will write back to the same file or a new one? Let's write to a new one then overwrite.
    output_path = os.path.join("Scripts", "status_update_examples_final.txt")
    
    with open(input_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    cleaned_lines = []
    skip_mode = False # To skip signature blocks
    
    # Compile regexes for speed and clarity
    # 1. Paragraphs (already mostly gone, but ensuring)
    re_section_headers = re.compile(r"^(Factual Background|Procedural Status|Factual Update):", re.IGNORECASE)
    
    # 2. Intro Sentences
    re_intro = re.compile(r"Please allow the .* to serve as", re.IGNORECASE)
    
    # 3. Salutation (Name + Comma)
    # Matches "Robin," or "David," but not "Hello," or "Dear Robin," necessarily, strictly single name + comma
    re_salutation = re.compile(r"^[A-Z][a-z]+,\s*$")
    
    # 4. Closing sentences
    re_closing_q = re.compile(r"^If you have any questions", re.IGNORECASE)
    
    # 5. "Best,"
    re_best = re.compile(r"^Best,\s*$", re.IGNORECASE)
    
    # 6. "Andrei" (just the word on a line)
    re_andrei_word = re.compile(r"^Andrei\s*$", re.IGNORECASE)
    
    # 7. Signature Start (Andrei V. Serpik, Attorney, Bordin Semmer)
    # We will match "ANDREI V. SERPIK" or "BORDIN SEMMER"
    re_sig_start = re.compile(r"^(ANDREI V\. SERPIK|BORDIN SEMMER)", re.IGNORECASE)
    
    # Delimiters that end a signature block
    re_delimiter = re.compile(r"^(={10,}|-{10,}|EXAMPLE SOURCE:)")

    for line in lines:
        stripped = line.strip()
        
        # Check if we are in skip mode (inside a signature)
        if skip_mode:
            # If we hit a delimiter, turn off skip mode and process this line (it's the delimiter)
            if re_delimiter.match(line): # Check original line for delimiter pattern
                skip_mode = False
            else:
                # Still in signature, skip this line
                continue
        
        # Check for Signature Start
        if re_sig_start.match(stripped):
            skip_mode = True
            continue

        # Check other removal criteria
        if re_section_headers.match(stripped):
            continue
            
        if re_intro.search(stripped):
            continue
            
        if re_salutation.match(stripped):
            continue
            
        if re_closing_q.match(stripped):
            continue
            
        if re_best.match(stripped):
            continue
            
        if re_andrei_word.match(stripped):
            continue
            
        # If we passed all checks, keep the line
        cleaned_lines.append(line)

    # Post-processing to remove excess blank lines
    final_lines = []
    blank_count = 0
    
    for line in cleaned_lines:
        if line.strip() == "":
            blank_count += 1
            if blank_count <= 2: # Allow max 2 consecutive blank lines
                final_lines.append(line)
        else:
            blank_count = 0
            final_lines.append(line)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.writelines(final_lines)
    
    # Overwrite the original
    os.replace(output_path, input_path)
    print(f"Advanced cleanup complete. Processed {len(lines)} lines -> {len(final_lines)} lines.")

if __name__ == "__main__":
    advanced_cleanup()
