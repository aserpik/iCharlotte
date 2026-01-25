import sys
import os
from docx import Document
from docx.shared import Pt

def analyze_docx(file_path):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    try:
        doc = Document(file_path)
        print(f"Analyzing: {file_path}")
        print("-" * 30)

        for i, para in enumerate(doc.paragraphs):
            # We assume the user is looking at the first few pages, so let's look at the first 50 paragraphs
            if i > 50: 
                break
            
            text = para.text
            style = para.style.name
            
            # Spacing
            pf = para.paragraph_format
            space_before = pf.space_before.pt if pf.space_before else "None"
            space_after = pf.space_after.pt if pf.space_after else "None"
            line_spacing = pf.line_spacing
            
            # Numbering (XML check for list items)
            is_list = False
            if para._p.pPr is not None and para._p.pPr.numPr is not None:
                is_list = True

            content_repr = repr(text)
            if len(content_repr) > 50:
                content_repr = content_repr[:47] + "..."

            print(f"Para {i}: Style='{style}' Text={content_repr}")
            print(f"    Spacing: Before={space_before}, After={space_after}, Line={line_spacing}")
            print(f"    Is List: {is_list}, Empty: {not text.strip()}")
            print("-" * 10)

    except Exception as e:
        print(f"Error reading docx: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        analyze_docx(sys.argv[1])
    else:
        print("Usage: python analyze_doc_structure.py <path_to_docx>")
