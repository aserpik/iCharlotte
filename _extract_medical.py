"""Extract medical treatment info from Haire case docx files."""
from docx import Document
import os

files = [
    r'Z:\Shared\Current Clients\3100 - PARSAC\094 - Haire v. City of Wildomar\NOTES\AI OUTPUT\Discovery_Responses_Plaintiff_ANDREW_HAIRE.docx',
    r'Z:\Shared\Current Clients\3100 - PARSAC\094 - Haire v. City of Wildomar\STATUS\[Draft] Carrier003 - Medical Record Review.docx',
    r'Z:\Shared\Current Clients\3100 - PARSAC\094 - Haire v. City of Wildomar\STATUS\[Draft] Carrier004.docx',
    r'Z:\Shared\Current Clients\3100 - PARSAC\094 - Haire v. City of Wildomar\STATUS\[draft2] Carrier004.docx',
]

output_path = r'C:\geminiterminal2\_medical_output.txt'

with open(output_path, 'w', encoding='utf-8') as out:
    for f in files:
        out.write(f'\n\n{"="*80}\n')
        out.write(f'FILE: {os.path.basename(f)}\n')
        out.write('='*80 + '\n')
        try:
            doc = Document(f)
            for p in doc.paragraphs:
                if p.text.strip():
                    out.write(p.text + '\n')
        except Exception as e:
            out.write(f'ERROR: {e}\n')

print(f"Done. Output written to {output_path}")
