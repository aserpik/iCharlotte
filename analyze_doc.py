import win32com.client as win32
import os

doc_path = r"C:\GeminiTerminal\Carrier.docx"
if not os.path.exists(doc_path):
    print(f"File not found: {doc_path}")
    # Try the draft one if Carrier.docx doesn't exist, just in case user renamed it in chat but not on disk
    doc_path = r"C:\GeminiTerminal\[DRAFT] Carrier 5800.057.docx"

print(f"Analyzing: {doc_path}")

word = win32.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

try:
    doc = word.Documents.Open(doc_path)
    print(f"Opened. Paragraph count: {doc.Paragraphs.Count}")
    
    target = "Plaintiffs"
    
    for i, para in enumerate(doc.Paragraphs):
        txt = para.Range.Text
        # repr() shows hidden chars like \r, \t, etc.
        if "Plaintiff" in txt:
             print(f"Para {i+1} Python sees: {repr(txt[:50])}...")
             if target in txt:
                 print(f"  -> Python 'in' check PASSED for '{target}'")
             else:
                 print(f"  -> Python 'in' check FAILED for '{target}'")
                 
             # Try Word Find
             rng = para.Range
             fnd = rng.Find
             fnd.ClearFormatting()
             fnd.Text = target
             fnd.MatchCase = True
             if fnd.Execute():
                 print(f"  -> Word .Find.Execute PASSED for '{target}'")
             else:
                 print(f"  -> Word .Find.Execute FAILED for '{target}'")
                 
    doc.Close(SaveChanges=False)
except Exception as e:
    print(f"Error: {e}")
finally:
    pass
