import win32com.client as win32
import os

doc_path = r"C:\GeminiTerminal\[DRAFT] Carrier 5800.057.docx"
pattern = "Plaintiffs"
replacement = "the Plaintiffs"

print(f"Opening {doc_path}")
word = win32.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

try:
    doc = word.Documents.Open(doc_path)
    print("Doc opened.")
    
    count = 0
    for i, para in enumerate(doc.Paragraphs):
        txt = para.Range.Text
        if pattern in txt:
            print(f"Match in Para {i+1}: {txt[:40]}...")
            
            rng = para.Range
            fnd = rng.Find
            fnd.ClearFormatting()
            fnd.Replacement.ClearFormatting()
            fnd.Text = pattern
            fnd.Replacement.Text = replacement
            fnd.MatchCase = True
            fnd.MatchWholeWord = False
            fnd.MatchWildcards = False
            
            success = fnd.Execute(Replace=2) # wdReplaceAll
            if success:
                print("  -> Replace Success")
                count += 1
            else:
                print("  -> Replace Failed (Execute returned False)")
                
    doc.Close(SaveChanges=False)
    print(f"Total Replacements: {count}")
    
except Exception as e:
    print(f"Error: {e}")
finally:
    # word.Quit()
    pass
