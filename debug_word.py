import win32com.client
import os

def main():
    f_path = r"C:\geminiterminal2\LLM Resources\Calendaring\LLM Scan\Extending Time in State Court Litigation (CA).doc"
    f_path = os.path.abspath(f_path)
    print(f"Path: {f_path}")
    
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True # Make it visible to see if errors pop up
        word.DisplayAlerts = -1 # wdAlertsAll
        
        print("Opening...")
        doc = word.Documents.Open(f_path)
        print(f"Doc type: {type(doc)}")
        print(f"Doc repr: {repr(doc)}")
        
        print("Text len:", len(doc.Content.Text))
        
        doc.Close(False)
        word.Quit()
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
