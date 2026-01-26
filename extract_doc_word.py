import sys
import os
import glob
import win32com.client
import time
import pythoncom

def main():
    target_dir = r"C:\geminiterminal2\LLM Resources\Calendaring\LLM Scan"
    files = glob.glob(os.path.join(target_dir, "*.doc"))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    files.sort()

    all_text = ""
    
    word = None
    try:
        pythoncom.CoInitialize()
        try:
            word = win32com.client.gencache.EnsureDispatch("Word.Application")
        except:
            word = win32com.client.Dispatch("Word.Application")
            
        word.Visible = False
        word.DisplayAlerts = 0
        
        for f_path in files:
            print(f"Processing: {os.path.basename(f_path)}")
            all_text += f"\n\n==============================================================================\n"
            all_text += f"FILE: {os.path.basename(f_path)}\n"
            all_text += f"==============================================================================\n\n"
            
            doc = None
            try:
                abs_path = os.path.abspath(f_path)
                doc = word.Documents.Open(FileName=abs_path, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False)
                
                # Try to get text
                try:
                    all_text += doc.Content.Text
                except Exception as e_text:
                    all_text += f"[Content Access Failed: {repr(e_text)}]"
                    try:
                        all_text += doc.Range().Text
                    except Exception as e_range:
                        all_text += f"[Range Access Failed: {repr(e_range)}]"

                doc.Close(False)
                doc = None
            except Exception as e_open:
                all_text += f"Error opening/processing {os.path.basename(f_path)}: {repr(e_open)}"
            finally:
                if doc:
                    try:
                        doc.Close(False)
                    except:
                        pass
            
            time.sleep(0.5)

    except Exception as e_app:
        all_text += f"CRITICAL WORD APP ERROR: {repr(e_app)}"
    finally:
        if word:
            try:
                word.Quit()
            except:
                pass
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    with open("extracted_text_word.txt", "w", encoding="utf-8") as f:
        f.write(all_text)

if __name__ == "__main__":
    main()