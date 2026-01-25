import os
import glob
import win32com.client
import pythoncom

def extract_updates():
    # 1. Setup Paths
    base_dir = os.getcwd()
    source_folder = os.path.join(base_dir, "Interim Updates")
    output_file = os.path.join(base_dir, "Scripts", "status_update_examples.txt")

    print(f"Scanning folder: {source_folder}")

    # 2. Initialize Outlook
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(f"Error initializing Outlook: {e}")
        return

    # 3. Find all .msg files
    # We use glob to handle the spaces in filenames easily
    msg_files = glob.glob(os.path.join(source_folder, "*.msg"))
    
    if not msg_files:
        print("No .msg files found!")
        return

    print(f"Found {len(msg_files)} .msg files. Extracting text...")

    success_count = 0
    
    with open(output_file, "w", encoding="utf-8") as out:
        # Add a header explaining this file
        out.write("=== STATUS UPDATE EXAMPLES ===\n")
        out.write("This file contains the text bodies of previous status updates.\n")
        out.write("Use these as style references for the LLM.\n\n")

        for msg_path in msg_files:
            try:
                # OpenSharedItem requires an absolute path
                abs_path = os.path.abspath(msg_path)
                
                # Load the .msg file
                item = namespace.OpenSharedItem(abs_path)
                
                # Extract details
                subject = item.Subject
                body = item.Body
                
                # Write to our text file with clear separators
                out.write("=" * 80 + "\n")
                out.write(f"EXAMPLE SOURCE: {os.path.basename(msg_path)}\n")
                out.write(f"SUBJECT: {subject}\n")
                out.write("-" * 80 + "\n")
                out.write(body)
                out.write("\n") # Extra newline for spacing
                
                # Cleanup (important for COM objects)
                # We don't save changes to the msg file
                item.Close(0) # 0 = olDiscard
                del item

                success_count += 1
                print(f"Extracted: {subject[:50]}...")

            except Exception as e:
                print(f"Failed to extract {os.path.basename(msg_path)}: {e}")

    print(f"\nExtraction complete.")
    print(f"Successfully processed {success_count} files.")
    print(f"Saved to: {output_file}")

if __name__ == "__main__":
    extract_updates()
