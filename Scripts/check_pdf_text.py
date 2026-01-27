"""
Diagnostic script to check what text is actually extracted from a PDF.
Usage: python check_pdf_text.py "path/to/file.pdf"
"""
import sys
import fitz

def check_pdf(path):
    print(f"Checking: {path}\n")
    print("=" * 60)

    try:
        doc = fitz.open(path)
        total_pages = len(doc)
        print(f"Total pages: {total_pages}\n")

        total_chars = 0
        for i, page in enumerate(doc):
            text = page.get_text()
            char_count = len(text.strip())
            total_chars += char_count

            print(f"--- Page {i+1} ({char_count} chars) ---")
            if char_count > 0:
                # Show first 500 chars of each page
                preview = text.strip()[:500]
                print(preview)
                if len(text.strip()) > 500:
                    print("... [truncated]")
            else:
                print("[NO TEXT EXTRACTED]")
            print()

        doc.close()

        print("=" * 60)
        print(f"TOTAL CHARACTERS EXTRACTED: {total_chars}")

        if total_chars < 100:
            print("\n⚠️  WARNING: Very little text extracted!")
            print("This PDF is likely a scanned image and needs OCR.")
        elif total_chars < 1000:
            print("\n⚠️  WARNING: Low text content.")
            print("This may be a scanned PDF with poor OCR or metadata only.")
        else:
            print("\n✓ PDF has extractable text.")

    except Exception as e:
        print(f"ERROR: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python check_pdf_text.py <pdf_path>")
        sys.exit(1)

    path = " ".join(sys.argv[1:]).strip('"').strip("'")
    check_pdf(path)
