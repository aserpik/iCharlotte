"""Extract Draw sections from GUILD001_GUILD2502 COMBINED.pdf"""
import fitz
import os

SOURCE = r"Z:\Shared\Current Clients\4500 - JBW DEFENSE\4553.001 - Mark Egerman\NOTES\GUILD001_GUILD2502 COMBINED.pdf"
OUT_DIR = os.path.dirname(SOURCE)

# (start_page, draw_name) — pages are 1-indexed as given by user
draws = [
    (231, "Draw 10"),
    (362, "Draw 11"),  # 362, 400 both listed — range starts at 362
    (410, "Draw 12"),
    (452, "Draw 13"),
    (497, "Draw 14"),
    (538, "Draw 15"),
    (606, "Draw 16"),
    (681, "Draw 17"),
    (779, "Draw 18"),
    (902, "Draw 19"),
    (1082, "Draw 2"),
    (1176, "Draw 20"),
    (1314, "Draw 21"),
    (1431, "Draw 22"),
    (1581, "Draw 23"),
    (1702, "Draw 24"),
    (1863, "Draw 25"),
    (1948, "Draw 26"),
    (2035, "Draw 27"),
    (2171, "Draw 28"),
    (2327, "Draw 3"),
    (2388, "Draw 9"),
]

# Sort by start page to compute ranges
draws.sort(key=lambda x: x[0])

doc = fitz.open(SOURCE)
total = doc.page_count  # 2502

for i, (start, name) in enumerate(draws):
    # End page is one before the next draw's start (or last page of doc)
    if i + 1 < len(draws):
        end = draws[i + 1][0] - 1
    else:
        end = total  # last draw goes to end

    # Convert to 0-indexed for PyMuPDF
    start_idx = start - 1
    end_idx = end - 1

    out_path = os.path.join(OUT_DIR, f"{name}.pdf")
    new_doc = fitz.open()
    new_doc.insert_pdf(doc, from_page=start_idx, to_page=end_idx)
    new_doc.save(out_path)
    new_doc.close()

    page_count = end - start + 1
    print(f"  {name}: pages {start}-{end} ({page_count} pages) -> {out_path}")

doc.close()
print(f"\nDone — {len(draws)} files created in {OUT_DIR}")
