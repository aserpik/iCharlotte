"""
Extract images from PDFs and save as separate JPEG files.
"""
import fitz  # PyMuPDF
from pathlib import Path
import sys

def extract_images_from_pdf(pdf_path: str, output_folder: str):
    """Extract all images from a PDF and save as JPEG files."""
    pdf_path = Path(pdf_path)
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)

    # Open the PDF
    doc = fitz.open(pdf_path)

    # Base name for output files (use PDF filename without extension)
    base_name = pdf_path.stem

    print(f"\nProcessing: {pdf_path.name}")
    print(f"Total pages: {len(doc)}")

    image_count = 0

    # Process each page
    for page_num in range(len(doc)):
        page = doc[page_num]

        # Get images on the page
        image_list = page.get_images()

        if image_list:
            # Extract embedded images
            for img_index, img in enumerate(image_list):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]

                # Save as JPEG
                output_filename = f"{base_name}_page{page_num + 1}_img{img_index + 1}.jpeg"
                output_path = output_folder / output_filename

                # If it's already a JPEG, just save it
                if image_ext == "jpeg" or image_ext == "jpg":
                    with open(output_path, "wb") as f:
                        f.write(image_bytes)
                else:
                    # Convert to JPEG using PIL
                    from PIL import Image
                    import io
                    img_pil = Image.open(io.BytesIO(image_bytes))
                    img_pil = img_pil.convert("RGB")
                    img_pil.save(output_path, "JPEG", quality=95)

                print(f"  Saved: {output_filename}")
                image_count += 1
        else:
            # No embedded images, render the page as an image
            # Use 2x zoom for good quality (300 DPI equivalent)
            mat = fitz.Matrix(2, 2)
            pix = page.get_pixmap(matrix=mat)

            output_filename = f"{base_name}_page{page_num + 1}.jpeg"
            output_path = output_folder / output_filename

            # Save as JPEG
            pix.save(str(output_path))
            print(f"  Rendered page as: {output_filename}")
            image_count += 1

    doc.close()
    print(f"Extracted {image_count} image(s) from {pdf_path.name}")
    return image_count

if __name__ == "__main__":
    # PDF files to process
    pdf_files = [
        r"Z:\Shared\Current Clients\3000 - PERMA\075 - Saltarelli v. City of Hesperia\FOCUS GROUP\Defense Presentation\Photos\Photos.pdf",
        r"Z:\Shared\Current Clients\3000 - PERMA\075 - Saltarelli v. City of Hesperia\FOCUS GROUP\Defense Presentation\Photos\59-Vehicle-Summary-Report-10-21-2022.pdf",
        r"Z:\Shared\Current Clients\3000 - PERMA\075 - Saltarelli v. City of Hesperia\FOCUS GROUP\Defense Presentation\Photos\Ex. 1.pdf"
    ]

    # Output folder
    output_folder = r"Z:\Shared\Current Clients\3000 - PERMA\075 - Saltarelli v. City of Hesperia\FOCUS GROUP\Defense Presentation\Photos"

    print("="*60)
    print("PDF Image Extractor")
    print("="*60)

    total_images = 0
    for pdf_path in pdf_files:
        if Path(pdf_path).exists():
            count = extract_images_from_pdf(pdf_path, output_folder)
            total_images += count
        else:
            print(f"\nWARNING: File not found: {pdf_path}")

    print("\n" + "="*60)
    print(f"Total images extracted: {total_images}")
    print("="*60)
