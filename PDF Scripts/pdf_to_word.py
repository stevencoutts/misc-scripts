import sys
import os
import pdfplumber
from docx import Document
from docx.shared import Inches
from PIL import Image
import io
import pytesseract


def pdf_to_word(pdf_path, docx_path):
    document = Document()
    with pdfplumber.open(pdf_path) as pdf:
        print(f"Opened PDF: {pdf_path}, {len(pdf.pages)} pages")
        for page_num, page in enumerate(pdf.pages, 1):
            print(f"Processing page {page_num}/{len(pdf.pages)}")
            # Extract text with basic layout
            text = page.extract_text(x_tolerance=2, y_tolerance=2)
            if text:
                print(f"  Extracting text from page {page_num}")
                for line in text.split('\n'):
                    document.add_paragraph(line)
            else:
                print(f"  No text found on page {page_num}, running OCR...")
                # Render page as image for OCR
                try:
                    page_image = page.to_image(resolution=300).original
                    ocr_text = pytesseract.image_to_string(page_image)
                    if ocr_text.strip():
                        print(f"    OCR extracted text from page {page_num}")
                        for line in ocr_text.split('\n'):
                            document.add_paragraph(line)
                    else:
                        print(f"    OCR found no text on page {page_num}")
                except Exception as e:
                    print(f"    OCR failed on page {page_num}: {e}")
            # Extract images
            for img_index, img in enumerate(page.images):
                print(f"  Extracting image {img_index+1}/{len(page.images)} on page {page_num}")
                # Extract image bytes
                x0, top, x1, bottom = img["x0"], img["top"], img["x1"], img["bottom"]
                # Clamp bbox to page bbox
                page_x0, page_top, page_x1, page_bottom = page.bbox
                x0 = max(x0, page_x0)
                top = max(top, page_top)
                x1 = min(x1, page_x1)
                bottom = min(bottom, page_bottom)
                if x0 >= x1 or top >= bottom:
                    print(f"    Skipping invalid bbox for image {img_index+1}")
                    continue  # Skip invalid bbox
                try:
                    cropped = page.within_bbox((x0, top, x1, bottom)).to_image(resolution=300)
                    img_bytes = cropped.original.convert("RGB")
                    img_stream = io.BytesIO()
                    img_bytes.save(img_stream, format="PNG")
                    img_stream.seek(0)
                    document.add_picture(img_stream, width=Inches(4))
                    print(f"    Image {img_index+1} added to document")
                except Exception as e:
                    print(f"Warning: Could not extract image on page {page_num}, image {img_index}: {e}")
            if page_num != len(pdf.pages):
                document.add_page_break()
    print(f"Saving DOCX to {docx_path}")
    document.save(docx_path)
    print(f"Saved DOCX to {docx_path}")


def main():
    if len(sys.argv) != 3:
        print(f"Usage: python {os.path.basename(__file__)} input.pdf output.docx")
        sys.exit(1)
    pdf_path = sys.argv[1]
    docx_path = sys.argv[2]
    if not os.path.isfile(pdf_path):
        print(f"Error: File '{pdf_path}' does not exist.")
        sys.exit(1)
    pdf_to_word(pdf_path, docx_path)
    print(f"Converted '{pdf_path}' to '{docx_path}' successfully.")


if __name__ == "__main__":
    main() 