import sys
import os
import pdfplumber
from docx import Document
from docx.shared import Inches
from PIL import Image
import io


def pdf_to_word(pdf_path, docx_path):
    document = Document()
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            # Extract text with basic layout
            text = page.extract_text(x_tolerance=2, y_tolerance=2)
            if text:
                for line in text.split('\n'):
                    document.add_paragraph(line)
            # Extract images
            for img_index, img in enumerate(page.images):
                # Extract image bytes
                x0, top, x1, bottom = img["x0"], img["top"], img["x1"], img["bottom"]
                # Clamp bbox to page bbox
                page_x0, page_top, page_x1, page_bottom = page.bbox
                x0 = max(x0, page_x0)
                top = max(top, page_top)
                x1 = min(x1, page_x1)
                bottom = min(bottom, page_bottom)
                if x0 >= x1 or top >= bottom:
                    continue  # Skip invalid bbox
                try:
                    cropped = page.within_bbox((x0, top, x1, bottom)).to_image(resolution=300)
                    img_bytes = cropped.original.convert("RGB")
                    img_stream = io.BytesIO()
                    img_bytes.save(img_stream, format="PNG")
                    img_stream.seek(0)
                    document.add_picture(img_stream, width=Inches(4))
                except Exception as e:
                    print(f"Warning: Could not extract image on page {page_num}, image {img_index}: {e}")
            if page_num != len(pdf.pages):
                document.add_page_break()
    document.save(docx_path)


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