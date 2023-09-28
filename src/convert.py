import os
from win32com import client
from pytesseract import pytesseract
from docx import Document
import fitz  # PyMuPDF
from PIL import Image
import pytesseract

def convert_pdf_to_docx(pdf_folder_path: str, docx_output_path: str) -> None:
    """
    Convert all PDF files in the specified folder to DOCX format using OCR.

    This function iterates over all files in the specified folder, and for each file with a .pdf extension,
    it performs OCR (Optical Character Recognition) to convert it to DOCX format. The DOCX file is saved
    in the specified output directory.

    :param pdf_folder_path: The path to the folder containing PDF files.
    :type pdf_folder_path: str
    :param docx_output_path: The path to the output directory where DOCX files will be saved.
    :type docx_output_path: str
    :return: None
    """
    for file_name in os.listdir(pdf_folder_path):
        if file_name.endswith('.pdf'):
            pdf_file_path = os.path.join(pdf_folder_path, file_name)
            doc = fitz.open(pdf_file_path)
            docx_file_name = file_name.replace('.pdf', '.docx')
            docx_file_path = os.path.join(docx_output_path, docx_file_name)
            word = Document()
            for page_number in range(len(doc)):
                page = doc.loadPage(page_number)
                image_list = page.get_images(full=True)
                for img_index, img in enumerate(image_list, start=1):
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    image_data = base_image["image"]
                    image = Image.open(io.BytesIO(image_data))
                    text = pytesseract.image_to_string(image)
                    word.add_paragraph(text)
            word.save(docx_file_path)

if __name__ == "__main__":
    # Example usage:
    # convert_pdf_to_docx('pdfs/', 'docx/')