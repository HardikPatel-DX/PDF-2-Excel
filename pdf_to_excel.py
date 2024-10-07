import pdfplumber
import pandas as pd
import camelot
import tesserocr
from pdf2image import convert_from_path
from PIL import Image
import io

# Function to clean and format extracted table data
def clean_table_data(df):
    """
    Cleans the extracted tabular data.
    - Sets the first row as header if applicable.
    - Resets the index.
    """
    if not df.empty:
        df.columns = df.iloc[0]  # Set the first row as header
        df = df.drop(0)  # Drop the first row
        df = df.reset_index(drop=True)  # Reset index for cleanliness
    return df

# Camelot table extraction function
def extract_with_camelot(pdf_file, pages="all"):
    """
    Extracts tables from the PDF using Camelot.
    Pages can be 'all' or a specific page range.
    """
    tables = camelot.read_pdf(pdf_file, pages=pages, flavor='stream')
    data_frames = [clean_table_data(table.df) for table in tables]
    
    if data_frames:
        return pd.concat(data_frames, ignore_index=True)
    return None

# Tesseract OCR extraction function
def extract_with_tesseract(pdf_file, pages="all"):
    """
    Converts each page to image and uses OCR to extract text from images.
    Pages can be 'all' or a specific page range.
    """
    images = convert_from_path(pdf_file, first_page=pages[0], last_page=pages[-1] if isinstance(pages, list) else None)
    text_data = []
    
    for image in images:
        text = tesserocr.image_to_text(image)
        text_data.extend(text.split("\n"))
    
    return pd.DataFrame(text_data, columns=["Text"])

# pdfplumber text extraction function
def extract_with_pdfplumber(pdf_file, pages="all"):
    """
    Extracts generic text from the PDF using pdfplumber.
    Pages can be 'all' or a specific page range.
    """
    with pdfplumber.open(pdf_file) as pdf:
        all_data = []
        
        # If specific page range is provided
        page_indices = [int(p) - 1 for p in pages] if isinstance(pages, list) else range(len(pdf.pages))
        
        for page_index in page_indices:
            page = pdf.pages[page_index]
            text = page.extract_text()
            if text:
                lines = text.split("\n")
                all_data.extend(lines)
                
        return pd.DataFrame(all_data, columns=["Text"])

# Main function to convert PDF to Excel
def pdf_to_excel(pdf_file, output_excel, mode="tables", page_range="all"):
    """
    Converts the PDF to Excel based on the specified mode.
    Mode can be 'tables', 'text', or 'ocr'.
    Page range can be 'all' or a list of specific pages.
    """
    pages = page_range if page_range != "all" else None

    # Handle table extraction mode
    if mode == "tables":
        try:
            table_df = extract_with_camelot(pdf_file, pages)
            if table_df is not None:
                table_df.to_excel(output_excel, index=False)
                return output_excel
        except Exception as e:
            print(f"Camelot failed: {e}")
    
    # Handle OCR extraction mode
    if mode == "ocr":
        try:
            ocr_df = extract_with_tesseract(pdf_file, pages)
            if not ocr_df.empty:
                ocr_df.to_excel(output_excel, index=False)
                return output_excel
        except Exception as e:
            print(f"Tesseract failed: {e}")

    # Handle text extraction mode
    if mode == "text":
        try:
            text_df = extract_with_pdfplumber(pdf_file, pages)
            if not text_df.empty:
                text_df.to_excel(output_excel, index=False)
                return output_excel
        except Exception as e:
            print(f"pdfplumber failed: {e}")

    return None
