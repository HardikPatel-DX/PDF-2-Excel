import pdfplumber
import pandas as pd

def pdf_to_excel(pdf_file, output_excel):
    # Open the PDF file
    with pdfplumber.open(pdf_file) as pdf:
        # Initialize a list to store the extracted text
        data = []
        
        # Loop through all pages in the PDF
        for page in pdf.pages:
            # Extract text from each page
            text = page.extract_text()
            if text:
                # Split the text into lines
                lines = text.split("\n")
                # Add each line to the data list
                data.extend(lines)
    
    # Convert the list to a DataFrame
    df = pd.DataFrame(data, columns=["Extracted Text"])
    
    # Save the DataFrame to Excel
    df.to_excel(output_excel, index=False)
    return output_excel
