import pdfplumber
import pandas as pd

def extract_table_data(page):
    # Try to extract table data from a page
    tables = page.extract_tables()
    if tables:
        # If tables are found, convert them into a DataFrame
        df = pd.DataFrame(tables[0])  # Taking the first table found
        return df
    return None

def extract_text_data(page):
    # Extract non-tabular data from the page
    text = page.extract_text()
    if text:
        lines = text.split("\n")
        # Convert the list of lines into a DataFrame (one line per row)
        df = pd.DataFrame(lines, columns=["Text"])
        return df
    return None

def pdf_to_excel(pdf_file, output_excel):
    # Initialize an empty list to store DataFrames from each page
    all_data = []
    
    # Open the PDF file
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # Try extracting tabular data first
            table_df = extract_table_data(page)
            if table_df is not None:
                all_data.append(table_df)
            else:
                # If no table data, extract text data
                text_df = extract_text_data(page)
                if text_df is not None:
                    all_data.append(text_df)

    # Concatenate all the DataFrames from each page
    final_df = pd.concat(all_data, ignore_index=True)
    
    # Save the DataFrame to an Excel file
    final_df.to_excel(output_excel, index=False)
    return output_excel
