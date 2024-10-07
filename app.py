import streamlit as st
from io import BytesIO
from pdf_to_excel import pdf_to_excel

def app():
    st.title("Advanced PDF to Excel Converter (Robust)")

    # Upload PDF file
    uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])
    
    # User options for extraction mode
    extraction_mode = st.selectbox(
        "Choose Extraction Mode", 
        ["Tables Only", "Text Only", "OCR (for Scanned PDFs)"]
    )

    # User input for page selection (optional)
    page_range_input = st.text_input("Specify page range (optional)", value="all")
    
    # Convert page range input into list of page numbers or 'all'
    page_range = (
        [int(page.strip()) for page in page_range_input.split(",")] 
        if page_range_input != "all" 
        else "all"
    )

    if uploaded_file is not None:
        st.write(f"Processing PDF in {extraction_mode} mode...")

        # Call the appropriate extraction function based on the user selection
        output = pdf_to_excel(uploaded_file, "output.xlsx", mode=extraction_mode, page_range=page_range)

        if output:
            # Prepare the file for download
            with open(output, "rb") as f:
                bytes_data = BytesIO(f.read())
            
            st.download_button(label="Download Excel", data=bytes_data, file_name="converted.xlsx")
        else:
            st.error("Failed to convert the PDF. Please check the file or try again.")

if __name__ == "__main__":
    app()
