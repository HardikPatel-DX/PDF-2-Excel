import streamlit as st
from pdf_to_excel import pdf_to_excel
from io import BytesIO

def app():
    st.title("PDF to Excel Converter")

    # Upload PDF file
    uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])
    
    if uploaded_file is not None:
        st.write("Converting PDF to Excel...")
        
        # Convert the PDF to Excel
        output = pdf_to_excel(uploaded_file, "output.xlsx")
        
        # Prepare file for download
        with open(output, "rb") as f:
            bytes_data = BytesIO(f.read())
        
        st.download_button(label="Download Excel", data=bytes_data, file_name="converted.xlsx")

if __name__ == "__main__":
    app()
