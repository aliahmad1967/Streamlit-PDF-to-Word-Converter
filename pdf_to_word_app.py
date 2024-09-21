import streamlit as st
import os
import base64
import fitz  # PyMuPDF
from docx import Document
import io

def convert_pdf_to_docx(pdf_file):
    # Get the file name without extension
    file_name = os.path.splitext(pdf_file.name)[0]
    
    # Create a new Word document
    doc = Document()
    
    # Open the PDF
    pdf_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
    
    # Process each page
    for page in pdf_document:
        # Extract text from the page
        text = page.get_text()
        
        # Add the extracted text to the Word document
        doc.add_paragraph(text)
        
        # Add a page break after each page except the last one
        if page.number < len(pdf_document) - 1:
            doc.add_page_break()
    
    # Close the PDF
    pdf_document.close()
    
    # Save the document to a BytesIO object
    docx_file = io.BytesIO()
    doc.save(docx_file)
    docx_file.seek(0)
    
    return docx_file, f"{file_name}.docx"

def get_binary_file_downloader_html(bin_file, file_name, file_label='File'):
    bin_str = base64.b64encode(bin_file.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{bin_str}" download="{file_name}">Download {file_label}</a>'
    return href

def main():
    st.title("PDF to Editable Word Converter")
    
    uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")
    
    if uploaded_file is not None:
        st.write("File uploaded successfully!")
        
        if st.button("Convert to Word"):
            with st.spinner("Converting... This may take a while for large documents."):
                docx_file, file_name = convert_pdf_to_docx(uploaded_file)
            
            st.success("Conversion complete!")
            st.markdown(get_binary_file_downloader_html(docx_file, file_name, 'Word Document'), unsafe_allow_html=True)

if __name__ == "__main__":
    main()