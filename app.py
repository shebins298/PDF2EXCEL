# app.py (Java-free PDF to Word/Excel converter)
import streamlit as st
import pdfplumber
from docx import Document
import pandas as pd
import io

def pdf_to_word(pdf_bytes):
    doc = Document()
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                doc.add_paragraph(text)
    return doc

def pdf_to_excel(pdf_bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        # Extract text
        text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
        
        # Extract tables
        all_tables = []
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                all_tables.extend(tables)
    
    # Create Excel file
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Add text
        pd.DataFrame({'Extracted Text': [text]}).to_excel(writer, sheet_name='Text', index=False)
        
        # Add tables
        for i, table in enumerate(all_tables):
            pd.DataFrame(table).to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
    
    return output

# Streamlit UI
st.title("PDF Converter App 📄➡️📊")
uploaded_file = st.file_uploader("Upload PDF", type="pdf")

if uploaded_file:
    pdf_bytes = uploaded_file.read()
    
    if st.button("Convert to Word"):
        doc = pdf_to_word(pdf_bytes)
        bio = io.BytesIO()
        doc.save(bio)
        st.download_button(
            label="Download Word Document",
            data=bio.getvalue(),
            file_name="converted.docx",
            mime="docx"
        )
    
    if st.button("Convert to Excel"):
        excel_file = pdf_to_excel(pdf_bytes)
        st.download_button(
            label="Download Excel File",
            data=excel_file.getvalue(),
            file_name="converted.xlsx",
            mime="xlsx"
        )
