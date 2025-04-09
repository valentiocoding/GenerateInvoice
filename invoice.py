import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
from docx2pdf import convert
import io
import requests
from datetime import datetime

st.title("Invoice Generator")

# Input fields
date = st.date_input("Date")
invoice_number = st.text_input("Invoice Number")
customer_name = st.text_input("Customer Name")
total = st.text_input("Total Amount")

# Initialize editable dataframe
if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame({
        'Description': [''],
        'Amount': ['']
    })

edited_df = st.data_editor(st.session_state.df, key='editor', num_rows='dynamic')

# Template URL (replace with your actual GitHub raw URL)
TEMPLATE_URL = "https://raw.githubusercontent.com/yourusername/yourrepo/main/template.docx"

if st.button("Generate Invoice"):
    # Download the template from GitHub
    response = requests.get(TEMPLATE_URL)
    template_bytes = io.BytesIO(response.content)
    
    # Create document from template
    doc = DocxTemplate(template_bytes)
    
    # Prepare invoice items
    invoice_list = []
    for index, row in edited_df.iterrows():
        if row['Description'] and row['Amount']:  # Only add non-empty rows
            invoice_list.append([
                str(index+1), 
                row['Description'],
                str(row['Amount'])
            ])
    
    # Render the document
    doc.render({
        "invoice_list": invoice_list,
        "date": date.strftime("%d/%m/%Y"),  # Format date as you prefer
        "invoice_number": invoice_number,
        "customer_name": customer_name,
        "total": total
    })
    
    # Save to memory instead of disk
    doc_bytes = io.BytesIO()
    doc.save(doc_bytes)
    doc_bytes.seek(0)
    
    # Convert to PDF in memory
    pdf_bytes = io.BytesIO()
    
    # This requires docx2pdf which needs COM on Windows - might not work in cloud deployment
    try:
        convert(doc_bytes, pdf_bytes)
        pdf_bytes.seek(0)
        
        # Download button for the PDF
        st.download_button(
            label="Download Invoice (PDF)",
            data=pdf_bytes,
            file_name=f"Invoice_{invoice_number}.pdf",
            mime="application/pdf"
        )
    except Exception as e:
        st.error(f"Failed to convert to PDF: {e}")
        st.info("As fallback, here's the Word document version:")
        st.download_button(
            label="Download Invoice (DOCX)",
            data=doc_bytes,
            file_name=f"Invoice_{invoice_number}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )