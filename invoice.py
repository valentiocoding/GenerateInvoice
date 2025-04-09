import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
import io
import requests
import tempfile
import os
import pythoncom
from docx2pdf import convert

# Inisialisasi COM untuk konversi PDF
pythoncom.CoInitialize()

st.title("Invoice Generator")

# URL template di GitHub
TEMPLATE_URL = "https://raw.githubusercontent.com/valentiocoding/GenerateInvoice/main/template.docx"

# Input fields dasar
col1, col2 = st.columns(2)
with col1:
    date = st.date_input("Date")
    invoice_number = st.text_input("Invoice Number*")
with col2:
    customer_name = st.text_input("Customer Name*")
    total = st.text_input("Total Amount*")

# Tabel item invoice
if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame({
        'Description': [''],
        'Amount': ['']
    })

edited_df = st.data_editor(
    st.session_state.df,
    key='editor',
    num_rows='dynamic',
    column_config={
        "Description": "Item Description",
        "Amount": st.column_config.NumberColumn(
            "Amount",
            format="%.2f"
        )
    }
)

if st.button("Generate Invoice"):
    # Validasi input
    errors = []
    if not invoice_number.strip():
        errors.append("Invoice Number is required")
    if not customer_name.strip():
        errors.append("Customer Name is required")
    if not total.strip():
        errors.append("Total Amount is required")
    
    # Validasi item invoice
    invoice_list = []
    for index, row in edited_df.iterrows():
        if pd.notna(row['Description']) and pd.notna(row['Amount']):
            desc = str(row['Description']).strip()
            amount = str(row['Amount']).strip()
            if desc and amount:
                try:
                    float(amount)
                    invoice_list.append([str(index+1), desc, amount])
                except ValueError:
                    errors.append(f"Row {index+1}: Amount must be a number")
    
    if not invoice_list:
        errors.append("Please add at least one invoice item")
    
    if errors:
        for error in errors:
            st.error(error)
    else:
        try:
            # Download template dari GitHub
            response = requests.get(TEMPLATE_URL)
            response.raise_for_status()
            
            # Buat temporary directory
            with tempfile.TemporaryDirectory() as tmpdir:
                docx_path = os.path.join(tmpdir, f"invoice_{invoice_number}.docx")
                pdf_path = os.path.join(tmpdir, f"invoice_{invoice_number}.pdf")
                
                # Render template
                doc = DocxTemplate(io.BytesIO(response.content))
                doc.render({
                    "invoice_list": invoice_list,
                    "date": date.strftime("%d/%m/%Y"),
                    "invoice_number": invoice_number,
                    "customer_name": customer_name,
                    "total": total
                })
                doc.save(docx_path)
                
                # Baca file DOCX
                with open(docx_path, "rb") as docx_file:
                    docx_bytes = docx_file.read()
                
                # Konversi ke PDF
                try:
                    convert(docx_path, pdf_path)
                    
                    # Baca file PDF
                    with open(pdf_path, "rb") as pdf_file:
                        pdf_bytes = pdf_file.read()
                    
                    st.success("Invoice generated successfully!")
                    
                    # Tombol download
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            "Download DOCX",
                            docx_bytes,
                            file_name=f"invoice_{invoice_number}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    with col2:
                        st.download_button(
                            "Download PDF",
                            pdf_bytes,
                            file_name=f"invoice_{invoice_number}.pdf",
                            mime="application/pdf"
                        )
                
                except Exception as e:
                    st.warning(f"PDF conversion failed: {str(e)}")
                    st.download_button(
                        "Download DOCX Only",
                        docx_bytes,
                        file_name=f"invoice_{invoice_number}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
        
        except requests.exceptions.RequestException as e:
            st.error(f"Failed to download template: {str(e)}")
        except Exception as e:
            st.error(f"Error generating invoice: {str(e)}")

# Cleanup COM
pythoncom.CoUninitialize()