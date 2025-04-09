import pandas as pd
import streamlit as st
from num2words import num2words
from docxtpl import DocxTemplate
import pythoncom
import base64
import os
from datetime import datetime

# Initialize pythoncom once at the start
pythoncom.CoInitialize()

# Set page config
st.set_page_config(page_title="Invoice Generator", layout="wide")

# Title
st.title("Invoice Generator")

# Session state initialization
if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame({
        'Description': [''],
        'Amount': ['']
    })
if 'invoice_generated' not in st.session_state:
    st.session_state.invoice_generated = False

# Sidebar for instructions
with st.sidebar:
    st.header("Instructions")
    st.markdown("""
    1. Fill in the invoice details
    2. Add line items in the table
    3. Click 'Generate Invoice'
    4. Preview and download your invoice
    """)

# Main form
with st.form("invoice_form"):
    col1, col2 = st.columns(2)
    
    with col1:
        date = st.date_input("Date", value=datetime.now())
        invoice_number = st.text_input("Invoice Number", value=f"INV-{datetime.now().strftime('%Y%m%d')}-001")
        customer_name = st.text_input("Customer Name", placeholder="Enter customer name")
    
    with col2:
        total = st.text_input("Total Amount", placeholder="0.00")
        # Convert number to words for display
        if total.replace('.', '', 1).isdigit():
            try:
                total_words = num2words(float(total), lang='en').title()
                st.caption(f"**In words:** {total_words} Only")
            except:
                pass
    
    # Editable dataframe for line items
    st.subheader("Line Items")
    edited_df = st.data_editor(
        st.session_state.df,
        key='editor',
        num_rows='dynamic',
        column_config={
            "Description": st.column_config.TextColumn(width="medium"),
            "Amount": st.column_config.NumberColumn(format="%.2f")
        },
        hide_index=True
    )
    
    submit_button = st.form_submit_button("Generate Invoice")

# Invoice generation
if submit_button:
    try:
        # Validate inputs
        if not all([date, invoice_number, customer_name, total]):
            st.error("Please fill in all required fields")
            st.stop()
        
        if edited_df.empty or edited_df['Description'].isna().all() or edited_df['Amount'].isna().all():
            st.error("Please add at least one line item")
            st.stop()
        
        # Prepare data for template
        invoice_list = []
        for index, row in edited_df.dropna().iterrows():
            invoice_list.append({
                str(index+1), row['Description'], f"{float(row['Amount']):.2f}"
            })
        
        # Calculate total from line items if needed
        calculated_total = edited_df['Amount'].sum()
        
        # Load and render template
        doc = DocxTemplate("template.docx")
        doc.render({
            "invoice_list": invoice_list,
            "date": date.strftime("%d %B %Y"),
            "invoice_number": invoice_number,
            "customer_name": customer_name,
            "total": f"{float(total):.2f}",
        })
        
        # Save to temporary file
        temp_docx = "temp_invoice.docx"
        temp_pdf = "temp_invoice.pdf"
        doc.save(temp_docx)
        
        # Convert to PDF (this would work if docx2pdf is properly installed)
        # For Streamlit sharing, we might need to use a different approach
        try:
            from docx2pdf import convert
            convert(temp_docx, temp_pdf)
            st.session_state.invoice_generated = True
            st.session_state.invoice_path = temp_pdf
            st.success("Invoice generated successfully!")
        except Exception as e:
            st.error(f"PDF conversion failed: {e}")
            # Fallback to DOCX download
            st.session_state.invoice_generated = True
            st.session_state.invoice_path = temp_docx
            st.success("Invoice generated (DOCX format)!")
            
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")

# Preview and download
if st.session_state.get('invoice_generated', False):
    st.divider()
    st.subheader("Invoice Preview")
    
    if st.session_state.invoice_path.endswith('.pdf'):
        # Display PDF
        with open(st.session_state.invoice_path, "rb") as f:
            base64_pdf = base64.b64encode(f.read()).decode('utf-8')
            pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800" type="application/pdf"></iframe>'
            st.markdown(pdf_display, unsafe_allow_html=True)
        
        # Download button for PDF
        with open(st.session_state.invoice_path, "rb") as f:
            st.download_button(
                label="Download Invoice (PDF)",
                data=f,
                file_name=f"Invoice_{invoice_number}.pdf",
                mime="application/pdf"
            )
    else:
        # Download button for DOCX
        with open(st.session_state.invoice_path, "rb") as f:
            st.download_button(
                label="Download Invoice (DOCX)",
                data=f,
                file_name=f"Invoice_{invoice_number}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

# Clean up at the end
def cleanup():
    if 'invoice_path' in st.session_state:
        try:
            os.remove(st.session_state.invoice_path)
        except:
            pass

import atexit
atexit.register(cleanup)