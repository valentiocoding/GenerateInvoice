import streamlit as st
from docxtpl import DocxTemplate
import pandas as pd
from datetime import datetime
import tempfile

st.title("INVOICE")
st.divider()

# Initialize session state for invoice items if not already present
if 'invoice_items' not in st.session_state:
    st.session_state.invoice_items = pd.DataFrame({
        'Part Number': [''],
        'Item Name': [''],
        'Brand': [''],
        'Remarks': [''],
        'Qty': [''],
        'UOM': [''],
        'Unit Price': [''],
    })

# Input fields for invoice details
col1, col2 = st.columns(2)
customer = col1.text_input("Customer Name")
invoicenumber = col2.text_input("Invoice Number")
tanggal = col2.date_input("Tanggal")
nopo = col2.text_input("No PO")
address = col1.text_area("Address")

st.divider()

col1, col2 = st.columns(2)
customertuju = col1.text_input("Customer Tuju")
addresstuju = col1.text_area("Address Tuju")

# Editable table for invoice items
edited_df = st.data_editor(
    st.session_state.invoice_items,
    hide_index=True,
    num_rows="dynamic",
    use_container_width=True    
)

discount = st.number_input("Discount", value=None, placeholder="Dalam Percent", step=1)
if discount is None:
    discount = 0

pname = st.text_input("Penanggung Jawab")

# Function to prepare invoice context
def prepare_invoice_context(df, discount_percent):
    df.insert(0, "No", range(1, len(df) + 1))
    df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').fillna(0)
    df['Unit Price'] = pd.to_numeric(df['Unit Price'], errors='coerce').fillna(0)
    df['Total Price'] = df['Qty'] * df['Unit Price']
    

    subtotal = df['Total Price'].sum()
    discount_value = (discount / 100) * subtotal if discount is not None else 0
    total = subtotal - discount_value
    dpp = total * 11 / 12
    vat = dpp * 0.12
    grandtotal = total + vat

    df['Total Price'] = df['Total Price'].apply(lambda x: f"IDR {x:,.0f}")
    df['Unit Price'] = df['Unit Price'].apply(lambda x: f"IDR {x:,.0f}")
    df['Qty'] = df['Qty'].apply(lambda x: f"{x:,.0f}")

    # Format invoice_list as list of tuples
    invoice_list = df[['No', 'Part Number', 'Item Name', 'Brand', 'Remarks', 'Qty', 'UOM', 'Unit Price', 'Total Price']].values.tolist()

    return {
        'customer': customer,
        'invoice_number': invoicenumber,
        'tanggal': tanggal.strftime("%d-%m-%Y"),
        'nopo': nopo,
        'address': address,
        'customertuju': customertuju,
        'addresstuju': addresstuju,
        'invoice_list': invoice_list,
        'total': f"IDR {total:,.0f}",
        'subtotal': f"IDR {subtotal:,.0f}",
        'discount': f"IDR {discount_value:,.0f} ({discount}%)",
        'vat': f"IDR {vat:,.0f}",
        'dpp': f"IDR {dpp:,.0f}",
        'grandtotal': f"IDR {grandtotal:,.0f}",
    }

# Button to generate the invoice
if st.button("Generate Invoice"):
    context = prepare_invoice_context(edited_df, discount)

    # Load and render the DOCX template
    try:
        doc = DocxTemplate("vellin.docx")  # Ensure your template is named 'vellin.docx'
        doc.render(context)

        # Save the document to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            doc.save(tmp.name)
            tmp_path = tmp.name

        # Provide a download button for the generated DOCX file
        with open(tmp_path, "rb") as f:
            st.download_button(
                label="Download Invoice DOCX",
                data=f,
                file_name=f"Invoice_{invoicenumber}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        st.success("Invoice generated successfully!")

    except Exception as e:
        st.error(f"Error generating invoice: {e}")
