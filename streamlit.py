import streamlit as st
import pandas as pd




Invoice = st.Page(
    page="invoice.py",
    title="Input",
    icon="📚",
)

another = st.Page(
    page="another.py",
    title="another"
)



pg = st.navigation({
    "Invoice": [Invoice, another]
})


pg.run()