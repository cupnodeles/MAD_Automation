import streamlit as st

st.set_page_config(page_title="Automation Hub", layout="wide")

PAGE_MAP = {
    "Home": "Main.py",
    "Daily Prod": "pages/1_Daily_Prod.py",
    "Collection Efforts Clean": "pages/2_CE_Clean.py",
    "Collection Efforts Callout": "pages/3_CE_CALL.py",
    "Collection Efforts Skip": "pages/3_CE_SKIP.py",
    "Daily Prods accu": "pages/dailyprod.py",
    "Collection Efforts Combined": "pages/3_CE_COMBINED.py",
    "Compare": "pages/Comp.py",
    "Productivity Efforts Clean": "pages/4_ProdE_Clean.py",
}

st.title("Automation Hub")
st.markdown("---")
st.write("Select a function from the left sidebar to switch tools.")

selected_page = st.selectbox("Choose a page", list(PAGE_MAP.keys()), index=0)
if selected_page != "Home":
    st.switch_page(PAGE_MAP[selected_page])
