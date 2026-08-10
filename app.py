import streamlit as st

st.set_page_config(page_title="Universal Digital Signature", page_icon="✍️", layout="centered")

st.sidebar.title("Navigation")
selection = st.sidebar.radio("Go to", ["Sign Document", "Verify Document"])

if selection == "Sign Document":
    import Universal_Generator_Streamlit
elif selection == "Verify Document":
    import Universal_Verifier_Streamlit
