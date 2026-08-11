import streamlit as st

pg = st.navigation([
    st.Page('Universal_Generator_Streamlit.py', title='Sign Document', icon=':pencil:'),
    st.Page('Universal_Verifier_Streamlit.py', title='Verify Document', icon=':white_check_mark:'),
])
pg.run()
