# dashboard_v2.py
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title='Demand Increase Monitoring Dashboard', layout='wide')

USERS = {
    'anusha': 'Anusha@2026',
    'cdma': 'Cdma@2026'
}

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title('Login')
    u = st.text_input('Username')
    p = st.text_input('Password', type='password')
    if st.button('Login'):
        if u in USERS and USERS[u] == p:
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error('Invalid credentials')
    st.stop()

st.title('Demand Increase Monitoring Dashboard V2')

df = pd.read_excel('master_file_updated.xlsx')

tabs = st.tabs([
    'Executive Summary',
    'Region Analysis',
    'ULB Analysis',
    'Ward Secretariat Analysis',
    'Detailed Reports'
])

for i, t in enumerate(tabs, start=1):
    with t:
        st.subheader(f'Tab {i}')
