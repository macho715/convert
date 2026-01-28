#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import streamlit as st


st.set_page_config(
    page_title="Email Search Dashboard",
    page_icon="E",
    layout="wide",
    initial_sidebar_state="expanded",
)

from pathlib import Path
css_path = Path(__file__).parent / "style.css"
if css_path.exists():
    st.markdown(f"<style>{css_path.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)

pages = [
    st.Page("pages/01_overview.py", title="Overview", icon=":material/home:"),
    st.Page("pages/02_search.py", title="Search", icon=":material/search:"),
    st.Page("pages/03_thread_explorer.py", title="Thread Explorer", icon=":material/account_tree:"),
    st.Page("pages/04_quality_audit.py", title="Quality / Audit", icon=":material/check_circle:"),
]

nav = st.navigation(pages, position="sidebar")
nav.run()
