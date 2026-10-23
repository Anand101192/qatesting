import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Alignment

st.set_page_config(page_title="Enterprise QA Generator - Unlimited Edition", layout="wide")

st.title("🧪 Enterprise E2E QA Test Generator (1000+ Functional Suite)")
st.markdown("Generate unlimited test cases covering complete ERL functional workflows or Deals & Rewards purchases.")

# Sidebar Configuration
st.sidebar.header("⚙️ Generator Settings")

# Added Deals & Rewards into your existing dropdown menu
test_category = st.sidebar.selectbox(
    "Target Domain / Feature Scope:",
    [
        "ERL Complete Functional Suite (Ticketing, Gates, Payments, Errors)", 
        "Deals & Rewards (Wallet, Cards & Touch 'n Go)",
        "ETS & KTMB Passenger Journey", 
        "All Systems"
    ]
)

# Number input field allowing up to 2,000 test cases
test_target_count = st.sidebar.number_input(
    "Number of Test Cases to Generate:",
    min_value=10,
    max_value=2000,
    value=1000,
    step=50
)

# File Uploader
uploaded_file = st.file_uploader(
    "Upload APK, Build File, Specification PDF, or Spreadsheet",
    type=["apk", "pdf", "docx", "xlsx", "csv"]
)

if uploaded_file is not None:
    st.info(f"File loaded successfully: **{uploaded_file.name}**")

if st.button("🚀 Generate Master Test Suite"):
    with st.spinner(f"Generating {test_target_count} test cases for {test_category}..."):
        
        origins = [
            ("KL Sentral", "KLS"),
            ("Bandar Tasik Selatan", "BTS"),
            ("Putrajaya & Cyberjaya", "PJC"),
            ("Salak Tinggi", "STG"),
            ("KLIA Terminal 1", "KLT1")
        ]
        
        destinations = [
            ("KLIA Terminal 2", "KLT2"),
            ("KLIA Terminal 1
