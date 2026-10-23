import streamlit as st
import pandas as pd
import zipfile
import io

st.set_page_config(page_title="Enterprise QA Generator - ERL Edition", layout="wide")

st.title("🧪 Enterprise E2E QA Test Generator (100+ Test Cases)")
st.markdown("Upload your APK file (`KitsStyleUAT.apk`), PDF, or spec document to automatically scan internal screens and generate 100+ detailed high-level and low-level test cases.")

# Sidebar Configuration
st.sidebar.header("⚙️ Generator Settings")

test_category = st.sidebar.selectbox(
    "Target Domain / Feature Scope:",
    ["ERL (Express Rail Link & KLIA Ekspres/Transit)", "ETS & KTMB Passenger Journey", "All Systems"]
)

test_target_count = st.sidebar.slider(
    "Number of Test Cases to Generate:",
    min_value=20,
    max_value=200,
    value=100,
    step=10
)

st.sidebar.markdown("---")
st.sidebar.markdown("**Included Coverage:**")
st.sidebar.markdown("- High-Level E2E Journeys\n- Low-Level Field Rules\n- Offline Barcode Scanning\n- Payment Gateway Timeouts\n- Error Handling & Edge Cases")

# Helper function to parse screens/activities directly from APK file
def inspect_apk(file_bytes):
    screens = []
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes), 'r') as z:
            for filename in z.namelist():
                if filename.startswith("res/layout/") and filename.endswith(".xml"):
                    clean_name = filename.replace("res/layout/", "").replace(".xml", "")
                    screens.append(clean_name)
    except Exception:
        pass
    return list(set(screens))

# File Uploader
uploaded_file = st.file_uploader(
    "Upload APK, Build File, Specification PDF, or Spreadsheet",
    type=["apk", "pdf", "docx", "xlsx", "csv"]
)

if uploaded_file is not None:
    st.info(f"File loaded: **{uploaded_file.name}** ({round(uploaded_file.size / (1024*1024), 1)} MB)")
    
    # Run APK inspection if an APK is uploaded
    extracted_screens = []
    if uploaded_file.name.endswith(".apk"):
        with st.spinner("Extracting internal layout modules and activities from APK..."):
            extracted_screens = inspect_apk(uploaded_file.getvalue())
            if extracted_screens:
                st.success(f"Successfully extracted {len(extracted_screens)} screen components from APK layout resources!")
                with st.expander("View Extracted Screen Layouts"):
                    st.write(", ".join(extracted_screens[:30]) + ("..." if len(extracted_screens) > 30 else ""))

    if st.button("🚀 Generate 100+ ERL & E2E Test Cases"):
        with st.spinner("Generating 100+ high-level and low-level test cases..."):
            
            modules = [
                "ERL KLIA Ekspres Ticketing",
                "ERL KLIA Transit Inter-City",
                "ERL QR Code Gate Reader & Offline Ticket",
                "ERL Station Schedule & Train Tracker",
                "Payment Gateway & Wallet",
                "User Account & Session Management"
            ]
            
            generated_suite = []
            
            # Base Master Templates cleanly formatted to avoid syntax breaks
            templates = [
                ("High-Level", "E2E KLIA Ekspres One-Way Ticket Purchase", "User logged in with valid card.", "1. Select Origin (KL Sentral) & Destination (KLIA T1/T2).\n2. Choose date & passenger count.\n3. Complete payment.\n4. Access E-Ticket tab.", "Payment succeeds; valid QR code generated; status marked ACTIVE.", "P1 - Critical"),
                ("High-Level", "E2E KLIA Transit Multi-Station Ticket Booking", "App active.", "1. Select origin (Bandar Tasik Selatan) & destination (Putrajaya & Cyberjaya).\
