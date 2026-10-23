import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Alignment

st.set_page_config(page_title="Enterprise QA Generator", layout="wide")

st.title("🧪 Enterprise E2E QA Test Generator (100+ Test Cases)")
st.markdown("Generate clean, simple, and structured ERL test cases with vertically stacked test steps.")

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

# File Uploader
uploaded_file = st.file_uploader(
    "Upload APK, Build File, Specification PDF, or Spreadsheet",
    type=["apk", "pdf", "docx", "xlsx", "csv"]
)

if uploaded_file is not None:
    st.info(f"File loaded successfully: **{uploaded_file.name}**")

if st.button("🚀 Generate Master Test Suite"):
    with st.spinner(f"Generating {test_target_count} test cases..."):
        
        modules = [
            "ERL Ticketing Module",
            "ERL Schedule & Train Tracker",
            "ERL QR Gate & Barcode Scanner",
            "Payment Gateway & Checkout",
            "User Account Profile",
            "Notification & Alerts"
        ]
        
        generated_suite = []
        
        for i in range(1, test_target_count + 1):
            mod = modules[(i - 1) % len(modules)]
            priority = "High" if i % 3 == 0 else ("Medium" if i % 2 == 0 else "Low")
            
            # Using actual newline characters so steps stack vertically down inside the cell
            test_steps = f"1. Open the app and go to {mod}.\n2. Enter valid test details for case #{i}.\n3. Tap on submit or confirm button."
            
            generated_suite.append({
                "Test ID": f"TC_ERL_{i:03d}",
                "Test Name": f"{mod} Test {i}",
                "Scenario": f"Verify user can successfully complete action in {mod}",
                "Description": f"Checking basic functionality and response for {mod}.",
                "Pre Condition": "User is logged into the app with active internet connection.",
                "Test Steps": test_steps,
                "Test Data": f"Route: KL Sentral to KLIA, ID: user_{i:03d}",
                "Expected Result": "System processes request successfully and shows confirmation message.",
                "Actual Result": "",
                "Status": "Pending",
                "Priority": priority,
                "Remarks": ""
            })
        
        df = pd.DataFrame(generated_suite)
        
        st.subheader(f"📋 Generated {len(df)} Test Cases for {test_category}")
        st.dataframe(df, use_container_width=True)
        
        # Export to Excel with vertical text wrapping and clean column widths
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Master_Test_Suite')
            
        output.seek(0)
        wb = openpyxl.load_workbook(output)
        ws = wb.active
        
        # Apply vertical alignment and wrap text to all cells
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
                
        # Set clean default column widths to prevent horizontal over-stretching
        column_widths = {
            'A': 12,  # Test ID
            'B': 22,  # Test Name
            'C': 30,  # Scenario
            'D': 25,  # Description
            'E': 25,  # Pre Condition
            'F': 35,  # Test Steps (stacked vertically)
            'G': 20,  # Test Data
            'H': 30,  # Expected Result
            'I': 15,  # Actual Result
            'J': 12,  # Status
            'K': 12,  # Priority
            'L': 15   # Remarks
        }
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width
            
        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)
        
        st.download_button(
            label=f"📥 Download All {len(df)} Test Cases (.xlsx)",
            data=final_output,
            file_name="Master_ERL_Test_Cases.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
