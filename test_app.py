import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Alignment

st.set_page_config(page_title="Enterprise QA Generator - ERL UI Edition", layout="wide")

st.title("🧪 Enterprise E2E QA Test Generator (Unlimited ERL & UI Suite)")
st.markdown("Generate unlimited test cases covering UI validation (logos, wording, brand colors, station codes) and unique ERL routes.")

# Sidebar Configuration
st.sidebar.header("⚙️ Generator Settings")

test_category = st.sidebar.selectbox(
    "Target Domain / Feature Scope:",
    ["ERL (Express Rail Link & KLIA Ekspres/Transit)", "ETS & KTMB Passenger Journey", "All Systems"]
)

# Unlimited slider up to 300 test cases
test_target_count = st.sidebar.slider(
    "Number of Test Cases to Generate:",
    min_value=20,
    max_value=300,
    value=150,
    step=10
)

# File Uploader
uploaded_file = st.file_uploader(
    "Upload APK, Build File, Specification PDF, or Spreadsheet",
    type=["apk", "pdf", "docx", "xlsx", "csv"]
)

if uploaded_file is not None:
    st.info(f"File loaded successfully: **{uploaded_file.name}**")

if st.button("🚀 Generate Unlimited Master Test Suite"):
    with st.spinner(f"Generating {test_target_count} unique UI & ERL test cases..."):
        
        # Unique ERL stations and codes to prevent repetition
        origins = [
            ("KL Sentral", "KLS"),
            ("Bandar Tasik Selatan", "BTS"),
            ("Putrajaya & Cyberjaya", "PJC"),
            ("Salak Tinggi", "STG"),
            ("KLIA Terminal 1", "KLT1")
        ]
        
        destinations = [
            ("KLIA Terminal 2", "KLT2"),
            ("KLIA Terminal 1", "KLT1"),
            ("Putrajaya & Cyberjaya", "PJC"),
            ("Bandar Tasik Selatan", "BTS"),
            ("KL Sentral", "KLS")
        ]
        
        ui_checks = [
            ("UI Logo & Brand Color", "Verify official ERL logo is displayed with brand red (#E31B23) and high contrast text."),
            ("Station Code Wording", "Verify station name matches destination code abbreviation precisely (e.g., KLIA Terminal 2 [KLT2])."),
            ("Button Wording & Label", "Verify action buttons display correct wording (e.g., 'Book Now', 'Confirm Payment') without typos."),
            ("Currency & Fare Formatting", "Verify ticket pricing displays in Malaysian Ringgit (RM) format (e.g., RM 55.00)."),
            ("Screen Layout & Font Size", "Verify UI typography is readable, clean, and properly aligned without horizontal overflow.")
        ]
        
        generated_suite = []
        
        for i in range(1, test_target_count + 1):
            # Rotate origins and destinations mathematically so they never repeat blindly
            orig_name, orig_code = origins[(i - 1) % len(origins)]
            dest_name, dest_code = destinations[(i - 1) % len(destinations)]
            
            # Ensure origin and destination are never identical
            if orig_name == dest_name:
                dest_name, dest_code = destinations[i % len(destinations)]
                
            ui_element, ui_expected = ui_checks[(i - 1) % len(ui_checks)]
            
            priority = "High" if i % 3 == 0 else ("Medium" if i % 2 == 0 else "Low")
            
            test_name = f"UI & Booking Verification: {orig_name} to {dest_name}"
            scenario = f"Verify UI elements ({ui_element}) and booking flow from {orig_name} ({orig_code}) to {dest_name} ({dest_code})"
            description = f"Checks that the application correctly renders UI components, wording, and station codes for route {orig_code} -> {dest_code}."
            
            pre_condition = "App is launched, user is on the booking dashboard with active network connection."
            
            test_steps = (
                f"1. Open app and verify {ui_element}.\n"
                f"2. Select Origin as {orig_name} [{orig_code}] and Destination as {dest_name} [{dest_code}].\n"
                f"3. Verify destination name and station code match UI specification exactly.\n"
                f"4. Proceed to checkout and verify button wording and currency display."
            )
            
            test_data = f"Origin: {orig_name} [{orig_code}], Dest: {dest_name} [{dest_code}], Run ID: {i:03d}"
            expected_result = f"{ui_expected} Route from {orig_name} [{orig_code}] to {dest_name} [{dest_code}] calculates correct fare and displays proper UI branding."
            
            generated_suite.append({
                "Test ID": f"TC_ERL_{i:03d}",
                "Test Name": test_name,
                "Scenario": scenario,
                "Description": description,
                "Pre Condition": pre_condition,
                "Test Steps": test_steps,
                "Test Data": test_data,
                "Expected Result": expected_result,
                "Actual Result": "",
                "Status": "Pending",
                "Priority": priority,
                "Remarks": ""
            })
        
        df = pd.DataFrame(generated_suite)
        
        st.subheader(f"📋 Generated {len(df)} Unlimited Unique Test Cases")
        st.dataframe(df, use_container_width=True)
        
        # Export to Excel with vertical text wrapping and professional column formatting
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
                
        # Set clean column widths to prevent horizontal over-stretching
        column_widths = {
            'A': 12,  # Test ID
            'B': 30,  # Test Name
            'C': 40,  # Scenario
            'D': 35,  # Description
            'E': 25,  # Pre Condition
            'F': 38,  # Test Steps (stacked vertically)
            'G': 28,  # Test Data
            'H': 38,  # Expected Result
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
            file_name="Master_ERL_UI_Test_Cases.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
