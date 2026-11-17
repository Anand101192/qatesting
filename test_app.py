import io
import openpyxl
from openpyxl.styles import Alignment
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise QA Generator - Dynamic E2E & UI Suite", layout="wide")

st.title("🧪 Enterprise E2E & UI Test Suite Generator")
st.markdown("Generates comprehensive test cases covering both Functional E2E workflows and dedicated UI validation for the ERL system.")

st.sidebar.header("⚙️ Generator Settings")

test_category = st.sidebar.selectbox(
    "Target Test Suite Scope:",
    [
        "ERL Functional E2E Suite (Ticketing, Payments, Errors)", 
        "ERL UI / Visual Inspection Suite (Layouts, Buttons, Typography)",
        "Combined Functional & UI Suite"
    ]
)

test_target_count = st.sidebar.slider(
    "Max Test Cases to Generate:",
    min_value=20,
    max_value=500,
    value=100,
    step=20
)

uploaded_file = st.file_uploader(
    "Upload APK, Build File, Specification PDF, or Spreadsheet",
    type=["apk", "pdf", "docx", "xlsx", "csv"]
)

if st.button("🚀 Generate Test Suite", type="primary"):
    with st.spinner("Analyzing app scope and generating structured test cases..."):
        
        origins = [
            "KL Sentral", 
            "Bandar Tasik Selatan", 
            "Putrajaya and Cyberjaya", 
            "Salak Tinggi", 
            "KLIA Terminal One"
        ]
        destinations = [
            "KLIA Terminal Two", 
            "KLIA Terminal One", 
            "Putrajaya and Cyberjaya", 
            "Bandar Tasik Selatan", 
            "KL Sentral"
        ]
        product_types = [
            ("Standard Express Ticket", "Single Trip"),
            ("Return Flexi Ticket", "Round Trip")
        ]
        passenger_types = [
            ("Adult Passenger", "Standard Fare"),
            ("Concession Passenger", "Discounted Fare")
        ]
        payment_scenarios = [
            (
                "Positive", 
                "Secure Authorized Checkout", 
                "Payment completes successfully and digital ticket renders clearly with automatic screen brightness enhancement for gantry scanning.", 
                "User account is authenticated with active payment linkage."
            ),
            (
                "Negative", 
                "Declined Card Payment Failure", 
                "System displays card decline notification, retains user cart session intact, and allows payment retry.", 
                "User account is authenticated with test card configured for decline response."
            ),
            (
                "Negative", 
                "Gateway Timeout Error", 
                "System catches connection delay, displays friendly alert message, and restores shopping session safely.", 
                "User account is authenticated under simulated network latency conditions."
            )
        ]

        ui_test_scenarios = [
            (
                "Positive",
                "ERL Homepage Icon Alignment and Visual Rendering",
                "Verify that the ERL service icon and title text are properly aligned, formatted cleanly, and responsive across screen dimensions.",
                "1. Open the application and view the primary homepage dashboard.\n2. Inspect the position, padding, and clarity of the ERL feature icon and label text.\n3. Verify visual contrast and touch target dimensions.",
                "The ERL feature icon renders with correct branding proportions, crisp typography, and full responsiveness without overlapping elements."
            ),
            (
                "Positive",
                "Station Dropdown Menu Readability and Typography",
                "Verify that departure and arrival station selection lists display legible text without truncation or overlapping characters.",
                "1. Navigate to the station selection screen.\n2. Open the station dropdown menu.\n3. Review font size, line spacing, and text contrast against the background.",
                "All station options display clearly in normal English text with complete visibility and proper padding."
            ),
            (
                "Positive",
                "Digital Ticket Barcode and QR Code Brightness Boost",
                "Verify that viewing an active ticket automatically triggers maximum screen brightness for gantry optical scanning.",
                "1. Navigate to active tickets and select a valid digital travel pass.\n2. Tap the view code option to display the optical barcode and QR code.\n3. Observe device display settings.",
                "Screen brightness boosts automatically to the maximum level, ensuring clear scannability under turnstile scanners."
            ),
            (
                "Negative",
                "Error Alert Box Text Truncation and Alignment",
                "Verify that negative payment error messages wrap correctly without clipping or overlapping border boundaries on smaller screen displays.",
                "1. Trigger a simulated payment failure or network timeout condition.\n2. Observe the layout and text alignment inside the alert pop-up window.",
                "Error notification text wraps cleanly within container boundaries with appropriate padding and legible font styling."
            )
        ]

        generated_suite = []
        tc_counter = 1

        if "Functional" in test_category or "Combined" in test_category:
            for orig in origins:
                for dest in destinations:
                    if orig == dest:
                        continue
                    for prod_name, prod_desc in product_types:
                        for pass_name, pass_desc in passenger_types:
                            for scen_type, pay_name, pay_expected, pay_pre in payment_scenarios:
                                if tc_counter > test_target_count:
                                    break
                                
                                test_id = f"TC_E2E_{tc_counter:03d}"
                                test_name = f"{orig} to {dest} {prod_name} {pass_name} {pay_name}"
                                scenario = f"{scen_type} validation of {prod_name.lower()} workflow from {orig} to {dest}"
                                
                                description = f"Validating system behavior for {prod_name.lower()} purchase between {orig} and {dest} utilizing {pass_name.lower()} under {pay_name.lower()} conditions."
                                pre_condition = f"User is successfully authenticated and logged into the application. {pay_pre}"
                                
                                test_steps = (
                                    f"1. Navigate to the main homepage and select the target transport service feature icon.\n"
                                    f"2. Choose {orig} as the departure station and {dest} as the arrival station.\n"
                                    f"3. Choose {prod_name} with {pass_desc} classification.\n"
                                    f"4. Proceed to review the itinerary details, select {pay_name}, and complete the checkout action."
                                )
                                
                                test_data = f"Departure Station: {orig}, Arrival Station: {dest}, Product Type: {prod_name}, Passenger Category: {pass_name}, Payment Mode: {pay_name}"
                                expected_result = f"Itinerary summary displays correct route details from {orig} to {dest}. {pay_expected}"
                                priority = "High" if scen_type == "Positive" else "Medium"
                                
                                generated_suite.append({
                                    "Test ID": test_id,
                                    "Suite Type": "Functional E2E",
                                    "Scenario Type": scen_type,
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
                                tc_counter += 1

        if "UI" in test_category or "Combined" in test_category:
            ui_counter = 1
            for scen_type, ui_name, ui_desc, ui_steps, ui_expected in ui_test_scenarios:
                if len(generated_suite) >= test_target_count:
                    break
                
                test_id = f"TC_UI_{ui_counter:03d}"
                generated_suite.append({
                    "Test ID": test_id,
                    "Suite Type": "UI Visual Inspection",
                    "Scenario Type": scen_type,
                    "Test Name": ui_name,
                    "Scenario": ui_desc,
                    "Description": f"Visual interface inspection focusing on {ui_name.lower()} for the ERL ticketing platform.",
                    "Pre Condition": "Application is installed, active, and displayed on target test device screen.",
                    "Test Steps": ui_steps,
                    "Test Data": "Device Screen Resolution and Theme Settings",
                    "Expected Result": ui_expected,
                    "Actual Result": "",
                    "Status": "Pending",
                    "Priority": "Medium",
                    "Remarks": ""
                })
                ui_counter += 1

        df = pd.DataFrame(generated_suite)
        df = df.drop_duplicates(subset=['Test Steps', 'Test Data']).reset_index(drop=True)
        
        # Save generated dataframe to session state so filtering works smoothly
        st.session_state['generated_df'] = df

# --- NEW: Sidebar Dropdown Filter for Specific Test Cases ---
if 'generated_df' in st.session_state and not st.session_state['generated_df'].empty:
    st.sidebar.markdown("---")
    st.sidebar.header("🔍 Filter Test Case View")
    
    df_cached = st.session_state['generated_df']
    
    # Create a list of options for the dropdown (combining Test ID and Test Name)
    test_case_options = ["All Test Cases"] + [f"{row['Test ID']} - {row['Test Name']}" for _, row in df_cached.iterrows()]
    
    selected_filter = st.sidebar.selectbox(
        "Select Specific Test Case to Inspect:",
        test_case_options
    )
    
    # Filter dataframe based on dropdown selection
    if selected_filter == "All Test Cases":
        filtered_df = df_cached
    else:
        selected_id = selected_filter.split(" - ")[0]
        filtered_df = df_cached[df_cached['Test ID'] == selected_id]

    st.success(f"Displaying {len(filtered_df)} test case(s) based on your selection.")
    st.dataframe(filtered_df, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        filtered_df.to_excel(writer, index=False, sheet_name='Master_Test_Suite')

    output.seek(0)
    wb = openpyxl.load_workbook(output)
    ws = wb.active

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')

    col_letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
    widths = [15, 20, 18, 35, 35, 40, 35, 50, 35, 45, 15, 12, 12, 15]
    for col_letter, width in zip(col_letters, widths):
        ws.column_dimensions[col_letter].width = width

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    st.download_button(
        label=f"📥 Download Selected Test Cases (.xlsx)",
        data=final_output,
        file_name=f"ERL_Filtered_Test_Suite.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
