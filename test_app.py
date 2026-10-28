import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Alignment

st.set_page_config(page_title="Enterprise QA Generator - Global E2E Suite", layout="wide")

st.title("🧪 Enterprise E2E Test Suite Generator (Global Login & Flexible Pre-Conditions)")
st.markdown("Generates comprehensive, non-redundant test cases with **globalized pre-condition login (Biometrics/PIN)** plus dynamic pre-conditions, starting directly from the **KITS homepage icons** through complete workflows (Journey -> Products -> Passenger -> Details -> Review -> Payment).")

st.sidebar.header("⚙️ Generator Settings")

test_category = st.sidebar.selectbox(
    "Target Domain / Feature Scope:",
    [
        "ERL Complete Functional Suite (Ticketing, Gates, Payments, Errors)", 
        "Deals & Rewards (Direct Purchase via Wallet, Cards & Touch 'n Go)",
        "ETS & KTMB Passenger Journey", 
        "All Systems"
    ]
)

test_target_count = st.sidebar.slider(
    "Max Test Cases to Generate:",
    min_value=50,
    max_value=1000,
    value=200,
    step=50
)

uploaded_file = st.file_uploader(
    "Upload APK, Build File, Specification PDF, or Spreadsheet",
    type=["apk", "pdf", "docx", "xlsx", "csv"]
)

if uploaded_file is not None:
    st.info(f"File loaded successfully: **{uploaded_file.name}**")

if st.button("🚀 Generate Master E2E Test Suite", type="primary"):
    with st.spinner("Generating unique E2E test suite with globalized login and modular pre-conditions..."):
        
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
        products_types = [
            ("Standard Express Ticket", "Single Trip"),
            ("Return Pass / Flexi Ticket", "Round Trip")
        ]
        passenger_types = [
            ("Adult (Standard)", "1 Passenger"),
            ("Child / Senior", "1 Concession")
        ]
        payment_scenarios = [
            ("Secure Checkout (Pass)", "Payment authorized successfully via 3D-Secure OTP; digital ticket issued with scannable QR Code & barcode rendered clearly with auto screen brightness boost.", "Active internet connection & saved payment card linked."),
            ("Payment Failure (Declined Card)", "App displays 'Payment failed: Card declined' error message, keeps cart intact, and permits payment retry.", "Active internet connection & test card configured to return 'Declined' response."),
            ("Payment Hold / Pending", "App handles pending gateway state gracefully, displays 'Transaction Processing', and polls status without double-charging.", "Active internet connection & simulated gateway latency (>30s)."),
            ("Network Error / Timeout (>180s)", "App catches network timeout, displays friendly error alert, and safely restores user cart session.", "Unstable network connection or simulated gateway timeout.")
        ]

        fixed_date = "2026-11-15 (10:30 AM)"
        generated_suite = []
        tc_counter = 1

        for orig_name, orig_code in origins:
            for dest_name, dest_code in destinations:
                if orig_name == dest_name:
                    continue
                for prod_name, prod_desc in products_types:
                    for pass_name, pass_desc in passenger_types:
                        for pay_name, pay_expected, pay_specific_pre in payment_scenarios:
                            if tc_counter > test_target_count:
                                break
                            
                            test_id = f"TC_KITS_{tc_counter:03d}"
                            test_name = f"[{orig_code}->{dest_code}] {prod_name} - {pass_name} ({pay_name.split()[0]})"
                            scenario = f"Verify E2E workflow: KITS Homepage -> ERL Icon -> Journey ({orig_name} to {dest_name}) -> Products -> Passenger -> Details -> Review -> Payment ({pay_name})"
                            
                            description = f"Globalized E2E test starting from KITS homepage tapping ERL icon, selecting route {orig_name} to {dest_name} on fixed date {fixed_date}, selecting product {prod_name}, configuring passenger {pass_name}, entering contact details, reviewing itinerary summary, and processing payment under {pay_name}."
                            
                            # Globalized Pre-Condition combining Login + Scenario-specific supplementary pre-conditions
                            pre_condition = f"1. [GLOBAL] User is authenticated and logged into the KITS app using Biometrics or PIN.\n2. [MODULE] {pay_specific_pre}"
                            
                            test_steps = (
                                f"1. On KITS app homepage, locate and tap the **ERL icon**.\n"
                                f"2. **Journey**: Select Origin as {orig_name} [{orig_code}] and Destination as {dest_name} [{dest_code}] for travel date {fixed_date}.\n"
                                f"3. **Products**: Choose '{prod_name}' ({prod_desc}).\n"
                                f"4. **Passenger**: Select '{pass_name}' ({pass_desc}).\n"
                                f"5. **Details**: Enter passenger contact information.\n"
                                f"6. **Review**: Verify complete itinerary and fare breakdown on Review screen.\n"
                                f"7. **Payment**: Select payment option ({pay_name}) and complete checkout."
                            )
                            
                            test_data = f"Origin: {orig_name}[{orig_code}], Dest: {dest_name}[{dest_code}], Date: {fixed_date}, Product: {prod_name}, Passenger: {pass_name}, Payment: {pay_name}"
                            
                            if "Pass" in pay_name:
                                expected_result = f"Review screen displays accurate details for {orig_name} -> {dest_name}. Payment succeeds; order confirmation is generated, and digital QR code / barcode renders clearly with automatic screen brightness boost for gantry scanning."
                                priority = "High"
                            else:
                                expected_result = f"Review screen displays accurate details for {orig_name} -> {dest_name}. Payment encounters {pay_name}; {pay_expected}"
                                priority = "Medium"
                                
                            generated_suite.append({
                                "Test ID": test_id,
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

        df = pd.DataFrame(generated_suite)
        
        # Ensure 100% zero duplicates based on Test Steps & Test Data combinations
        df = df.drop_duplicates(subset=['Test Steps', 'Test Data']).reset_index(drop=True)
        
        st.success(f"Successfully generated {len(df)} unique E2E test cases with globalized login and modular pre-conditions!")
        st.dataframe(df, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='KITS_Global_Suite')

        output.seek(0)
        wb = openpyxl.load_workbook(output)
        ws = wb.active

        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')

        column_widths = {
            'A': 15, 'B': 35, 'C': 45, 'D': 40, 'E': 25,
            'F': 50, 'G': 32, 'H': 45, 'I': 15, 'J': 12, 'K': 12, 'L': 15
        }
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width

        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        st.download_button(
            label=f"📥 Download All {len(df)} Unique Test Cases (.xlsx)",
            data=final_output,
            file_name=f"KITS_Global_Workflow_{len(df)}_Test_Cases.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
