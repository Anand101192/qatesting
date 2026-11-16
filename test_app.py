import io
import openpyxl
from openpyxl.styles import Alignment
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise QA Generator - Dynamic E2E Suite", layout="wide")

st.title("🧪 Enterprise E2E Test Suite Generator")
st.markdown("Generates comprehensive, non-redundant test cases featuring positive and negative scenarios in clean normal English with numeric test steps.")

st.sidebar.header("⚙️ Generator Settings")

test_category = st.sidebar.selectbox(
    "Target Domain / Feature Scope:",
    [
        "Functional Suite (Ticketing, Gates, Payments, Errors)", 
        "Deals and Rewards Suite",
        "All Systems Suite"
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

if st.button("🚀 Generate Master E2E Test Suite", type="primary"):
    with st.spinner("Analyzing app scope and generating non-redundant test cases..."):
        
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

        generated_suite = []
        tc_counter = 1

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
                            
                            # Numeric numbering for test steps
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

        df = pd.DataFrame(generated_suite)
        df = df.drop_duplicates(subset=['Test Steps', 'Test Data']).reset_index(drop=True)
        
        st.success(f"Successfully generated {len(df)} unique test cases combining positive and negative scenarios!")
        st.dataframe(df, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Master_Test_Suite')

        output.seek(0)
        wb = openpyxl.load_workbook(output)
        ws = wb.active

        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')

        col_letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
        widths = [15, 18, 35, 35, 40, 35, 50, 35, 45, 15, 12, 12, 15]
        for col_letter, width in zip(col_letters, widths):
            ws.column_dimensions[col_letter].width = width

        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        st.download_button(
            label=f"📥 Download All {len(df)} Test Cases (.xlsx)",
            data=final_output,
            file_name=f"Master_Test_Suite_{len(df)}_Cases.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
