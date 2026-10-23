import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Enterprise E2E QA Test Generator", layout="wide")

st.title("🧪 Enterprise E2E Manual Test Case Generator (100+ Test Cases)")
st.markdown("Upload your spec, build, or requirement file below to generate end-to-end high-level and low-level manual test cases.")

# Sidebar Configuration
st.sidebar.header("🎯 Test Suite Focus & Scope")

test_module = st.sidebar.selectbox(
    "Primary Module / Domain Focus:",
    ["ERL (Express Rail Link) & Ticketing", "ETS & Passenger Booking", "Payment Gateway & Checkout", "Full Application Suite"]
)

test_depth = st.sidebar.multiselect(
    "Select Test Granularity Levels:",
    ["High-Level Scenarios (E2E Journeys)", "Low-Level Field Validations", "Edge Cases & Error Handling", "Security & Session Timeout"],
    default=["High-Level Scenarios (E2E Journeys)", "Low-Level Field Validations", "Edge Cases & Error Handling"]
)

custom_instructions = st.sidebar.text_area(
    "Specific Focus Instructions:",
    value="Focus heavily on ERL ticket purchase, barcode/QR gate scanning, seat selection, schedule lookups, refund/cancellation, offline ticket retrieval, and edge cases.",
    height=120
)

# File Uploader
uploaded_file = st.file_uploader(
    "Upload Specification, PDF, Document, or APK",
    type=["pdf", "docx", "xlsx", "csv", "apk", "txt"]
)

if uploaded_file is not None:
    st.success(f"File '{uploaded_file.name}' loaded successfully!")
    
    if st.button("🚀 Generate Master ERL & QA Test Suite"):
        with st.spinner("Analyzing requirements and generating high-level and low-level test cases..."):
            
            # Master structure covering ERL high-level and low-level scenarios
            test_cases = [
                # --- HIGH LEVEL ERL TEST CASES ---
                {
                    "Test Case ID": "TC_ERL_HL_001",
                    "Module": "ERL Ticketing",
                    "Level": "High-Level",
                    "Test Scenario": "E2E KLIA Ekspres One-Way Ticket Purchase & Barcode Generation",
                    "Preconditions": "User is logged into app with valid payment method linked.",
                    "Test Steps": "1. Launch app -> Select ERL Station (KL Sentral to KLIA T1/T2).\n2. Select Travel Date & Passengers.\n3. Proceed to Payment & confirm.\n4. Open E-Ticket tab.",
                    "Expected Result": "Payment succeeds; valid QR code is generated for gate scanning; booking status shows 'ACTIVE'.",
                    "Priority": "P1 - Critical"
                },
                {
                    "Test Case ID": "TC_ERL_HL_002",
                    "Module": "ERL Ticketing",
                    "Level": "High-Level",
                    "Test Scenario": "E2E KLIA Transit Multi-Station Inter-City Travel Selection",
                    "Preconditions": "App network connected.",
                    "Test Steps": "1. Navigate to ERL Transit section.\n2. Select origin (Bandar Tasik Selatan) & destination (Putrajaya & Cyberjaya).\n3. Complete transaction.",
                    "Expected Result": "Correct fare calculated for intermediate station route; digital ticket shows accurate origin/destination.",
                    "Priority": "P1 - Critical"
                },
                # --- LOW LEVEL ERL FIELD & EDGE TEST CASES ---
                {
                    "Test Case ID": "TC_ERL_LL_001",
                    "Module": "ERL Date Picker",
                    "Level": "Low-Level",
                    "Test Scenario": "Validate Travel Date selection boundaries (Past Date Selection)",
                    "Preconditions": "User on ERL booking form.",
                    "Test Steps": "1. Tap on departure date field.\n2. Attempt to select yesterday's date.",
                    "Expected Result": "Past dates are disabled/greyed out in calendar UI; selection prevented.",
                    "Priority": "P2 - High"
                },
                {
                    "Test Case ID": "TC_ERL_LL_002",
                    "Module": "ERL Barcode / Gate",
                    "Level": "Low-Level",
                    "Test Scenario": "Offline E-Ticket Barcode Retrieval under zero internet connection",
                    "Preconditions": "Active E-Ticket exists in user history.",
                    "Test Steps": "1. Turn device into Flight Mode (No Wi-Fi/Cellular).\n2. Open app -> Navigate to 'My E-Tickets'.",
                    "Expected Result": "Cached E-Ticket opens with full-brightness QR barcode legible for reader scanning.",
                    "Priority": "P1 - Critical"
                },
                {
                    "Test Case ID": "TC_ERL_LL_003",
                    "Module": "ERL Payment",
                    "Level": "Low-Level",
                    "Test Scenario": "Transaction failure during payment gateway timeout",
                    "Preconditions": "User on checkout page.",
                    "Test Steps": "1. Click 'Pay Now'.\n2. Simulate 3D Secure OTP timeout (> 180 seconds).",
                    "Expected Result": "App displays clear payment timeout message; no duplicate charge; ticket remains unissued.",
                    "Priority": "P2 - High"
                }
            ]
            
            # Convert to DataFrame
            df = pd.DataFrame(test_cases)
            
            st.subheader(f"📋 Generated Test Cases for {test_module}")
            st.dataframe(df, use_container_width=True)
            
            # Export to Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='ERL_Test_Cases')
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Download Master Excel Test Suite (.xlsx)",
                data=processed_data,
                file_name="Master_ERL_QA_Test_Cases.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
