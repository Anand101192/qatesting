import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Alignment

st.set_page_config(page_title="Enterprise QA Generator - Unlimited Edition", layout="wide")

st.title("🧪 Enterprise E2E QA Test Generator (1000+ Functional Suite)")
st.markdown("Generate unlimited test cases covering complete ERL functional workflows: Ticketing, QR Gates, Payments (Success/Failure/Timeouts), Schedules, Login, and Error Recovery.")

# Sidebar Configuration
st.sidebar.header("⚙️ Generator Settings")

test_category = st.sidebar.selectbox(
    "Target Domain / Feature Scope:",
    ["ERL Complete Functional Suite (Ticketing, Gates, Payments, Errors)", "ETS & KTMB Passenger Journey", "All Systems"]
)

# Number input field replacing slider to allow 1000+ or any custom count
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

if st.button("🚀 Generate Master ERL Test Suite"):
    with st.spinner(f"Generating {test_target_count} detailed ERL test cases..."):
        
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
        
        # Comprehensive functional scenarios covering ticketing, payments, gates, errors, and accounts
        functional_templates = [
            {
                "module": "ERL Ticketing & Booking",
                "name": "One-Way Ekspres Booking",
                "scenario": "Verify user can complete a full one-way KLIA Ekspres booking workflow from origin to destination",
                "desc": "Validates end-to-end ticket selection, passenger count, fare calculation, and checkout confirmation.",
                "pre": "User is logged into the app with a verified account and saved payment method.",
                "steps": "1. Launch app and tap 'Book Ticket'.\n2. Select service type as 'KLIA Ekspres'.\n3. Select Origin and Destination stations.\n4. Choose travel date, time slot, and passenger type.\n5. Proceed to payment and complete checkout.",
                "expected": "Booking is confirmed successfully, order summary is displayed, and digital QR ticket is generated."
            },
            {
                "module": "ERL Ticketing & Booking",
                "name": "Transit Multi-Station Booking",
                "scenario": "Verify user can book intermediate transit stops on KLIA Transit route",
                "desc": "Validates multi-station fare calculation and route selection for KLIA Transit.",
                "pre": "User is on the ticket booking home screen.",
                "steps": "1. Tap on 'KLIA Transit' tab.\n2. Select intermediate origin and destination stations.\n3. Verify intermediate fare calculation.\n4. Proceed to payment screen.",
                "expected": "Correct fare for intermediate transit route is calculated and ticket is issued with valid QR code."
            },
            {
                "module": "ERL QR Gate & Turnstile",
                "name": "Digital QR Code Gantry Entry",
                "scenario": "Verify active ticket generates scannable QR code for station gantry entry turnstile",
                "desc": "Tests QR ticket rendering on mobile screen for station turnstile scanning.",
                "pre": "User has a valid active ticket in 'My Tickets'.",
                "steps": "1. Navigate to 'My Tickets' and select active ticket.\n2. Tap 'View QR Code' for entry gate.\n3. Verify barcode clarity and auto screen brightness boost.",
                "expected": "QR code renders clearly, barcode is scannable, and screen brightness increases automatically."
            },
            {
                "module": "ERL QR Gate & Turnstile",
                "name": "Offline E-Ticket Barcode Retrieval",
                "scenario": "Verify user can access cached e-ticket QR code when offline at station gates",
                "desc": "Tests offline ticket retrieval feature when mobile data/Wi-Fi is unavailable.",
                "pre": "User has previously downloaded the ticket while online.",
                "steps": "1. Enable Flight Mode on mobile device.\n2. Open app and navigate to 'My Tickets'.\n3. Tap on active ticket to view QR code.",
                "expected": "Cached QR ticket loads successfully from local storage without requiring internet connection."
            },
            {
                "module": "ERL Schedule & Tracker",
                "name": "Live Train Departure Timetable Search",
                "scenario": "Verify user can view upcoming train departures for selected station",
                "desc": "Tests timetable search functionality for upcoming departures.",
                "pre": "User is on the home dashboard.",
                "steps": "1. Tap on 'Schedules / Timetable'.\n2. Select target station.\n3. View list of upcoming train departures and platform info.",
                "expected": "App displays accurate departure times, train numbers, and platform numbers."
            },
            {
                "module": "Payment Gateway",
                "name": "Credit/Debit Card Secure Checkout Success",
                "scenario": "Verify successful payment processing using valid credit card and 3D-Secure OTP",
                "desc": "Tests secure payment gateway integration with credit card entry and OTP verification.",
                "pre": "User is on checkout payment screen with ticket selected.",
                "steps": "1. Select 'Credit/Debit Card' payment method.\n2. Enter valid card number, expiry, CVV, and cardholder name.\n3. Tap 'Pay Now' and complete bank 3D-Secure OTP verification.",
                "expected": "Payment is authorized successfully, order receipt is generated, and ticket is issued."
            },
            {
                "module": "Payment Gateway",
                "name": "Payment Failure on Declined Card",
                "scenario": "Verify app handles declined card payment gracefully without crashing",
                "desc": "Tests negative scenario when bank declines payment transaction due to insufficient funds.",
                "pre": "User is on checkout screen.",
                "steps": "1. Enter test card configured to return 'Declined' response.\n2. Tap 'Pay Now' and submit transaction.",
                "expected": "App displays clear error message 'Payment failed: Card declined' and allows user to retry."
            },
            {
                "module": "Payment Gateway",
                "name": "Payment Gateway Timeout Handling",
                "scenario": "Verify app behavior when payment gateway session times out (> 180 seconds)",
                "desc": "Tests system resilience when banking gateway takes longer than expected to respond.",
                "pre": "User initiates payment checkout.",
                "steps": "1. Tap 'Pay Now'.\n2. Simulate gateway network delay exceeding 180 seconds.",
                "expected": "App catches timeout exception, displays friendly error notification, and restores cart without double-charging."
            },
            {
                "module": "User Account & Auth",
                "name": "User Login with Registered Credentials",
                "scenario": "Verify user can successfully log into the app with valid email and password",
                "desc": "Tests authentication workflow for returning registered users.",
                "pre": "User has a registered account.",
                "steps": "1. Open app and navigate to Login screen.\n2. Enter registered email address and password.\n3. Tap '
