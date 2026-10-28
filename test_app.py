import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Alignment

st.set_page_config(page_title="Enterprise QA Generator - Unlimited Edition", layout="wide")

st.title("🧪 Enterprise E2E QA Test Generator (1000+ Functional Suite)")
st.markdown("Generate unlimited test cases covering complete ERL functional workflows or Deals & Rewards direct money purchases.")

# Sidebar Configuration
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

test_target_count = st.sidebar.number_input(
    "Number of Test Cases to Generate:",
    min_value=10,
    max_value=2000,
    value=1000,
    step=50
)

uploaded_file = st.file_uploader(
    "Upload APK, Build File, Specification PDF, or Spreadsheet",
    type=["apk", "pdf", "docx", "xlsx", "csv"]
)

if uploaded_file is not None:
    st.info(f"File loaded successfully: **{uploaded_file.name}**")

if st.button("🚀 Generate Master Test Suite"):
    with st.spinner(f"Generating {test_target_count} test cases for {test_category}..."):
        
        origins = [("KL Sentral", "KLS"), ("Bandar Tasik Selatan", "BTS"), ("Putrajaya & Cyberjaya", "PJC"), ("Salak Tinggi", "STG"), ("KLIA Terminal 1", "KLT1")]
        destinations = [("KLIA Terminal 2", "KLT2"), ("KLIA Terminal 1", "KLT1"), ("Putrajaya & Cyberjaya", "PJC"), ("Bandar Tasik Selatan", "BTS"), ("KL Sentral", "KLS")]
        partner_brands = [("Zus Coffee", "ZUS"), ("FamilyMart", "FMT"), ("Tealive", "TLV"), ("GrabCar", "GRB"), ("Boost Juice", "BST")]
        
        # 1. ERL Functional Templates
        erl_templates = [
            {
                "module": "ERL Ticketing & Booking", "name": "One-Way Ekspres Booking",
                "scenario": "Verify user can complete a full one-way KLIA Ekspres booking workflow from origin to destination",
                "desc": "Validates end-to-end ticket selection, passenger count, fare calculation, and checkout confirmation.",
                "pre": "User is logged into the app with a verified account and saved payment method.",
                "steps": ["1. Launch app and tap 'Book Ticket'.", "2. Select service type as 'KLIA Ekspres'.", "3. Select Origin and Destination stations.", "4. Choose travel date, time slot, and passenger type.", "5. Proceed to payment and complete checkout."],
                "expected": "Booking is confirmed successfully, order summary is displayed, and digital QR ticket is generated."
            },
            {
                "module": "ERL Ticketing & Booking", "name": "Transit Multi-Station Booking",
                "scenario": "Verify user can book intermediate transit stops on KLIA Transit route",
                "desc": "Validates multi-station fare calculation and route selection for KLIA Transit.",
                "pre": "User is on the ticket booking home screen.",
                "steps": ["1. Tap on 'KLIA Transit' tab.", "2. Select intermediate origin and destination stations.", "3. Verify intermediate fare calculation.", "4. Proceed to payment screen."],
                "expected": "Correct fare for intermediate transit route is calculated and ticket is issued with valid QR code."
            },
            {
                "module": "ERL QR Gate & Turnstile", "name": "Digital QR Code Gantry Entry",
                "scenario": "Verify active ticket generates scannable QR code for station gantry entry turnstile",
                "desc": "Tests QR ticket rendering on mobile screen for station turnstile scanning.",
                "pre": "User has a valid active ticket in 'My Tickets'.",
                "steps": ["1. Navigate to 'My Tickets' and select active ticket.", "2. Tap 'View QR Code' for entry gate.", "3. Verify barcode clarity and auto screen brightness boost."],
                "expected": "QR code renders clearly, barcode is scannable, and screen brightness increases automatically."
            },
            {
                "module": "ERL QR Gate & Turnstile", "name": "Offline E-Ticket Barcode Retrieval",
                "scenario": "Verify user can access cached e-ticket QR code when offline at station gates",
                "desc": "Tests offline ticket retrieval feature when mobile data/Wi-Fi is unavailable.",
                "pre": "User has previously downloaded the ticket while online.",
                "steps": ["1. Enable Flight Mode on mobile device.", "2. Open app and navigate to 'My Tickets'.", "3. Tap on active ticket to view QR code."],
                "expected": "Cached QR ticket loads successfully from local storage without requiring internet connection."
            },
            {
                "module": "ERL Schedule & Tracker", "name": "Live Train Departure Timetable Search",
                "scenario": "Verify user can view upcoming train departures for selected station",
                "desc": "Tests timetable search functionality for upcoming departures.",
                "pre": "User is on the home dashboard.",
                "steps": ["1. Tap on 'Schedules / Timetable'.", "2. Select target station.", "3. View list of upcoming train departures and platform info."],
                "expected": "App displays accurate departure times, train numbers, and platform numbers."
            },
            {
                "module": "Payment Gateway", "name": "Credit/Debit Card Secure Checkout Success",
                "scenario": "Verify successful payment processing using valid credit card and 3D-Secure OTP",
                "desc": "Tests secure payment gateway integration with credit card entry and OTP verification.",
                "pre": "User is on checkout payment screen with ticket selected.",
                "steps": ["1. Select 'Credit/Debit Card' payment method.", "2. Enter valid card number, expiry, CVV, and cardholder name.", "3. Tap 'Pay Now' and complete bank 3D-Secure OTP verification."],
                "expected": "Payment is authorized successfully, order receipt is generated, and ticket is issued."
            },
            {
                "module": "Payment Gateway", "name": "Payment Failure on Declined Card",
                "scenario": "Verify app handles declined card payment gracefully without crashing",
                "desc": "Tests negative scenario when bank declines payment transaction due to insufficient funds.",
                "pre": "User is on checkout screen.",
                "steps": ["1. Enter test card configured to return 'Declined' response.", "2. Tap 'Pay Now' and submit transaction."],
                "expected": "App displays clear error message 'Payment failed: Card declined' and allows user to retry."
            },
            {
                "module": "Payment Gateway", "name": "Payment Gateway Timeout Handling",
                "scenario": "Verify app behavior when payment gateway session times out (> 180 seconds)",
                "desc": "Tests system resilience when banking gateway takes longer than expected to respond.",
                "pre": "User initiates payment checkout.",
                "steps": ["1. Tap 'Pay Now'.", "2. Simulate gateway network delay exceeding 180 seconds."],
                "expected": "App catches timeout exception, displays friendly error notification, and restores cart without double-charging."
            },
            {
                "module": "User Account & Auth", "name": "User Login with Registered Credentials",
                "scenario": "Verify user can successfully log into the app with valid email and password",
                "desc": "Tests authentication workflow for returning registered users.",
                "pre": "User has a registered account.",
                "steps": ["1. Open app and navigate to Login screen.", "2. Enter registered email address and password.", "3. Tap 'Login'."],
                "expected": "User is authenticated successfully and redirected to home dashboard."
            },
            {
                "module": "Error Handling & Recovery", "name": "App Crash Recovery During Payment Processing",
                "scenario": "Verify app syncs transaction state upon relaunch if killed during payment processing",
                "desc": "Tests system state recovery when app process is unexpectedly terminated during payment spinner.",
                "pre": "Payment transaction is in progress.",
                "steps": ["1. Kill app process while processing spinner is active.", "2. Relaunch app and check booking history."],
                "expected": "App syncs with backend; if debited, ticket shows under active bookings; if failed, cart is restored."
            }
        ]

        # 2. Deals & Rewards Money Purchase Templates (In-App Wallet, Saved/New Cards, Touch 'n Go)
        deals_templates = [
            {
                "module": "Deals & Rewards - Direct Purchase", "name": "View Available Deals & Rewards Catalog",
                "scenario": "Verify user can view list of available partner deals with direct monetary pricing (RM)",
                "desc": "Validates that the Deals/Rewards catalog loads active partner offers with accurate Ringgit pricing and category filters.",
                "pre": "User is logged into the mobile application with an active internet connection.",
                "steps": ["1. Launch the app and tap on the 'Deals' or 'Rewards' tab.", "2. Verify featured banners and category filters load with monetary prices (e.g., RM 5.00, RM 10.00).", "3. Tap on a category filter (e.g., 'Dining' or 'Travel').", "4. Verify deal list updates to show matching partner offers."],
                "expected": "Deals catalog loads successfully; category filters work correctly and display relevant partner deals with RM pricing."
            },
            {
                "module": "Deals & Rewards - Direct Purchase", "name": "View Deal Details & Price Breakdown",
                "scenario": "Verify user can view full deal details, original price, discounted purchase price (RM), and validity period",
                "desc": "Tests tapping on a specific deal card to view detailed description, partner locations, and terms of purchase.",
                "pre": "User is on the Deals home screen.",
                "steps": ["1. Tap on any active partner deal card.", "2. Verify deal details screen opens displaying banner, partner name, and monetary purchase price.", "3. Scroll down to check 'Terms & Conditions' and validity period."],
                "expected": "Deal details page opens correctly, showing accurate partner information, validity dates, and readable T&Cs."
            },
            {
                "module": "Deals Checkout - In-App Wallet", "name": "Direct Purchase Using In-App Stored Wallet Balance",
                "scenario": "Verify user can purchase a deal voucher directly using available in-app wallet balance",
                "desc": "Validates direct checkout flow, instant debit of stored wallet funds, and receipt generation.",
                "pre": "User has sufficient in-app wallet balance and is viewing an active deal.",
                "steps": ["1. Tap 'Buy Now' on an active deal detail page.", "2. Select 'In-App Wallet' as the payment method.", "3. Review the price breakdown and tap 'Pay Now'."],
                "expected": "In-app wallet balance is debited correctly, purchase receipt is generated, and voucher appears in 'My Vouchers'."
            },
            {
                "module": "Deals Checkout - In-App Wallet", "name": "Top-Up Prompt on Insufficient In-App Wallet Balance",
                "scenario": "Verify user prompts user to top-up when in-app wallet balance is insufficient for deal purchase",
                "desc": "Tests negative balance handling and quick navigation to wallet top-up screen.",
                "pre": "User has low in-app wallet balance and attempts to buy a deal.",
                "steps": ["1. Select 'In-App Wallet' for a deal purchase exceeding current balance.", "2. Tap 'Pay Now'."],
                "expected": "App displays 'Insufficient balance' warning and provides a direct button to top up wallet funds."
            },
            {
                "module": "Deals Checkout - Saved Cards", "name": "Purchase Deal Using Saved Credit/Debit Card",
                "scenario": "Verify user can checkout a deal using a previously saved credit or debit card",
                "desc
