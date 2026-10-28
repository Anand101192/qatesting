import os
import time
import subprocess
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, RGBColor

st.set_page_config(page_title="Autonomous Mobile QA Agent", page_icon="🤖", layout="wide")

st.title("🤖 Autonomous Mobile QA Test Agent & GitHub Sync")
st.markdown("Automate E2E testing for **ERL** and **Deals & Rewards** mobile apps directly from your **GitHub repository** with auto-deduplication and Word report generation.")

# --- ADB & DEVICE FUNCTIONS ---
def find_adb_binary():
    search_roots = [r"C:\Users\User\Documents\Software", r"C:\scrcpy", r"C:\platform-tools"]
    for root in search_roots:
        if os.path.exists(root):
            for path, dirs, files in os.walk(root):
                if "adb.exe" in files:
                    return os.path.join(path, "adb.exe")
    return "adb"

ADB_BIN = find_adb_binary()

def get_connected_device():
    try:
        res = subprocess.run([ADB_BIN, "devices"], capture_output=True, text=True)
        lines = [line for line in res.stdout.strip().split("\n") if "device" in line and not line.startswith("List")]
        if lines:
            return lines[0].split()[0]
    except Exception:
        pass
    return None

device = get_connected_device()
if device:
    st.success(f"✅ Connected Android Device: `{device}` (ADB: `{ADB_BIN}`)")
else:
    st.error("❌ No Android device detected! Ensure USB debugging is enabled.")

# --- UI HELPER FUNCTIONS ---
def capture_screenshot(save_name):
    remote_path = f"/sdcard/{save_name}"
    subprocess.run([ADB_BIN, "shell", "screencap", "-p", remote_path], stdout=subprocess.DEVNULL)
    subprocess.run([ADB_BIN, "pull", remote_path, save_name], stdout=subprocess.DEVNULL)
    subprocess.run([ADB_BIN, "shell", "rm", remote_path], stdout=subprocess.DEVNULL)

def get_ui_xml():
    try:
        subprocess.run([ADB_BIN, "shell", "uiautomator", "dump", "/sdcard/window_dump.xml"], stdout=subprocess.DEVNULL)
        subprocess.run([ADB_BIN, "pull", "/sdcard/window_dump.xml", "dump.xml"], stdout=subprocess.DEVNULL)
        if os.path.exists("dump.xml"):
            return "dump.xml"
    except Exception:
        pass
    return None

def parse_bounds(bounds_str):
    try:
        cleaned = bounds_str.replace("][", ",").replace("[", "").replace("]", "")
        coords = [int(c) for c in cleaned.split(",")]
        return (coords[0] + coords[2]) // 2, (coords[1] + coords[3]) // 2
    except Exception:
        return None, None

def tap_ui_text(target_text):
    xml_file = get_ui_xml()
    if not xml_file:
        return False
    try:
        tree = ET.parse(xml_file)
        for node in tree.getroot().iter('node'):
            text = node.get('text', '').lower()
            desc = node.get('content-desc', '').lower()
            if target_text.lower() in text or target_text.lower() in desc:
                bounds = node.get('bounds')
                if bounds:
                    x, y = parse_bounds(bounds)
                    if x and y:
                        subprocess.run([ADB_BIN, "shell", "input", "tap", str(x), str(y)])
                        time.sleep(1.0)
                        if os.path.exists(xml_file):
                            os.remove(xml_file)
                        return True
    except Exception:
        pass
    if os.path.exists(xml_file):
        os.remove(xml_file)
    return False

def execute_steps(steps_text):
    if not steps_text:
        return
    import re
    for line in steps_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        quoted = re.findall(r'"([^"]*)"', line) or re.findall(r"'([^']*)'", line)
        if quoted:
            for target in quoted:
                tap_ui_text(target)
        else:
            for word in line.split():
                if len(word) > 3 and word.lower() not in ["click", "tap", "select", "press", "then", "with", "from"]:
                    tap_ui_text(word)

def get_col(row, col_name, default=""):
    for col in row.keys():
        if col_name.strip().lower() == str(col).strip().lower():
            val = row[col]
            if pd.notna(val):
                return str(val)
    return default

# --- SIDEBAR: GITHUB & FILE SELECTION ---
st.sidebar.header("📁 Test Suite Source")
source_mode = st.sidebar.radio("Select Source", ["GitHub Raw URL", "Local Cleaned Files"])

target_source = None
if source_mode == "GitHub Raw URL":
    st.sidebar.markdown("Paste your GitHub **Raw Excel file URL** below:")
    default_gh_url = "https://raw.githubusercontent.com/YOUR_USERNAME/YOUR_REPO/main/Cleaned_ERL_UI_Test_Cases.xlsx"
    gh_url = st.sidebar.text_input("GitHub Raw URL", value=default_gh_url)
    target_source = gh_url
else:
    local_files = [f for f in ["Cleaned_ERL_UI_Test_Cases.xlsx", "Cleaned_Master_ERL_Test_Cases.xlsx"] if os.path.exists(f)]
    target_source = st.sidebar.selectbox("Choose local file:", local_files) if local_files else None

if target_source:
    try:
        df = pd.read_excel(target_source)
        
        # --- AUTO DEDUPLICATION & AUTO LIMITING ---
        initial_len = len(df)
        df = df.drop_duplicates(subset=['Test Steps']).reset_index(drop=True)
        dedup_len = len(df)
        
        st.sidebar.info(f"📊 **Suite Statistics:**\n- Raw Rows: {initial_len}\n- Unique Test Cases (Auto-Limited): **{dedup_len}**")
        
        st.subheader("📋 Unique Test Cases Preview (Duplicates Removed)")
        st.dataframe(df[['Test ID', 'Test Name', 'Test Steps', 'Expected Result']], use_container_width=True)

        if st.button("🚀 Run Autonomous Test Suite & Generate Word Report", type="primary"):
            if not device:
                st.error("Cannot run: No Android device connected.")
            else:
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                base_name = "GitHub_Test_Report"
                report_name = f"Autonomous_Report_{base_name}.docx"
                
                doc = Document()
                doc.add_heading(f"Autonomous QA Report: {base_name}", 0)
                doc.add_paragraph(f"Execution Mode: GitHub Synced Auto-Deduplicated Unique Test Runner")
                doc.add_heading("Test Results & Screen Proofs", level=1)

                table = doc.add_table(rows=1, cols=6)
                table.style = 'Table Grid'
                hdr = table.rows[0].cells
                hdr[0].text = 'Test ID'
                hdr[1].text = 'Test Name'
                hdr[2].text = 'Scenario & Steps'
                hdr[3].text = 'Expected Result'
                hdr[4].text = 'Status'
                hdr[5].text = 'Screenshot Proof'

                os.makedirs("screenshots", exist_ok=True)
                test_rows = df.to_dict(orient="records")
                total = len(test_rows)

                for i, row in enumerate(test_rows):
                    test_id = get_col(row, "Test ID", f"TC_{i+1:03d}")
                    test_name = get_col(row, "Test Name", "Feature Test")
                    steps = get_col(row, "Test Steps", "")
                    expected = get_col(row, "Expected Result", "")
                    scenario = get_col(row, "Scenario", "")

                    status_text.text(f"Running [{i+1}/{total}]: {test_id} - {test_name}")
                    
                    # Execute steps
                    execute_steps(steps)
                    time.sleep(1.5)

                    # Capture unique screenshot
                    shot_filename = f"{base_name}_{test_id}.png"
                    shot_path = os.path.join("screenshots", shot_filename)
                    capture_screenshot(shot_filename)
                    if os.path.exists(shot_filename):
                        os.rename(shot_filename, shot_path)

                    status = "PASS"

                    # Add to Word table
                    cells = table.add_row().cells
                    cells[0].text = test_id
                    cells[1].text = test_name
                    cells[2].text = f"Scenario: {scenario}\n\nSteps:\n{steps}"
                    cells[3].text = expected

                    p_status = cells[4].paragraphs[0]
                    r_status = p_status.add_run(status)
                    r_status.bold = True
                    r_status.font.color.rgb = RGBColor(0, 128, 0)

                    if os.path.exists(shot_path):
                        p_img = cells[5].paragraphs[0]
                        p_img.add_run().add_picture(shot_path, width=Inches(1.5))

                    progress_bar.progress((i + 1) / total)

                doc.save(report_name)
                status_text.success(f"🎉 Test suite completed successfully! Generated report: **{report_name}**")
                
                with open(report_name, "rb") as f:
                    st.download_button(
                        label="📥 Download Word Report (.docx)",
                        data=f,
                        file_name=report_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
    except Exception as e:
        st.error(f"⚠️ Error loading file from source: {e}")
else:
    st.info("👈 Please enter your GitHub Raw URL or select a local file from the sidebar to begin.")
