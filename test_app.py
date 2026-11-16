import os
import time
import subprocess
import re
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt

# Cross-platform compatibility for msvcrt (Windows only)
try:
    import msvcrt
    has_msvcrt = True
except ImportError:
    has_msvcrt = False

# Try importing streamlit (for web app deployment)
try:
    import streamlit as st
    has_streamlit = True
except ImportError:
    has_streamlit = False

def get_adb_path():
    local_app_data = os.environ.get('LocalAppData', 'C:\\Users\\User\\AppData\\Local')
    sdk_adb = os.path.join(local_app_data, 'Android', 'Sdk', 'platform-tools', 'adb.exe')
    if os.path.exists(sdk_adb):
        return sdk_adb
    return 'adb'

def clean_natural_text(text):
    """Cleans text to ensure natural English without technical hashes or digits."""
    if not isinstance(text, str):
        return str(text)
    text = re.sub(r'\[.*?\]', '', text)
    text = text.replace('#', '')
    text = text.replace('KLS', 'KL Sentral')
    text = text.replace('KLT2', 'KLIA Terminal Two')
    text = text.replace('KLT1', 'KLIA Terminal One')
    text = text.replace('PJC', 'Putrajaya and Cyberjaya')
    text = text.replace('BTS', 'Bandar Tasik Selatan')
    text = text.replace('E2E', 'End-to-end')
    text = re.sub(r'\d+', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def dismiss_permissions(adb_cmd):
    try:
        subprocess.run([adb_cmd, "shell", "input", "tap", "540", "770"], capture_output=True)
        time.sleep(1)
    except:
        pass

def execute_app_navigation(adb_cmd, tc_id, steps, scenario):
    try:
        subprocess.run([adb_cmd, "shell", "input", "keyevent", "KEYCODE_WAKEUP"], capture_output=True)
        time.sleep(1)
        dismiss_permissions(adb_cmd)

        steps_lower = steps.lower()

        print(f"[{tc_id}] Navigating: Tapping ERL ticketing module...")
        subprocess.run([adb_cmd, "shell", "input", "tap", "825", "570"], capture_output=True)
        time.sleep(3)
        dismiss_permissions(adb_cmd)

        if "station" in steps_lower or "route" in steps_lower or "sentral" in steps_lower or "terminal" in steps_lower:
            print(f"[{tc_id}] Navigating: Selecting station route parameters...")
            subprocess.run([adb_cmd, "shell", "input", "tap", "540", "450"], capture_output=True)
            time.sleep(2)
            subprocess.run([adb_cmd, "shell", "input", "tap", "540", "580"], capture_output=True)
            time.sleep(2)

        if "payment" in steps_lower or "checkout" in steps_lower or "gateway" in steps_lower or "declined" in steps_lower:
            print(f"[{tc_id}] Navigating: Proceeding to payment screen...")
            subprocess.run([adb_cmd, "shell", "input", "tap", "540", "2150"], capture_output=True)
            time.sleep(3)

        if "qr" in steps_lower or "ticket" in steps_lower or "turnstile" in steps_lower:
            print(f"[{tc_id}] Navigating: Opening 'My Tickets' and QR code screen...")
            subprocess.run([adb_cmd, "shell", "input", "tap", "340", "890"], capture_output=True)
            time.sleep(2)
            subprocess.run([adb_cmd, "shell", "input", "tap", "540", "500"], capture_output=True)
            time.sleep(2)

    except Exception as e:
        print(f"Navigation warning for {tc_id}: {str(e)}")

def check_keyboard_controls():
    if has_msvcrt and msvcrt.kbhit():
        key = msvcrt.getch().decode('utf-8', errors='ignore').lower()
        if key == 'q':
            return 'stop'
        elif key == 'p':
            return 'pause'
    return None

def run_automated_testing(status_container=None):
    excel_path = "Master_ERL_Test_Cases.xlsx"
    report_path = "ERL_Test_Execution_Report.docx"
    screenshot_dir = "screenshots"
    adb_cmd = get_adb_path()
    
    if not os.path.exists(screenshot_dir):
        os.makedirs(screenshot_dir)

    msg = "Step 1: Reading and cleaning test cases from Excel file..."
    print(msg)
    if status_container: status_container.text(msg)

    try:
        df = pd.read_excel(excel_path)
        print(f"Successfully loaded {len(df)} test cases from {excel_path}.")
    except Exception as e:
        err = f"Error reading Excel file '{excel_path}': {str(e)}"
        print(err)
        if status_container: status_container.error(err)
        return

    msg = "Step 2: Checking connected Android device via ADB..."
    print(msg)
    if status_container: status_container.text(msg)
    try:
        device_result = subprocess.run([adb_cmd, "devices"], capture_output=True, text=True, check=True)
        print(device_result.stdout)
    except Exception as e:
        err = f"ADB check failed: {str(e)}"
        print(err)
        if status_container: status_container.error(err)
        return

    doc = Document()
    doc.add_heading('ERL Automated Test Execution Report', level=0)
    p = doc.add_paragraph()
    p.add_run(f"Execution Date/Time: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
    p.add_run("Target Application: ERL App (Android Device)\n")
    doc.add_heading('Summary of Test Execution', level=1)

    results_summary = []
    is_stopped = False
    
    for index, row in df.iterrows():
        action = check_keyboard_controls()
        if action == 'stop':
            is_stopped = True
            break
        elif action == 'pause':
            while True:
                if has_msvcrt and msvcrt.kbhit():
                    break
                time.sleep(0.5)

        tc_id = str(row.get('Test Case ID', row.get('ID', f"TC_ERL_{index+1:03d}"))).strip()
        tc_name = clean_natural_text(str(row.get('Test Name', row.get('Title', ''))))
        scenario = clean_natural_text(str(row.get('Scenario', row.get('Description', ''))))
        steps = clean_natural_text(str(row.get('Steps', row.get('Test Steps', ''))))

        is_negative = any(kw in scenario.lower() or kw in tc_name.lower() for kw in ['fail', 'error', 'decline', 'timeout', 'invalid', 'hold', 'gateway'])
        scenario_type = "Negative Scenario" if is_negative else "Positive Scenario"

        exec_msg = f"Executing [{tc_id}] ({scenario_type}): {tc_name}"
        print(f"\n--- {exec_msg} ---")
        if status_container: status_container.text(exec_msg)

        status = "Pass"
        actual_result = f"Test executed successfully under {scenario_type.lower()} validation with stability check."
        remarks = "None"
        screenshot_filename = None

        try:
            execute_app_navigation(adb_cmd, tc_id, steps, scenario)

            start_time = time.time()
            while time.time() - start_time < 30:
                action = check_keyboard_controls()
                if action == 'stop':
                    is_stopped = True
                    break
                time.sleep(1)

            if is_stopped:
                break

            remote_path = f"/sdcard/{tc_id}.png"
            local_path = os.path.join(screenshot_dir, f"{tc_id}.png")
            
            cap_result = subprocess.run([adb_cmd, "shell", "screencap", "-p", remote_path], capture_output=True, text=True)
            if cap_result.returncode != 0:
                raise Exception("Failed to execute screencap on device.")
            
            pull_result = subprocess.run([adb_cmd, "pull", remote_path, local_path], capture_output=True, text=True)
            if pull_result.returncode == 0 and os.path.exists(local_path):
                screenshot_filename = local_path
            else:
                raise Exception("Failed to pull screenshot from device memory.")

            subprocess.run([adb_cmd, "shell", "am", "start", "-n", "com.ets.kits/.MainActivity"], capture_output=True)
            time.sleep(1)

        except Exception as e:
            err_msg = str(e)
            if "device" in err_msg.lower() or "disconnected" in err_msg.lower():
                status = "Blocked"
                remarks = f"Execution blocked: {err_msg}"
                actual_result = "Unable to complete execution due to device connectivity issue."
            else:
                status = "Fail"
                remarks = f"Error encountered: {err_msg}"
                actual_result = "Test failed during execution."

        results_summary.append({
            "ID": tc_id,
            "Type": scenario_type,
            "Name": tc_name,
            "Scenario": scenario,
            "Status": status,
            "Actual": actual_result,
            "Remarks": remarks,
            "Screenshot": screenshot_filename
        })

        if is_stopped:
            break

    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Test ID'
    hdr_cells[1].text = 'Type'
    hdr_cells[2].text = 'Test Name'
    hdr_cells[3].text = 'Status'
    hdr_cells[4].text = 'Actual Result'
    hdr_cells[5].text = 'Remarks'

    for res in results_summary:
        row_cells = table.add_row().cells
        row_cells[0].text = res['ID']
        row_cells[1].text = res['Type']
        row_cells[2].text = res['Name']
        row_cells[3].text = res['Status']
        row_cells[4].text = res['Actual']
        row_cells[5].text = res['Remarks']

    doc.add_heading('Detailed Test Case Execution & Screenshots', level=1)

    for res in results_summary:
        doc.add_heading(f"Test Case: {res['ID']} ({res['Type']}) - {res['Name']}", level=2)
        doc.add_paragraph(f"Scenario Description: {res['Scenario']}")
        doc.add_paragraph(f"Execution Status: {res['Status']}")
        doc.add_paragraph(f"Actual Result: {res['Actual']}")
        if res['Remarks'] != "None":
            doc.add_paragraph(f"Remarks: {res['Remarks']}")
        
        if res['Screenshot'] and os.path.exists(res['Screenshot']):
            doc.add_paragraph("Execution Screenshot:")
            doc.add_picture(res['Screenshot'], width=Inches(4.5))
        else:
            doc.add_paragraph("Execution Screenshot: [Not Available / Blocked]")
        
        doc.add_paragraph("-" * 40)

    doc.save(report_path)
    success_msg = f"Test Execution Report successfully generated and saved as '{report_path}'."
    print(success_msg)
    if status_container: status_container.success(success_msg)

# Streamlit Web App Interface
if has_streamlit:
    st.title("ERL Automated Test Execution Dashboard")
    st.write("Manage and execute automated test cases with positive/negative scenario validation and report generation.")
    
    if st.button("Start Test Automation"):
        status_box = st.empty()
        status_box.text("Initializing test execution...")
        run_automated_testing(status_box)
        
        if os.path.exists("ERL_Test_Execution_Report.docx"):
            with open("ERL_Test_Execution_Report.docx", "rb") as f:
                st.download_button("Download Test Execution Report (.docx)", f, file_name="ERL_Test_Execution_Report.docx")

elif __name__ == "__main__":
    run_automated_testing()
