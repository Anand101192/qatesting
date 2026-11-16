import os
import time
import subprocess
import re
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt

# Try importing streamlit (for web app deployment)
try:
    import streamlit as st
    has_streamlit = True
except ImportError:
    has_streamlit = False

# Cross-platform compatibility for msvcrt (Windows only)
try:
    import msvcrt
    has_msvcrt = True
except ImportError:
    has_msvcrt = False

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
