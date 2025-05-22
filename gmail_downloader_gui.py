# gmail_downloader_gui.py
import streamlit as st
import imaplib
import email
from email.header import decode_header
import os
import datetime
import json
from pathlib import Path

# --- Setup ---
import zipfile
import io
CONFIG_FILE = Path("config.json")
st.session_state.setdefault("stop_requested", False)

# --- Load saved credentials if available ---
def load_credentials():
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {"email": "", "password": ""}

# --- Save credentials to config file ---
def save_credentials(email, password):
    with open(CONFIG_FILE, "w") as f:
        json.dump({"email": email, "password": password}, f)

# --- Main App Layout ---
st.set_page_config(layout="wide")
st.title("📥 Gmail Attachment Downloader")
st.caption("Automatically fetch attachments from Gmail or Outlook with filters and controls")

creds = load_credentials()
account_type = st.selectbox("📡 Email Provider", ["Gmail", "Outlook"], index=0, key="account_type")
col1, col2 = st.columns([1, 1])

with col1:
    email_user = st.text_input("📧 Email Address", value=creds.get("email", ""))
    email_pass = st.text_input("🔒 App Password", type="password", value=creds.get("password", ""))
    st.caption("ℹ️ For Outlook accounts, use an App Password if two-factor authentication is enabled.")
    save_creds = st.checkbox("Remember credentials (store locally)")
    mailbox_options = ["inbox", "[Gmail]/Sent Mail", "[Gmail]/All Mail", "[Gmail]/Drafts", "[Gmail]/Starred", "[Gmail]/Important", "[Gmail]/Spam", "[Gmail]/Trash"]
    mailbox = st.selectbox("📂 Folder/Label to search", options=mailbox_options, index=0)
    save_folder = "downloads"

with col2:
    start_date = st.date_input("📅 Start Date", value=datetime.date.today().replace(day=1))
    end_date = st.date_input("📅 End Date", value=datetime.date.today())
    subject_filter = st.text_input("🔍 Subject contains")
    sender_filter = st.text_input("👤 Sender contains")
    body_filters = st.text_input("✉️ Body contains (comma-separated: rsm, cv, resume)")
    
