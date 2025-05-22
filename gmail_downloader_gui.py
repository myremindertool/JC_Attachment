# gmail_downloader_gui.py
import streamlit as st
import imaplib
import email
from email.header import decode_header
import os
import datetime
import json
from pathlib import Path
import zipfile
import io

# Setup
CONFIG_FILE = Path("config.json")
st.session_state.setdefault("stop_requested", False)

def load_credentials():
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {"email": "", "password": ""}

def save_credentials(email, password):
    with open(CONFIG_FILE, "w") as f:
        json.dump({"email": email, "password": password}, f)

# Stylish page setup
st.set_page_config(page_title="📥 JC Attachment Downloader", page_icon="📬", layout="wide")
st.markdown("""
    <style>
        body {
            background-color: #f8f9fa;
        }
        .block-container {
            padding: 2rem 2rem 2rem 2rem;
            border-radius: 10px;
            background-color: #ffffff;
            box-shadow: 0px 0px 10px rgba(0,0,0,0.05);
        }
        .stButton>button {
            border-radius: 8px;
            color: white;
            background: linear-gradient(90deg, #2b5876 0%, #4e4376 100%);
            padding: 0.6rem 1.2rem;
        }
        .stDownloadButton>button {
            border-radius: 8px;
            color: white;
            background: linear-gradient(90deg, #11998e 0%, #38ef7d 100%);
            padding: 0.6rem 1.2rem;
        }
        .stCheckbox>label, .stSelectbox>label, .stTextInput>label, .stDateInput>label, .stMultiSelect>label {
            color: #2c3e50;
            font-weight: 600;
            font-size: 1rem;
        }
        input, textarea {
            border-radius: 6px;
            padding: 0.5rem;
            font-size: 1rem;
            background-color: #f7f9fc;
            border: 1px solid #ccc;
            transition: all 0.2s;
        }
        input:focus, textarea:focus {
            border-color: #38ef7d;
            box-shadow: 0 0 5px rgba(56, 239, 125, 0.4);
            background-color: #ffffff;
        }
        .stTextInput>div>div>input, .stDateInput input, .stSelectbox>div>div, .stMultiSelect>div>div, .stTextArea textarea {
            border: 1px solid #ccc;
            border-radius: 8px;
            padding: 0.6rem;
            font-size: 1rem;
            background-color: #f4f6fa;
            color: #333;
        }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
    <div style='background: linear-gradient(to right, #00b09b, #96c93d); padding: 2rem; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); text-align: center;'>
        <h1 style='color: white;'>📥 JC Email Attachment Downloader</h1>
        <p style='color: white;'>Download Gmail/Outlook attachments by date, subject, and file type</p>
    </div>
""", unsafe_allow_html=True)

# Credential inputs
creds = load_credentials()
account_type = st.selectbox("📡 Email Provider", ["Gmail", "Outlook"], index=0)
col1, col2 = st.columns(2)

with col1:
    email_user = st.text_input("📧 Email Address", value=creds.get("email", ""))
    email_pass = st.text_input("🔒 App Password", type="password", value=creds.get("password", ""))
    st.caption("ℹ️ Use App Password if 2FA is enabled")
    save_creds = st.checkbox("Remember credentials")
    mailbox = st.selectbox("📂 Folder/Label to search", ["inbox", "[Gmail]/Sent Mail", "[Gmail]/All Mail"])
    save_folder = "downloads"

with col2:
    start_date = st.date_input("📅 Start Date", value=datetime.date.today().replace(day=1))
    end_date = st.date_input("📅 End Date", value=datetime.date.today())
    subject_filter = st.text_input("🔍 Subject contains")
    sender_filter = st.text_input("👤 Sender contains")
    body_filters = st.text_input("✉️ Body keywords (comma-separated: resume, cv, job)")
    file_types = st.multiselect("📎 Attachment types", [".pdf", ".docx", ".jpg", ".png"], default=[".pdf", ".docx"])
    custom_types = st.text_input("➕ Additional file types (.zip, .mp3)")
    if custom_types:
        file_types += [ft.strip() for ft in custom_types.split(",") if ft.strip().startswith(".")]

# Execution controls
st.markdown("---")
col_stop, col_start = st.columns([1, 6])
with col_stop:
    stop_button = st.button("🛑 Stop")
with col_start:
    start_button = st.button("🚀 Start Download")

status_text = st.empty()
progress = st.empty()
log_box = st.empty()

# Main download logic
if stop_button:
    st.session_state["stop_requested"] = True

if start_button:
    st.session_state["stop_requested"] = False
    start_time = datetime.datetime.now()
    log = []

    if save_creds:
        save_credentials(email_user, email_pass)

    os.makedirs(save_folder, exist_ok=True)
    keywords = [k.strip().lower() for k in body_filters.split(",") if k.strip()]

    try:
        imap_server = "imap.gmail.com" if account_type == "Gmail" else "outlook.office365.com"
        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(email_user, email_pass)

        status, _ = imap.select(f'"{mailbox}"')
        if status != "OK":
            raise Exception(f"Cannot access folder: {mailbox}")

        since = start_date.strftime('%d-%b-%Y')
        before = (end_date + datetime.timedelta(days=1)).strftime('%d-%b-%Y')
        _, messages = imap.search(None, f'(SINCE "{since}" BEFORE "{before}")')

        email_ids = messages[0].split()
        total = len(email_ids)
        log.append(f"📬 Found {total} emails in '{mailbox}'")

        for idx, eid in enumerate(email_ids, 1):
            if st.session_state["stop_requested"]:
                log.append("🛑 Stopped by user")
                break

            status_text.text(f"📨 Reading email {idx} of {total}...")
            _, msg_data = imap.fetch(eid, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])

            subject = msg.get("Subject", "")
            from_email = msg.get("From", "")
            body = ""

            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode(errors="ignore")
                        break
            else:
                body = msg.get_payload(decode=True).decode(errors="ignore")

            if subject_filter and subject_filter.lower() not in subject.lower():
                continue
            if sender_filter and sender_filter.lower() not in from_email.lower():
                continue
            if keywords and not any(k in body.lower() for k in keywords):
                continue

            for part in msg.walk():
                if part.get_content_disposition() == "attachment":
                    filename = part.get_filename()
                    if filename:
                        decoded = decode_header(filename)[0][0]
                        if isinstance(decoded, bytes):
                            decoded = decoded.decode()

                        if not any(decoded.lower().endswith(ft) for ft in file_types):
                            continue

                        email_date = email.utils.parsedate_to_datetime(msg["Date"])
                        folder = os.path.join(save_folder, email_date.strftime("%Y-%m"))
                        os.makedirs(folder, exist_ok=True)

                        filepath = os.path.join(folder, decoded)
                        if os.path.exists(filepath):
                            base, ext = os.path.splitext(decoded)
                            counter = 1
                            while os.path.exists(filepath):
                                filepath = os.path.join(folder, f"{base}_{counter}{ext}")
                                counter += 1

                        with open(filepath, "wb") as f:
                            f.write(part.get_payload(decode=True))
                        log.append(f"✅ Saved: {decoded}")

            progress.progress(idx / total)
            log_box.markdown(f"<pre style='background:#f4f4f4;padding:1em;border-radius:6px;font-size:0.9em'>{chr(10).join(log[-8:])}</pre>", unsafe_allow_html=True)

        imap.logout()

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, _, files in os.walk(save_folder):
                for f in files:
                    full = os.path.join(root, f)
                    zf.write(full, arcname=os.path.relpath(full, save_folder))
        zip_buffer.seek(0)

    except Exception as e:
        log.append(f"❌ Error: {str(e)}")

    duration = (datetime.datetime.now() - start_time).seconds
    log.append(f"\n⏱️ Time: {duration} seconds")
    log.append("✅ Done.")
    status_text.text("✔️ Complete.")
    log_box.markdown(f"<pre style='background:#f4f4f4;padding:1em;border-radius:6px;font-size:0.9em'>{chr(10).join(log[-10:])}</pre>", unsafe_allow_html=True)
    st.download_button("⬇️ Download ZIP", zip_buffer, "attachments.zip", "application/zip")
