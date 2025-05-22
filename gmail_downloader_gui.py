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
col1, col2 = st.columns(2)

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

file_types = st.multiselect("📎 Attachment types to download", [".pdf", ".docx", ".jpg", ".png"], default=[".pdf", ".docx"])
custom_types = st.text_input("➕ Add custom file types (comma-separated: .mp3, .zip)")
if custom_types:
    file_types += [ft.strip() for ft in custom_types.split(",") if ft.strip().startswith(".")]

progress = st.empty()
log_box = st.empty()
status_text = st.empty()
stop_button = st.button("🛑 Stop")

if stop_button:
    st.session_state["stop_requested"] = True

if st.button("🚀 Start Download"):
    st.session_state["stop_requested"] = False
    start_time = datetime.datetime.now()
    log = []

    if save_creds:
        save_credentials(email_user, email_pass)

    os.makedirs(save_folder, exist_ok=True)
    body_keywords = [word.strip().lower() for word in body_filters.split(",") if word.strip()] if body_filters else []

    try:
        imap_server = "imap.gmail.com" if account_type == "Gmail" else "outlook.office365.com"
        imap = imaplib.IMAP4_SSL(imap_server)
        try:
            imap.login(email_user, email_pass)
        except imaplib.IMAP4.error as login_error:
            raise Exception("Login failed. Please check your email/password. If using Outlook, generate an App Password from your account security settings.") from login_error

        status, _ = imap.select(f'"{mailbox}"')
        if status != "OK":
            raise Exception(f"Unable to access folder: {mailbox}. Please check folder name or IMAP label.")

        since = start_date.strftime('%d-%b-%Y')
        before = (end_date + datetime.timedelta(days=1)).strftime('%d-%b-%Y')
        status, messages = imap.search(None, f'(SINCE "{since}" BEFORE "{before}")')

        email_ids = messages[0].split()
        total = len(email_ids)
        log.append(f"🔍 Found {total} emails between {since} and {before} in '{mailbox}'")

        for idx, email_id in enumerate(email_ids, 1):
            if st.session_state["stop_requested"]:
                log.append("🛑 Stopped by user.")
                break

            status_text.text(f"📨 Reading email {idx} of {total}...")
            res, msg_data = imap.fetch(email_id, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])

            subject = msg.get("Subject", "")
            from_email = msg.get("From", "")
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        try:
                            body = part.get_payload(decode=True).decode(errors='ignore')
                        except:
                            continue
                        break
            else:
                body = msg.get_payload(decode=True).decode(errors='ignore')

            if subject_filter and subject_filter.lower() not in subject.lower():
                continue
            if sender_filter and sender_filter.lower() not in from_email.lower():
                continue
            if body_keywords and not any(keyword in body.lower() for keyword in body_keywords):
                continue

            for part in msg.walk():
                if part.get_content_disposition() == "attachment":
                    filename = part.get_filename()
                    if filename:
                        decoded_filename = decode_header(filename)[0][0]
                        if isinstance(decoded_filename, bytes):
                            decoded_filename = decoded_filename.decode()

                        if not any(decoded_filename.lower().endswith(ft.lower()) for ft in file_types):
                            continue

                        email_date = email.utils.parsedate_to_datetime(msg["Date"])
                        month_folder = email_date.strftime("%Y-%m")
                        full_path = os.path.join(save_folder, month_folder)
                        os.makedirs(full_path, exist_ok=True)

                        filepath = os.path.join(full_path, decoded_filename)
                        if os.path.exists(filepath):
                            base, ext = os.path.splitext(decoded_filename)
                            counter = 1
                            while os.path.exists(filepath):
                                filepath = os.path.join(full_path, f"{base}_{counter}{ext}")
                                counter += 1

                        with open(filepath, "wb") as f:
                            f.write(part.get_payload(decode=True))
                        log.append(f"✅ Saved: {decoded_filename}")

            progress.progress(idx / total)
            log_box.text("\n".join(log[-8:]))

        imap.logout()

        # Create zip of saved files
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(save_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, arcname=os.path.relpath(file_path, save_folder))
        zip_buffer.seek(0)

    except Exception as e:
        log.append(f"❌ Error: {str(e)}")

    end_time = datetime.datetime.now()
    duration = (end_time - start_time).seconds
    log.append(f"\n⏱️ Time taken: {duration} seconds")
    log.append(f"📁 Files saved in: {save_folder}")
    log.append("✅ Done.")
    status_text.text("✔️ Complete.")
    st.download_button(label="⬇️ Download Attachments (ZIP)", data=zip_buffer, file_name="attachments.zip", mime="application/zip")
    log_box.text("\n".join(log[-10:]))
