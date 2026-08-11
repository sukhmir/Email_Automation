import streamlit as st
import smtplib
import os
import time
import hashlib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO
from dotenv import load_dotenv
from extract_and_format_emails import extract_rich_text_from_docx, extract_emails_subjects_bodies

load_dotenv()


def get_secret(key: str, default: str = None) -> str:
    """Prefer Streamlit secrets (Cloud), then .env (local)."""
    try:
        if key in st.secrets:
            return str(st.secrets[key]).strip()
    except Exception:
        pass
    value = os.getenv(key, default)
    return value.strip() if isinstance(value, str) else value


def send_email(
    sender_email: str,
    sender_password: str,
    recipient_email: str,
    subject: str,
    body: str,
    attachment: BytesIO = None,
) -> bool:
    try:
        sender_password = sender_password.replace(" ", "")

        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = recipient_email
        msg["Subject"] = subject

        styled_body = f'<div style="font-family:\'Times New Roman\', serif;">{body}</div>'
        msg.attach(MIMEText(styled_body, "html"))

        if attachment:
            attachment.seek(0)
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f'attachment; filename="{attachment.name}"',
            )
            msg.attach(part)

        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)

        return True
    except smtplib.SMTPAuthenticationError:
        st.error(
            "Email login failed. Use your Gmail address and a Gmail App Password "
            "(not your normal Google password)."
        )
        return False
    except Exception as e:
        st.error(f"Failed to send email to {recipient_email}: {str(e)}")
        return False


def send_emails_automatically(
    email_data: list,
    sender_email: str,
    sender_password: str,
    subject: str,
    attachment: BytesIO = None,
) -> None:
    progress_bar = st.progress(0)
    status_text = st.empty()
    total_emails = len(email_data)
    success_count = 0

    for i, entry in enumerate(email_data):
        recipient_email = entry["email"]
        body = entry["body"] or "No body content"
        current_subject = subject if subject else entry.get("subject", "No Subject")

        status_text.text(f"Sending email {i + 1} of {total_emails} to {recipient_email}...")
        if send_email(
            sender_email,
            sender_password,
            recipient_email,
            current_subject,
            body,
            attachment,
        ):
            success_count += 1
            st.success(f"Email sent to {recipient_email}!")

        progress_bar.progress((i + 1) / total_emails)
        time.sleep(1)

    status_text.text(f"Completed! {success_count} of {total_emails} emails sent successfully.")


def check_login(username: str, password: str) -> bool:
    expected_user = get_secret("LOGIN_USER")
    expected_password = get_secret("LOGIN_PASSWORD")

    if not expected_user or not expected_password:
        st.error(
            "Login is not configured. Add LOGIN_USER and LOGIN_PASSWORD "
            "in Streamlit Secrets (Cloud) or .env (local)."
        )
        return False

    return username.strip() == expected_user and password == expected_password


def main() -> None:
    st.set_page_config(page_title="Email Automation", layout="wide")
    st.title("Email Automation App")

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "email_data" not in st.session_state:
        st.session_state.email_data = []
    if "loaded_docx_key" not in st.session_state:
        st.session_state.loaded_docx_key = None

    # ---- Simple login (username + password only) ----
    if not st.session_state.authenticated:
        st.subheader("Login")
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")

        if st.button("Login"):
            if check_login(username, password):
                st.session_state.authenticated = True
                st.success("Logged in successfully!")
                st.rerun()
            else:
                st.error("Invalid credentials")
        return

    # ---- Main app ----
    st.subheader("Compose and Send Emails")

    if st.sidebar.button("Logout"):
        st.session_state.authenticated = False
        st.session_state.email_data = []
        st.session_state.loaded_docx_key = None
        st.rerun()

    st.markdown("**Sender account (used to send emails)**")
    sender_email = st.text_input(
        "Sender Email",
        value=get_secret("SENDER_EMAIL") or "",
        placeholder="your_email@gmail.com",
    )
    sender_password = st.text_input(
        "Sender Password (Gmail App Password)",
        value=get_secret("SENDER_PASSWORD") or "",
        type="password",
        placeholder="16-character app password",
    )

    col1, col2 = st.columns(2)
    with col1:
        docx_file = st.file_uploader("Upload DOCX File", type=["docx"])
    with col2:
        attachment = st.file_uploader(
            "Attachment (Optional)",
            type=["pdf", "docx", "xlsx"],
        )

    # Always extract recipients from the uploaded file CONTENT (bytes), never by filename
    if docx_file is None:
        st.session_state.email_data = []
        st.session_state.loaded_docx_key = None
    else:
        file_bytes = docx_file.getvalue()
        content_hash = hashlib.sha256(file_bytes).hexdigest()

        if content_hash != st.session_state.loaded_docx_key:
            with st.spinner("Reading emails from uploaded file content..."):
                try:
                    html_content = extract_rich_text_from_docx(BytesIO(file_bytes))
                    parsed = extract_emails_subjects_bodies(html_content)
                    st.session_state.email_data = parsed
                    st.session_state.loaded_docx_key = content_hash
                    st.success(
                        f"Found {len(parsed)} recipient(s) inside the uploaded file content."
                    )
                except Exception as e:
                    st.session_state.email_data = []
                    st.session_state.loaded_docx_key = None
                    st.error(f"Error reading uploaded file content: {str(e)}")

        if st.button("Re-read uploaded file"):
            st.session_state.loaded_docx_key = None
            st.rerun()

    if st.session_state.email_data:
        st.markdown("**Recipients pulled from uploaded file content:**")
        st.write([entry["email"] for entry in st.session_state.email_data])
    elif docx_file is not None:
        st.warning("No email addresses were found in the uploaded file content.")

    default_subject = (
        st.session_state.email_data[0]["subject"]
        if st.session_state.email_data and st.session_state.email_data[0].get("subject")
        else ""
    )

    subject = st.text_input(
        "Email Subject",
        value=default_subject,
        placeholder="Enter email subject",
    )

    if st.session_state.email_data:
        if st.button("Send All Emails", type="primary"):
            if not sender_email or not sender_password:
                st.error("Enter Sender Email and Sender Password before sending.")
            elif not subject and not all(entry.get("subject") for entry in st.session_state.email_data):
                st.error("Please enter a subject or ensure subjects are in the DOCX file.")
            else:
                with st.spinner("Sending emails..."):
                    send_emails_automatically(
                        st.session_state.email_data,
                        sender_email,
                        sender_password,
                        subject,
                        attachment,
                    )


if __name__ == "__main__":
    main()
