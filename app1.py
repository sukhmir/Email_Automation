import streamlit as st
import smtplib
import os
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO
from dotenv import load_dotenv
from extract_and_format_emails import extract_rich_text_from_docx, extract_emails_subjects_bodies

# Load environment variables
load_dotenv()

def send_email(
    sender_email: str,
    sender_password: str,
    recipient_email: str,
    subject: str,
    body: str,
    attachment: BytesIO = None
) -> bool:
    """
    Sends email ensuring Times New Roman is used in the body.
    """
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        
        # Apply Times New Roman font
        styled_body = f'<div style="font-family:\'Times New Roman\', serif;">{body}</div>'
        msg.attach(MIMEText(styled_body, 'html'))

        if attachment:
            attachment.seek(0)
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename="{attachment.name}"'
            )
            msg.attach(part)

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        
        return True
    except Exception as e:
        st.error(f"Failed to send email to {recipient_email}: {str(e)}")
        return False


def send_emails_automatically(
    email_data: list,
    sender_email: str,
    sender_password: str,
    subject: str,
    attachment: BytesIO = None
) -> None:
    """
    Sends emails to all recipients with progress tracking.
    """
    progress_bar = st.progress(0)
    status_text = st.empty()
    total_emails = len(email_data)
    success_count = 0
    
    for i, entry in enumerate(email_data):
        recipient_email = entry['email']
        body = entry['body'] or "No body content"
        current_subject = subject if subject else entry.get('subject', 'No Subject')
        
        status_text.text(f"Sending email {i+1} of {total_emails} to {recipient_email}...")
        if send_email(
            sender_email,
            sender_password,
            recipient_email,
            current_subject,
            body,
            attachment
        ):
            success_count += 1
            st.success(f"Email sent to {recipient_email}!")
        
        progress_bar.progress((i + 1) / total_emails)
        time.sleep(1)  # Small delay to avoid rate limiting
    
    status_text.text(f"Completed! {success_count} of {total_emails} emails sent successfully.")

def login(username: str, password: str) -> bool:
    """
    Handles login with environment variables.
    """
    return (
        username == os.getenv("LOGIN_USER") and
        password == os.getenv("LOGIN_PASSWORD")
    )

def main() -> None:
    """
    Main Streamlit application with simplified UI.
    """
    st.set_page_config(page_title="Email Automation", layout="wide")
    st.title("📧 Email Automation App")
    
    # Initialize session state
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "email_data" not in st.session_state:
        st.session_state.email_data = []
    if "docx_processed" not in st.session_state:
        st.session_state.docx_processed = False
    
    # Login screen
    if not st.session_state.authenticated:
        st.subheader("🔒 Login")
        col1, col2 = st.columns([1, 2])
        with col1:
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            
            if st.button("Login"):
                if login(username, password):
                    st.session_state.authenticated = True
                    st.success("Logged in successfully!")
                    st.rerun()
                else:
                    st.error("Invalid credentials")
        with col2:
            st.markdown("""
            **📌 Instructions:**
            1. Enter your credentials to login
            2. Upload a DOCX file containing email addresses and content
            3. Optionally attach a file
            4. Send emails
            """)
        return
    
    # Main application
    st.subheader("✉️ Compose and Send Emails")
    
    # Logout button
    if st.sidebar.button("🚪 Logout"):
        st.session_state.authenticated = False
        st.session_state.email_data = []
        st.session_state.docx_processed = False
        st.rerun()
    
    sender_email = os.getenv("SENDER_EMAIL")
    sender_password = os.getenv("SENDER_PASSWORD")
    
    # File upload section
    col1, col2 = st.columns(2)
    with col1:
        docx_file = st.file_uploader(
            "Upload DOCX File",
            type=["docx"],
            help="Upload a DOCX file containing email content"
        )
    with col2:
        attachment = st.file_uploader(
            "Attachment (Optional)",
            type=["pdf", "docx", "xlsx"],
            help="Attach a file to be sent with all emails"
        )
    
    # Process DOCX file
    if docx_file and not st.session_state.docx_processed:
        with st.spinner("Processing DOCX file..."):
            try:
                html_content = extract_rich_text_from_docx(BytesIO(docx_file.read()))
                st.session_state.email_data = extract_emails_subjects_bodies(html_content)
                st.session_state.docx_processed = True
                st.success("DOCX file processed successfully!")
            except Exception as e:
                st.error(f"Error processing DOCX file: {str(e)}")
    
    # Email composition
    default_subject = st.session_state.email_data[0]['subject'] if st.session_state.email_data and st.session_state.email_data[0].get('subject') else ""
    
    subject = st.text_input(
        "📝 Email Subject",
        value=default_subject,
        placeholder="Enter email subject"
    )
    
    # Send emails button
    if st.session_state.email_data:
        if st.button("🚀 Send All Emails", type="primary"):
            if not subject and not all(entry.get('subject') for entry in st.session_state.email_data):
                st.error("Please enter a subject or ensure subjects are included in the DOCX file")
            else:
                with st.spinner("Sending emails..."):
                    send_emails_automatically(
                        st.session_state.email_data,
                        sender_email,
                        sender_password,
                        subject,
                        attachment
                    )

if __name__ == '__main__':
    main()