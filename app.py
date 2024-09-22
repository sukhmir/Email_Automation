import streamlit as st
import os
import base64
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from extract_and_format_emails import extract_text_from_docx, extract_emails_subjects_bodies
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Get sensitive information from .env
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
CLIENT_SECRET_FILE = os.getenv("CLIENT_SECRET_FILE")
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# Global variable for Gmail service
gmail_service = None

def authenticate_gmail_service():
    """Authenticate and get the Gmail service instance only once."""
    global gmail_service
    creds = None
    token_file = f'token_{SENDER_EMAIL}.json'
    
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(
                port=8081, 
                redirect_uri="https://emailautomation.streamlit.app"
             )


        # Save the credentials
        with open(token_file, 'w') as token:
            token.write(creds.to_json())

    # Initialize the Gmail service only once
    gmail_service = build('gmail', 'v1', credentials=creds)

def send_email(recipient_email, subject, body, attachment=None):
    global gmail_service
    message = MIMEMultipart()
    message['From'] = SENDER_EMAIL
    message['To'] = recipient_email
    message['Subject'] = subject
    message.attach(MIMEText(body, 'plain'))

    if attachment:
        try:
            with open(attachment, 'rb') as f:
                mime_base = MIMEBase('application', 'octet-stream')
                mime_base.set_payload(f.read())
                encoders.encode_base64(mime_base)
                mime_base.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment)}')
                message.attach(mime_base)
        except Exception as e:
            st.error(f"Failed to attach file {attachment}: {e}")

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

    try:
        message = {'raw': raw_message}
        send_message = gmail_service.users().messages().send(userId="me", body=message).execute()
        st.success(f"Email sent to {recipient_email} with message ID: {send_message['id']}")
    except Exception as e:
        st.error(f"An error occurred: {e}")

def send_emails_automatically(email_data, attachment=None):
    for entry in email_data:
        recipient_email = entry['email']
        subject = entry['subject']
        body = entry['body']
        send_email(recipient_email, subject, body, attachment)

    # Remove token after sending
    token_file = f'token_{SENDER_EMAIL}.json'
    if os.path.exists(token_file):
        os.remove(token_file)

# Streamlit App
st.title("Email Automation App")

# Display sender email (secured via environment variables)
st.write(f"Sender Email: {SENDER_EMAIL}")

# Upload Word file and PDF attachment
word_file = st.file_uploader("Upload Word file containing email data", type=["docx"])
pdf_file = st.file_uploader("Upload PDF Attachment", type=["pdf"])

if word_file is not None:
    text = extract_text_from_docx(word_file)
    email_data = extract_emails_subjects_bodies(text)

if st.button("Send Emails"):
    if word_file is not None and email_data:
        attachment = pdf_file.name if pdf_file else None

        # Authenticate once before sending all emails
        authenticate_gmail_service()

        # Send emails one by one
        send_emails_automatically(email_data, attachment)
    else:
        st.error("Please upload a valid Word file with email data.")
