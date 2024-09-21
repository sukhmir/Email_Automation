import streamlit as st
import os
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from extract_and_format_emails import extract_text_from_docx, extract_emails_subjects_bodies

SCOPES = ['https://www.googleapis.com/auth/gmail.send']
gmail_service = None

def authenticate_gmail_service(sender_email):
    """Authenticate and get the Gmail service instance only once."""
    global gmail_service
    creds = None
    token_file = f'token_{sender_email}.json'

    # Check if token file exists and load the credentials from it
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)

    # If credentials are not available or are invalid, handle the authentication
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret_967285117566-b7671q8ftnpli46r68nqtq9e5m507fr7.apps.googleusercontent.com.json', SCOPES)
            creds = flow.run_console()


        # Save the credentials for future runs
        with open(token_file, 'w') as token:
            token.write(creds.to_json())

    # Initialize the Gmail service using the loaded or newly created credentials
    gmail_service = build('gmail', 'v1', credentials=creds)



def send_email(sender_email, recipient_email, subject, body, attachment=None):
    global gmail_service
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = recipient_email
    message['Subject'] = subject
    message.attach(MIMEText(body, 'plain'))

    if attachment:
        try:
            mime_base = MIMEBase('application', 'octet-stream')
            mime_base.set_payload(attachment.read())
            encoders.encode_base64(mime_base)
            mime_base.add_header('Content-Disposition', f'attachment; filename={attachment.name}')
            message.attach(mime_base)
        except Exception as e:
            st.error(f"Failed to attach file {attachment.name}: {e}")

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    try:
        message = {'raw': raw_message}
        send_message = gmail_service.users().messages().send(userId="me", body=message).execute()
        st.success(f"Email sent to {recipient_email} with message ID: {send_message['id']}")
    except Exception as e:
        st.error(f"An error occurred: {e}")

def send_emails_automatically(email_data, sender_email, attachment=None):
    for entry in email_data:
        recipient_email = entry['email']
        subject = entry['subject']
        body = entry['body']
        send_email(sender_email, recipient_email, subject, body, attachment)

    token_file = f'token_{sender_email}.json'
    if os.path.exists(token_file):
        os.remove(token_file)

# Streamlit App
st.title("Email Automation App")
sender_email = "mark.rothman.coach@gmail.com"
st.write(f"Sender Email: {sender_email}")

word_file = st.file_uploader("Upload Word file containing email data", type=["docx"])
pdf_file = st.file_uploader("Upload PDF Attachment", type=["pdf"])

if word_file is not None:
    text = extract_text_from_docx(word_file)
    email_data = extract_emails_subjects_bodies(text)

if st.button("Send Emails"):
    if word_file is not None and email_data:
        attachment = pdf_file  # Use the uploaded file object directly
        authenticate_gmail_service(sender_email)
        send_emails_automatically(email_data, sender_email, attachment)
    else:
        st.error("Please upload a valid Word file with email data.")
