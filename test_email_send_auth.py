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

# If modifying the SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def get_gmail_service():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                '', SCOPES)
            creds = flow.run_local_server(port=8081)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)
    return service

def send_email(sender_email, recipient_email, subject, body, attachment=None):
    """Send an email using Gmail API with OAuth2 authentication."""
    service = get_gmail_service()

    # Create the email message
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = recipient_email
    message['Subject'] = subject

    # Attach the email body
    message.attach(MIMEText(body, 'plain'))

    # Attach a file (if provided)
    if attachment:
        try:
            with open(attachment, 'rb') as f:
                mime_base = MIMEBase('application', 'octet-stream')
                mime_base.set_payload(f.read())
                encoders.encode_base64(mime_base)
                mime_base.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment)}')
                message.attach(mime_base)
        except Exception as e:
            print(f"Failed to attach file {attachment}: {e}")

    # Encode the message in base64
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

    # Send the email using Gmail API
    try:
        message = {'raw': raw_message}
        send_message = service.users().messages().send(userId="me", body=message).execute()
        print(f"Email sent to {recipient_email} with message ID: {send_message['id']}")
    except Exception as e:
        print(f"An error occurred: {e}")

def send_emails_automatically(email_data, sender_email, attachment=None):
    """Send emails using extracted data and Gmail API with OAuth2."""
    for entry in email_data:
        recipient_email = entry['email']
        subject = entry['subject']
        body = entry['body']

        # Send email to each recipient
        send_email(sender_email, recipient_email, subject, body, attachment)
        print(f"Email sent to {recipient_email}!")

# Sample usage of the email sending function
if __name__ == '__main__':
    sender_email = ""  # Replace with your email

    # Load the DOCX file and extract email data
    docx_file = 'Test_mail.docx'  # Update this to your file path
    text = extract_text_from_docx(docx_file)
    email_data = extract_emails_subjects_bodies(text)

    # PDF attachment (optional)
    attachment = 'Lines In The Sand MCLE.pdf'  # Path to the PDF file you want to attach

    # Send emails using the extracted data and attachment
    send_emails_automatically(email_data, sender_email, attachment)
