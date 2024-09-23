import streamlit as st
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO
from dotenv import load_dotenv  # Import dotenv to load .env file
from extract_and_format_emails import extract_text_from_docx, extract_emails_subjects_bodies

# Load environment variables from .env file
load_dotenv()

# Function to send emails with an attachment
def send_email(sender_email, sender_password, recipient_email, subject, body, attachment=None):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()  # Start TLS encryption
    server.login(sender_email, sender_password)

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    if attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename=attachment.pdf')
        msg.attach(part)

    server.sendmail(sender_email, recipient_email, msg.as_string())
    server.quit()

def send_emails_automatically(email_data, sender_email, sender_password, attachment=None):
    for entry in email_data:
        recipient_email = entry['email']
        subject = entry['subject']
        body = entry['body']
        send_email(sender_email, sender_password, recipient_email, subject, body, attachment)
        st.success(f"Email sent to {recipient_email}!")

def main():
    st.title("Email Automation App")

    # Retrieve email and password from environment variables
    sender_email = os.getenv("SENDER_EMAIL")
    sender_password = os.getenv("SENDER_PASSWORD")

    docx_file = st.file_uploader("Upload DOCX File for Email Data", type=["docx"])
    pdf_attachment = st.file_uploader("Attach a PDF (Optional)", type=["pdf"])

    if st.button("Send Emails"):
        if docx_file is not None:
            email_data = extract_emails_subjects_bodies(extract_text_from_docx(BytesIO(docx_file.read())))
            send_emails_automatically(email_data, sender_email, sender_password, pdf_attachment)
        else:
            st.error("Please upload a DOCX file with email data.")

if __name__ == '__main__':
    main()
