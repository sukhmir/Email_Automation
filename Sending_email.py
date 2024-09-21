import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from extract_and_format_emails import extract_text_from_docx, extract_emails_subjects_bodies

# Function to send emails with an attachment
def send_email(sender_email, sender_password, recipient_email, subject, body, attachment=None):
    # Set up the email server (Gmail example)
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()  # Start TLS encryption

    # Log in to your email account
    server.login(sender_email, sender_password)

    # Create the email content
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Attach the body as plain text
    msg.attach(MIMEText(body, 'plain'))

    # Check if there is an attachment
    if attachment:
        try:
            # Open the file to attach
            with open(attachment, 'rb') as file:
                # Create a MIMEBase instance
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file.read())
                # Encode to base64
                encoders.encode_base64(part)
                # Add headers
                part.add_header('Content-Disposition', f'attachment; filename={attachment}')
                # Attach to the email
                msg.attach(part)
        except Exception as e:
            print(f"Failed to attach file {attachment}: {e}")

    # Send the email
    server.sendmail(sender_email, recipient_email, msg.as_string())

    # Close the server
    server.quit()

# Sending emails automatically with optional PDF attachments
def send_emails_automatically(email_data, sender_email, sender_password, attachment=None):
    for entry in email_data:
        recipient_email = entry['email']
        subject = entry['subject']
        body = entry['body']

        # Send email to each recipient
        send_email(sender_email, sender_password, recipient_email, subject, body, attachment)
        print(f"Email sent to {recipient_email}!")

# Sample usage of the email sending function
if __name__ == '__main__':
    sender_email = "shoaibsukhmir10@gmail.com"  # Replace with your email
    sender_password = "brbh ftzq orag ybig"  # Replace with your app password

    # Load the DOCX file and extract email data
    docx_file = 'Test_mail.docx'  # Update this to your file path
    text = extract_text_from_docx(docx_file)
    email_data = extract_emails_subjects_bodies(text)

    # PDF attachment (optional)
    attachment = 'Lines In The Sand MCLE.pdf'  # Path to the PDF file you want to attach

    # Send emails using the extracted data and attachment
    send_emails_automatically(email_data, sender_email, sender_password, attachment)
