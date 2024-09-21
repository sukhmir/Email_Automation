import re
import docx

# Function to extract text from the DOCX file
def extract_text_from_docx(docx_file):
    doc = docx.Document(docx_file)
    text = []
    for para in doc.paragraphs:
        if para.text.strip():  # Only append non-empty paragraphs
            text.append(para.text.strip())
    return text

# Function to extract emails, subjects, and bodies from the text
def extract_emails_subjects_bodies(text):
    emails_data = []
    email_pattern = r'[\w\.-]+@[\w\.-]+'  # Regex to match email addresses
    subject_start_pattern = r'Lines\s+In\s+the\s+Sand:'   # Regex to match subject start pattern
    body_start_pattern = r'\bDear\b'       # Regex to detect the start of the body
    
    current_email = None
    current_subject = None
    current_body = []
    
    # Process the text line by line
    for line in text:
        if re.match(email_pattern, line):  # Detect email address
            # If there is an existing email, save the previous data
            if current_email:
                emails_data.append({
                    'email': current_email,
                    'subject': current_subject,
                    'body': format_email_body("\n".join(current_body).strip())
                })
            current_email = line
            current_subject = None
            current_body = []
        elif re.search(subject_start_pattern, line):  # Detect subject start using regex
            current_subject = extract_subject_after_pattern(line, subject_start_pattern)
        elif re.search(body_start_pattern, line):  # Detect start of body using regex
            current_body = [line]
        else:
            current_body.append(line)  # Continue accumulating the body

    # Add the last email entry after the loop ends
    if current_email:
        emails_data.append({
            'email': current_email,
            'subject': current_subject,
            'body': format_email_body("\n".join(current_body).strip())
        })
    
    return emails_data

# Function to extract the subject after the specific pattern and end at the first full stop
def extract_subject_after_pattern(line, pattern):
    # Find the part of the line that starts after the regex pattern
    match = re.search(pattern, line)
    if match:
        # Get the remaining text after the pattern
        subject_part = line[match.end():].strip()
        # Split the remaining part at the first full stop and return the subject
        if '.' in subject_part:
            return subject_part.split('.')[0].strip() + '.'  # Return text until the first full stop
        return subject_part  # Return the whole subject if no full stop found
    return None

# Function to format the email body professionally
def format_email_body(body):
    # Add greetings, signatures, and structure
    formatted_body = []

    # Add greeting if "Dear" is not in the first line
    if not body.startswith("Dear"):
        formatted_body.append("Dear Recipient,")

    # Break paragraphs with double line breaks and indentations
    paragraphs = body.split('\n')
    for paragraph in paragraphs:
        formatted_body.append(paragraph.strip())  # Remove any extra spaces

    return "\n\n".join(formatted_body)

# Example usage
docx_file = 'Test_mail.docx' 
text = extract_text_from_docx(docx_file)

# Extract emails, subjects, and bodies with professional formatting
emails_subjects_bodies = extract_emails_subjects_bodies(text)

# Output results
if not emails_subjects_bodies:
    print("No email found")
else:
    for entry in emails_subjects_bodies:
        print(f"Email: {entry['email']}")
        print(f"Subject: {entry['subject']}")
        print(f"Body:\n{entry['body']}")
        print('-' * 50)
