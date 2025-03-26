import re
import docx
from docx.oxml.ns import qn
from io import BytesIO
from docx.shared import RGBColor

def extract_rich_text_from_docx(docx_file):
    """
    Extracts rich text from a DOCX file and converts it to HTML while preserving 
    line breaks, paragraph spacing, and inline text formatting.
    """
    doc = docx.Document(docx_file)
    html_content = []

    for para in doc.paragraphs:
        if not para.text.strip():  # Preserve blank lines as <p>&nbsp;</p>
            html_content.append("<p>&nbsp;</p>")
            continue
        
        # Handle indentation for bullet-like or indented content
        if para.text.startswith("    "):  
            html_content.append(f'<p style="text-indent: 2em;">{para.text.strip()}</p>')
            continue

        para_html = []
        current_hyperlink = None
        p = para._p  # XML element access

        for elem in p.iterchildren():
            if elem.tag.endswith('hyperlink'):  
                rel_id = elem.get(qn('r:id'))
                if rel_id and para.part and rel_id in para.part.rels:
                    current_hyperlink = para.part.rels[rel_id].target_ref
                    for run_elem in elem.iterchildren():
                        if run_elem.tag.endswith('r'):
                            run_text, formatting = extract_run_text_and_formatting(run_elem)
                            formatted_text = apply_formatting(run_text, formatting)
                            para_html.append(f'<a href="{current_hyperlink}" style="color: blue; text-decoration: underline;">{formatted_text}</a>')
                    current_hyperlink = None
            elif elem.tag.endswith('r'):  
                run_text, formatting = extract_run_text_and_formatting(elem)
                run_text = run_text.replace("\n", "<br>")
                formatted_text = apply_formatting(run_text, formatting)
                if current_hyperlink:
                    formatted_text = f'<a href="{current_hyperlink}" style="color: blue; text-decoration: underline;">{formatted_text}</a>'
                para_html.append(formatted_text)

        if para.text.startswith("Where are") or para.text.endswith("?"):
            html_content.append(f'<p style="margin-left: 20px;">{"".join(para_html)}</p>')
        else:
            html_content.append(f"<p>{''.join(para_html)}</p>")

    return "".join(html_content)

def extract_run_text_and_formatting(run_elem):
    """
    Extracts text and formatting (bold, italic, underline, color) from a run element.
    """
    run_text = ''
    formatting = {'bold': False, 'italic': False, 'underline': False, 'color': None}

    for t in run_elem.iterchildren():
        if t.tag.endswith('t'):  
            run_text += t.text
        elif t.tag.endswith('rPr'):  
            for prop in t.iterchildren():
                if prop.tag.endswith('b'):
                    formatting['bold'] = True
                elif prop.tag.endswith('i'):
                    formatting['italic'] = True
                elif prop.tag.endswith('u'):
                    formatting['underline'] = True
                elif prop.tag.endswith('color') and 'val' in prop.attrib:
                    formatting['color'] = prop.attrib['val']

    return run_text, formatting

def apply_formatting(text, formatting):
    """
    Applies HTML tags based on the formatting dictionary.
    """
    formatted_text = text
    if formatting['color']:
        formatted_text = f'<span style="color: #{formatting["color"]};">{formatted_text}</span>'
    if formatting['bold']:
        formatted_text = f'<strong>{formatted_text}</strong>'
    if formatting['italic']:
        formatted_text = f'<em>{formatted_text}</em>'
    if formatting['underline']:
        formatted_text = f'<u>{formatted_text}</u>'
    return formatted_text

def extract_emails_subjects_bodies(html_content):
    """
    Extracts emails, subjects, and bodies from the provided HTML content.
    This version ensures better handling of emails, subject headers, and body text.
    """
    emails_data = []
    
    if isinstance(html_content, list):
        html_content = "".join(html_content)
    
    paragraphs = re.split(r'(?=<p>|</p>)', html_content)
    paragraphs = [p for p in paragraphs if p.strip() and p not in ['<p>', '</p>']]
    
    email_pattern = re.compile(r'[\w\.-]+@[\w\.-]+(?:\.[\w]+)+')
    subject_pattern = re.compile(r'Lines\s+In\s+the\s+Sand:\s*(.*?)(?:\.|$)', re.IGNORECASE)
    body_start_pattern = re.compile(r'Dear\s+[A-Za-z]+', re.IGNORECASE)
    
    current_email = None
    current_subject = None
    current_body = []
    in_body = False
    
    for para in paragraphs:
        email_match = email_pattern.search(para)
        if email_match:
            if current_email:
                emails_data.append({
                    'email': current_email,
                    'subject': current_subject,
                    'body': format_email_body("".join(current_body))
                })
            current_email = email_match.group()
            current_subject = None
            current_body = []
            in_body = False
            continue
        
        if not in_body and not current_subject:
            subject_match = subject_pattern.search(para)
            if subject_match:
                current_subject = subject_match.group(1).strip()
                if not current_subject.endswith('.'):
                    current_subject += '.'
                continue
        
        if not in_body and body_start_pattern.search(para):
            in_body = True
        
        if in_body or (current_email and not current_subject):
            current_body.append(para)
    
    if current_email:
        emails_data.append({
            'email': current_email,
            'subject': current_subject,
            'body': format_email_body("".join(current_body))
        })
    
    return emails_data
def format_email_body(body_html):
    """
    Formats the email body with proper HTML structure.
    Ensures proper greetings and signatures.
    """
    # Clean up the HTML
    body_html = re.sub(r'<p>\s*</p>', '', body_html)  # Remove empty paragraphs
    
    # Add greeting if not present
    if not re.search(r'Dear\s+[A-Za-z]+', body_html, re.IGNORECASE):
        body_html = f"<p>Dear Recipient,</p>{body_html}"
    
    # Add signature if not present
    #if not re.search(r'(Sincerely|Regards|Best\s+wishes)', body_html, re.IGNORECASE):
        #body_html = f"{body_html}<p>Sincerely,</p><p>The Team</p>"
    
    return body_html