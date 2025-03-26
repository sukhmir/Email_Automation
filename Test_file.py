import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO
from docx.shared import RGBColor

def extract_rich_text_from_docx(docx_file):
    """
    Extracts rich text with formatting and hyperlinks from DOCX to HTML.
    Uses direct XML parsing to ensure hyperlinks are captured.
    """
    doc = docx.Document(docx_file)
    html_content = []
    
    for para in doc.paragraphs:
        if not para.text.strip():
            continue
            
        para_html = []
        current_hyperlink = None
        
        # Get the paragraph's XML element
        p = para._p
        
        # Process all elements in the paragraph (runs, hyperlinks, etc.)
        for elem in p.iterchildren():
            if elem.tag.endswith('hyperlink'):  # Found a hyperlink
                rel_id = elem.get(qn('r:id'))
                if rel_id and para.part and rel_id in para.part.rels:
                    current_hyperlink = para.part.rels[rel_id].target_ref
                    # Process all runs within this hyperlink
                    for run_elem in elem.iterchildren():
                        if run_elem.tag.endswith('r'):  # This is a run
                            run_text = ''
                            for t in run_elem.iterchildren():
                                if t.tag.endswith('t'):  # Text element
                                    run_text += t.text
                            if run_text.strip():
                                para_html.append(f'<a href="{current_hyperlink}">{run_text}</a>')
                    current_hyperlink = None
            elif elem.tag.endswith('r'):  # Regular run
                run_text = ''
                is_bold = False
                is_italic = False
                is_underline = False
                
                # Extract text and formatting
                for t in elem.iterchildren():
                    if t.tag.endswith('t'):  # Text element
                        run_text += t.text
                    elif t.tag.endswith('rPr'):  # Run properties
                        for prop in t.iterchildren():
                            if prop.tag.endswith('b'):  # Bold
                                is_bold = True
                            elif prop.tag.endswith('i'):  # Italic
                                is_italic = True
                            elif prop.tag.endswith('u'):  # Underline
                                is_underline = True
                
                if run_text.strip():
                    formatted_text = run_text
                    if is_bold:
                        formatted_text = f'<strong>{formatted_text}</strong>'
                    if is_italic:
                        formatted_text = f'<em>{formatted_text}</em>'
                    if is_underline:
                        formatted_text = f'<span style="text-decoration:underline">{formatted_text}</span>'
                    
                    if current_hyperlink:
                        formatted_text = f'<a href="{current_hyperlink}">{formatted_text}</a>'
                    
                    para_html.append(formatted_text)
        
        if para_html:
            html_content.append(f"<p>{''.join(para_html)}</p>")
    
    return "".join(html_content)

# Example usage
if __name__ == "__main__":
    docx_file_path = "Merge Test 1 (1).docx"
    
    with open(docx_file_path, "rb") as file:
        docx_file = BytesIO(file.read())
    
    html_content = extract_rich_text_from_docx(docx_file)
    
    print("Formatted HTML Content:")
    print(html_content)