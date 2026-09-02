import os
import re
from pptx import Presentation
from pypdf import PdfReader

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas

PAGE_WIDTH, PAGE_HEIGHT = letter
MARGIN = 36   # Standard 0.5 inch margin
GAP = 18      # Gap between columns
COL_WIDTH = (PAGE_WIDTH - (MARGIN * 2) - GAP) / 2  # ~261 pt per column

COLOR_PURPLE = colors.HexColor('#990099')  # Heading 1
COLOR_GREEN = colors.HexColor('#008000')   # Heading 2
COLOR_RED_QUOTE = colors.HexColor('#8B0000') # Quotes
COLOR_TEXT = colors.HexColor('#111111')

class CleanCanvas(canvas.Canvas):
    """Canvas without headers, footers, page numbers, or horizontal lines."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

def dynamic_highlight(text):
    """Dynamically detects and highlights key terms, dates, proper nouns, and exam warnings in ANY text."""
    if not text or len(text.strip()) == 0:
        return text

    # Red Highlights for Warning/Trap/Limitation keywords
    red_pattern = r'\b(EXAM TRAP|WARNING|CAUTION|IMPORTANT|LIMITATION|PITFALL|CRITICAL|NOTE)\b'
    text = re.sub(red_pattern, r'<font backColor="#FF0000" color="white"><b>\g<0></b></font>', text, flags=re.IGNORECASE)

    # Yellow Highlights for Dates (e.g., June 12, 1956, 19th Century, 1872, 2026)
    date_pattern = r'\b((?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{1,2}(?:,\s+\d{4})?|\d{1,2}/\d{1,2}/\d{2,4}|\d{4}s?|\d{1,2}(?:st|nd|rd|th)\s+[Cc]entury)\b'
    text = re.sub(date_pattern, r'<font backColor="#FFFF00"><b>\g<0></b></font>', text)

    # Yellow Highlights for Laws, Acts, or Code symbols (e.g. Republic Act No. 1425, RA 1425)
    law_pattern = r'\b(Republic Act\s+(?:No\.\s*)?\d+|RA\s+\d+|Article\s+\d+|Section\s+\d+)\b'
    text = re.sub(law_pattern, r'<font backColor="#FFFF00"><b>\g<0></b></font>', text, flags=re.IGNORECASE)

    return text

def parse_pptx_file(filepath):
    """Parses a PowerPoint file into structured slides with titles and body content."""
    slides_data = []
    try:
        prs = Presentation(filepath)
        for idx, slide in enumerate(prs.slides):
            slide_title = f"Slide {idx+1}"
            slide_lines = []
            
            # Find slide title
            if slide.shapes.title and slide.shapes.title.text.strip():
                slide_title = slide.shapes.title.text.strip()
            
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        txt = paragraph.text.strip()
                        if txt and txt != slide_title:
                            slide_lines.append(txt)
            
            slides_data.append({
                'title': slide_title,
                'content': slide_lines
            })
    except Exception as e:
        slides_data.append({
            'title': f"Error Reading {os.path.basename(filepath)}",
            'content': [str(e)]
        })
    return slides_data

def parse_pdf_file(filepath):
    """Parses a PDF file into structured pages/sections."""
    pages_data = []
    try:
        reader = PdfReader(filepath)
        for idx, page in enumerate(reader.pages):
            txt = page.extract_text()
            if txt:
                lines = [l.strip() for l in txt.split('\n') if l.strip()]
                title = lines[0] if lines else f"Page {idx+1}"
                content = lines[1:] if len(lines) > 1 else lines
                pages_data.append({
                    'title': title,
                    'content': content
                })
    except Exception as e:
        pages_data.append({
            'title': f"Error Reading {os.path.basename(filepath)}",
            'content': [str(e)]
        })
    return pages_data

def convert_materials_to_pdf(input_path, output_pdf_path, progress_callback=None, log_callback=None):
    """Generic Dynamic Converter: Converts ANY subject materials (PPTX/PDF) into a 2-column Reviewer PDF."""
    if log_callback:
        log_callback(f"Scanning materials in: {input_path}")

    files_to_process = []
    if os.path.isdir(input_path):
        for root, dirs, files in os.walk(input_path):
            for f in sorted(files):
                if f.lower().endswith(('.pptx', '.pdf')):
                    files_to_process.append(os.path.join(root, f))
    elif os.path.isfile(input_path):
        files_to_process.append(input_path)

    if not files_to_process:
        raise ValueError("No PowerPoint (.pptx) or PDF (.pdf) files found in selected location.")

    total_files = len(files_to_process)
    if log_callback:
        log_callback(f"Found {total_files} file(s) for dynamic reviewer compilation.")

    # Subject / Course Title derived from folder name or first file
    course_title = os.path.basename(os.path.normpath(input_path)).replace('_', ' ').replace('-', ' ').title()
    if not course_title or course_title.lower() == 'downloads':
        course_title = "Exam Reviewer & Study Guide"

    # Extract all files dynamically
    all_modules = []
    for idx, filepath in enumerate(files_to_process):
        fname = os.path.basename(filepath)
        clean_name = os.path.splitext(fname)[0].replace('_', ' ')
        if log_callback:
            log_callback(f"Parsing [{idx+1}/{total_files}]: {fname}")

        if filepath.lower().endswith('.pptx'):
            sections = parse_pptx_file(filepath)
        else:
            sections = parse_pdf_file(filepath)

        all_modules.append({
            'module_num': idx + 1,
            'module_title': f"{idx+1}. {clean_name}",
            'sections': sections
        })

        if progress_callback:
            progress_callback(int(((idx + 1) / total_files) * 50))

    if log_callback:
        log_callback("Formatting 2-column layout and generating PDF document...")

    # Build ReportLab Document
    doc = BaseDocTemplate(
        output_pdf_path,
        pagesize=letter,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=MARGIN,
        bottomMargin=MARGIN
    )

    height_normal = PAGE_HEIGHT - (MARGIN * 2)
    frame1 = Frame(MARGIN, MARGIN, COL_WIDTH, height_normal, id='col1', leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    frame2 = Frame(MARGIN + COL_WIDTH + GAP, MARGIN, COL_WIDTH, height_normal, id='col2', leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)

    two_col_template = PageTemplate(id='TwoColPage', frames=[frame1, frame2])
    doc.addPageTemplates([two_col_template])

    styles = getSampleStyleSheet()
    
    body_style = ParagraphStyle(
        'CustomBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11.5,
        textColor=COLOR_TEXT,
        spaceAfter=5
    )

    bullet_style = ParagraphStyle(
        'CustomBullet',
        parent=body_style,
        leftIndent=10,
        firstLineIndent=-8,
        spaceAfter=3
    )

    h1_style = ParagraphStyle(
        'CustomH1',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=11.5,
        leading=13.5,
        textColor=COLOR_PURPLE,
        spaceBefore=9,
        spaceAfter=3,
        keepWithNext=True
    )

    h2_style = ParagraphStyle(
        'CustomH2',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=9.5,
        leading=11.5,
        textColor=COLOR_GREEN,
        spaceBefore=6,
        spaceAfter=2,
        keepWithNext=True
    )

    story = []

    # Document Title Header
    story.append(Paragraph(course_title, h1_style))
    story.append(Spacer(1, 4))

    # Dynamically build story from extracted modules and sections
    for module in all_modules:
        story.append(Paragraph(module['module_title'], h1_style))
        
        for section in module['sections']:
            sec_title = section['title']
            if sec_title and len(sec_title.strip()) > 0:
                story.append(Paragraph(dynamic_highlight(sec_title), h2_style))
            
            for line in section['content']:
                if not line or len(line.strip()) == 0:
                    continue
                
                # Format bullet points vs paragraphs
                if line.startswith(('•', '-', '*', '1.', '2.', '3.', '4.', '5.')):
                    formatted_bullet = f"• {dynamic_highlight(line.lstrip('•-*123456789. '))}"
                    story.append(Paragraph(formatted_bullet, bullet_style))
                elif ':' in line and len(line.split(':')[0].split()) <= 4:
                    # Key-value / term definition pair
                    key, val = line.split(':', 1)
                    formatted_kv = f"<font backColor=\"#FFFF00\"><b>{key.strip()}</b></font>: {dynamic_highlight(val.strip())}"
                    story.append(Paragraph(formatted_kv, body_style))
                else:
                    story.append(Paragraph(dynamic_highlight(line), body_style))

    if progress_callback:
        progress_callback(85)

    doc.build(story, canvasmaker=CleanCanvas)
    
    if progress_callback:
        progress_callback(100)

    if log_callback:
        log_callback(f"Successfully generated Reviewer PDF at: {output_pdf_path}")

    return output_pdf_path
