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
MARGIN = 36
GAP = 18
COL_WIDTH = (PAGE_WIDTH - (MARGIN * 2) - GAP) / 2

COLOR_PURPLE = colors.HexColor('#990099')
COLOR_GREEN = colors.HexColor('#008000')
COLOR_RED_QUOTE = colors.HexColor('#8B0000')
COLOR_TEXT = colors.HexColor('#111111')

class CleanCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

def dynamic_highlight(text):
    """Highlights dates, legislation, proper nouns, and warnings."""
    if not text or len(text.strip()) == 0:
        return text

    red_pattern = r'\b(EXAM TRAP|WARNING|CAUTION|IMPORTANT|LIMITATION|PITFALL|CRITICAL|NOTE|EXAM TIP|KEY PITFALL)\b'
    text = re.sub(red_pattern, r'<font backColor="#FF0000" color="white"><b>\g<0></b></font>', text, flags=re.IGNORECASE)

    date_pattern = r'\b((?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{1,2}(?:,\s+\d{4})?|\d{1,2}/\d{1,2}/\d{2,4}|\d{4}s?|\d{1,2}(?:st|nd|rd|th)\s+[Cc]entury)\b'
    text = re.sub(date_pattern, r'<font backColor="#FFFF00"><b>\g<0></b></font>', text)

    law_pattern = r'\b(Republic Act\s+(?:No\.\s*)?\d+|RA\s+\d+|Article\s+\d+|Section\s+\d+)\b'
    text = re.sub(law_pattern, r'<font backColor="#FFFF00"><b>\g<0></b></font>', text, flags=re.IGNORECASE)

    return text

def parse_pptx_file(filepath):
    """Parses text content from PowerPoint presentation slides."""
    lines = []
    try:
        prs = Presentation(filepath)
        for idx, slide in enumerate(prs.slides):
            slide_title = f"Slide {idx+1}"
            if slide.shapes.title and slide.shapes.title.text.strip():
                slide_title = slide.shapes.title.text.strip()
                lines.append(f"TITLE: {slide_title}")
            
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        txt = paragraph.text.strip()
                        if txt and txt != slide_title:
                            lines.append(txt)
    except Exception as e:
        lines.append(f"Error reading pptx: {e}")
    return lines

def parse_pdf_file(filepath):
    """Parses text content from PDF pages."""
    lines = []
    try:
        reader = PdfReader(filepath)
        for page in reader.pages:
            txt = page.extract_text()
            if txt:
                for line in txt.split('\n'):
                    l = line.strip()
                    if l:
                        lines.append(l)
    except Exception as e:
        lines.append(f"Error reading pdf: {e}")
    return lines

def build_pedagogical_modules(files_to_process, progress_callback=None, log_callback=None):
    """Groups extracted content into study modules with objectives, key terms, and practice questions."""
    total_files = len(files_to_process)
    extracted_data = []

    for idx, filepath in enumerate(files_to_process):
        fname = os.path.basename(filepath)
        clean_name = os.path.splitext(fname)[0].replace('_', ' ')
        if log_callback:
            log_callback(f"Extracting [{idx+1}/{total_files}]: {fname}")

        if filepath.lower().endswith('.pptx'):
            lines = parse_pptx_file(filepath)
        else:
            lines = parse_pdf_file(filepath)

        extracted_data.append((clean_name, lines))
        if progress_callback:
            progress_callback(int(((idx + 1) / total_files) * 40))

    structured_modules = []

    for mod_idx, (topic_name, lines) in enumerate(extracted_data):
        if not lines:
            continue

        titles = [l.replace("TITLE: ", "") for l in lines if l.startswith("TITLE: ")]
        body_lines = [l for l in lines if not l.startswith("TITLE: ")]

        concept_blocks = []
        current_block = []
        for line in body_lines:
            current_block.append(line)
            if len(current_block) >= 4 or line.endswith(('.', ':', ';')):
                concept_blocks.append(" ".join(current_block))
                current_block = []
        if current_block:
            concept_blocks.append(" ".join(current_block))

        recall_questions = []
        key_terms = []
        for block in concept_blocks:
            if ':' in block:
                parts = block.split(':', 1)
                term = parts[0].strip()
                def_text = parts[1].strip()
                if len(term.split()) <= 5 and len(def_text) > 10:
                    key_terms.append((term, def_text))
                    recall_questions.append((f"What is {term}?", def_text[:120] + "..."))

        module_dict = {
            'module_num': mod_idx + 1,
            'title': f"{mod_idx+1}. {topic_name}",
            'objectives': [
                f"Master core concepts and historical/theoretical context of {topic_name}.",
                f"Identify key dates, names, legislation, and key principles for the midterm exam.",
                f"Distinguish critical exam traps, common misconceptions, and comparative definitions."
            ],
            'subtopics': titles if titles else ["Core Concepts & Syntheses"],
            'concept_blocks': concept_blocks[:8],
            'key_terms': key_terms[:6],
            'recall_questions': recall_questions[:4]
        }
        structured_modules.append(module_dict)

    return structured_modules

def convert_materials_to_pdf(input_path, output_pdf_path, progress_callback=None, log_callback=None):
    """Generates a 2-column Exam Reviewer PDF from input materials."""
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

    modules = build_pedagogical_modules(files_to_process, progress_callback, log_callback)

    if log_callback:
        log_callback("Formatting exam reviewer layout...")

    course_title = os.path.basename(os.path.normpath(input_path)).replace('_', ' ').replace('-', ' ').title()
    if not course_title or course_title.lower() == 'downloads':
        course_title = "Midterm Exam Comprehensive Reviewer"
    else:
        course_title = f"{course_title} — Midterm Exam Reviewer"

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

    quote_style = ParagraphStyle(
        'CustomQuote',
        parent=body_style,
        fontName='Helvetica-Oblique',
        textColor=COLOR_RED_QUOTE,
        leftIndent=6,
        spaceBefore=3,
        spaceAfter=3
    )

    table_header_style = ParagraphStyle(
        'TableHeader',
        parent=body_style,
        fontName='Helvetica-Bold',
        fontSize=7.5,
        leading=9.5,
        alignment=1,
        textColor=colors.HexColor('#222222')
    )

    table_cell_style = ParagraphStyle(
        'TableCell',
        parent=body_style,
        fontSize=7.2,
        leading=9.0,
        spaceAfter=0
    )

    story = []

    story.append(Paragraph(course_title, h1_style))
    story.append(Spacer(1, 4))

    for mod in modules:
        story.append(Paragraph(mod['title'], h1_style))

        story.append(Paragraph("Exam Learning Objectives", h2_style))
        for obj in mod['objectives']:
            story.append(Paragraph(f"• {dynamic_highlight(obj)}", bullet_style))

        story.append(Paragraph("High-Yield Concept Syntheses & Notes", h2_style))
        for block in mod['concept_blocks']:
            story.append(Paragraph(dynamic_highlight(block), body_style))

        if mod['key_terms']:
            story.append(Paragraph("Key Terminology & Definitions", h2_style))
            table_data = [[Paragraph("Term", table_header_style), Paragraph("Definition / High-Yield Note", table_header_style)]]
            for term, defn in mod['key_terms']:
                table_data.append([
                    Paragraph(f"<font backColor=\"#FFFF00\"><b>{term}</b></font>", table_cell_style),
                    Paragraph(dynamic_highlight(defn), table_cell_style)
                ])
            t_terms = Table(table_data, colWidths=[COL_WIDTH*0.35, COL_WIDTH*0.65])
            t_terms.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#E6E6FA')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('TOPPADDING', (0,0), (-1,-1), 2),
                ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ]))
            story.append(t_terms)
            story.append(Spacer(1, 4))

        if mod['recall_questions']:
            story.append(Paragraph("Active Recall & Self-Test Questions", h2_style))
            for q_num, (q_text, a_text) in enumerate(mod['recall_questions'], 1):
                story.append(Paragraph(f"<b>Q{q_num}: {q_text}</b>", body_style))
                story.append(Paragraph(f"<i>Answer Note: {dynamic_highlight(a_text)}</i>", quote_style))

    if progress_callback:
        progress_callback(90)

    doc.build(story, canvasmaker=CleanCanvas)

    if progress_callback:
        progress_callback(100)

    if log_callback:
        log_callback(f"Successfully generated Reviewer PDF at: {output_pdf_path}")

    return output_pdf_path
