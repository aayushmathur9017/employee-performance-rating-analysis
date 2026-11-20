"""
High-Fidelity DOCX to PDF Conversion Engine
Using PyMuPDF (fitz) and python-docx.
Custom-tailored for ABES Engineering College MCA Academic Submissions.
Student: Aayush Mathur | Enrolment No: 2500320140001
"""

import os
import re
import fitz
import docx
from docx.text.paragraph import Paragraph
from docx.table import Table

# A4 dimensions in points
PAGE_W = 595.0
PAGE_H = 842.0
MARGIN_LEFT = 45.0
MARGIN_RIGHT = 550.0
USABLE_W = MARGIN_RIGHT - MARGIN_LEFT
MARGIN_TOP = 65.0
MARGIN_BOTTOM = 785.0
USABLE_H = MARGIN_BOTTOM - MARGIN_TOP

# Color Palette
NAVY_HEADER = (0.10, 0.20, 0.45)   # #1A3373
BLUE_ACCENT = (0.15, 0.35, 0.75)   # #2659BF
TEXT_DARK   = (0.12, 0.14, 0.18)   # #1F242E
TEXT_MUTED  = (0.40, 0.45, 0.50)   # #667380
BORDER_COL  = (0.80, 0.84, 0.88)   # #CCD6E0
HEADER_BG   = (0.12, 0.23, 0.54)   # #1F3B8A
ZEBRA_BG    = (0.96, 0.97, 0.99)   # #F5F7FC
WHITE       = (1.00, 1.00, 1.00)


def wrap_text(text: str, font: str, size: float, max_width: float):
    """Wraps text into a list of lines that fit within max_width."""
    if not text:
        return ['']
    words = text.split()
    if not words:
        return ['']
    lines = []
    cur_line = []
    for w in words:
        test_line = ' '.join(cur_line + [w])
        try:
            w_len = fitz.get_text_length(test_line, fontname=font, fontsize=size)
        except Exception:
            w_len = len(test_line) * size * 0.5
        if w_len <= max_width:
            cur_line.append(w)
        else:
            if cur_line:
                lines.append(' '.join(cur_line))
                cur_line = [w]
            else:
                lines.append(w)
                cur_line = []
    if cur_line:
        lines.append(' '.join(cur_line))
    return lines


def convert_docx(docx_path: str, pdf_path: str, doc_title: str = "Academic Document"):
    """Converts a Word .docx file into a beautifully formatted PDF."""
    if not os.path.exists(docx_path):
        print(f"Error: {docx_path} not found.")
        return False

    doc_docx = docx.Document(docx_path)
    doc_pdf = fitz.open()

    def add_page():
        p = doc_pdf.new_page(width=PAGE_W, height=PAGE_H)
        # Header banner
        p.draw_line(fitz.Point(MARGIN_LEFT, 45), fitz.Point(MARGIN_RIGHT, 45), color=NAVY_HEADER, width=0.8)
        p.insert_text((MARGIN_LEFT, 40), "ABES Engineering College, Ghaziabad | MCA HCL Training", fontsize=8, fontname="helv", color=TEXT_MUTED)
        p.insert_text((MARGIN_RIGHT - 130, 40), "Aayush Mathur (2500320140001)", fontsize=8, fontname="Helvetica-Bold", color=NAVY_HEADER)
        return p

    page = add_page()
    y = MARGIN_TOP

    for child in doc_docx.element.body:
        tag = child.tag
        if tag.endswith('p'):
            p = Paragraph(child, doc_docx)
            text = p.text.strip()
            if not text:
                y += 5
                continue

            style_name = p.style.name.lower()
            is_title = 'title' in style_name or (y < 120 and len(doc_pdf) == 1 and ('document' in text.lower() or 'analysis' in text.lower() or 'step' in text.lower()))
            is_h1 = 'heading 1' in style_name or text.startswith('Section ') or re.match(r'^\d+\.\s+[A-Z]', text)
            is_h2 = 'heading 2' in style_name or re.match(r'^\d+\.\d+\s+[A-Z]', text)
            is_bullet = 'bullet' in style_name or text.startswith(('•', '▪', '-', '*'))

            if is_title:
                font = 'Helvetica-Bold'
                size = 14.0
                color = NAVY_HEADER
                line_h = 18.0
                pad_top = 10.0
                pad_bot = 6.0
                indent = 0
            elif is_h1:
                font = 'Helvetica-Bold'
                size = 11.5
                color = NAVY_HEADER
                line_h = 15.0
                pad_top = 12.0
                pad_bot = 4.0
                indent = 0
            elif is_h2:
                font = 'Helvetica-Bold'
                size = 10.0
                color = BLUE_ACCENT
                line_h = 13.0
                pad_top = 8.0
                pad_bot = 3.0
                indent = 0
            elif is_bullet:
                font = 'helv'
                size = 9.0
                color = TEXT_DARK
                line_h = 12.0
                pad_top = 2.0
                pad_bot = 2.0
                indent = 15.0
            else:
                font = 'helv'
                size = 9.0
                color = TEXT_DARK
                line_h = 12.0
                pad_top = 2.0
                pad_bot = 4.0
                indent = 0

            lines = wrap_text(text, font, size, USABLE_W - indent)
            block_h = pad_top + len(lines) * line_h + pad_bot

            if y + block_h > MARGIN_BOTTOM:
                page = add_page()
                y = MARGIN_TOP

            y += pad_top
            for l in lines:
                page.insert_text((MARGIN_LEFT + indent, y + size), l, fontname=font, fontsize=size, color=color)
                y += line_h
            y += pad_bot

        elif tag.endswith('tbl'):
            tbl = Table(child, doc_docx)
            num_cols = len(tbl.columns)
            if num_cols == 0:
                continue

            col_w = USABLE_W / num_cols
            header_lines = []

            for row_idx, row in enumerate(tbl.rows):
                is_header = (row_idx == 0)
                font = 'Helvetica-Bold' if is_header else 'helv'
                size = 8.0
                tcolor = WHITE if is_header else TEXT_DARK
                bgcolor = HEADER_BG if is_header else (ZEBRA_BG if row_idx % 2 == 1 else WHITE)
                line_h = 10.5

                cell_lines = []
                for cell in row.cells:
                    ctext = cell.text.strip().replace('\n', ' ')
                    clines = wrap_text(ctext, font, size, col_w - 8)
                    cell_lines.append(clines)

                if is_header:
                    header_lines = cell_lines

                row_h = max(len(cl) for cl in cell_lines) * line_h + 6.0
                row_h = max(row_h, 15.0)

                if y + row_h > MARGIN_BOTTOM:
                    page = add_page()
                    y = MARGIN_TOP

                    # Repeat header row on new page
                    if not is_header and header_lines:
                        h_row_h = max(len(cl) for cl in header_lines) * line_h + 6.0
                        for c_idx, clines in enumerate(header_lines):
                            cx0 = MARGIN_LEFT + c_idx * col_w
                            cx1 = cx0 + col_w
                            page.draw_rect(fitz.Rect(cx0, y, cx1, y + h_row_h), color=BORDER_COL, fill=HEADER_BG, width=0.5)
                            cy = y + 3.0 + size
                            for cl in clines:
                                page.insert_text((cx0 + 4, cy), cl, fontname='Helvetica-Bold', fontsize=size, color=WHITE)
                                cy += line_h
                        y += h_row_h

                for c_idx, clines in enumerate(cell_lines):
                    cx0 = MARGIN_LEFT + c_idx * col_w
                    cx1 = cx0 + col_w
                    page.draw_rect(fitz.Rect(cx0, y, cx1, y + row_h), color=BORDER_COL, fill=bgcolor, width=0.5)
                    cy = y + 3.0 + size
                    for cl in clines:
                        page.insert_text((cx0 + 4, cy), cl, fontname=font, fontsize=size, color=tcolor)
                        cy += line_h
                y += row_h
            y += 8.0

    # If ER Diagram document, insert high-res PNG at end
    if "09_er_diagram" in docx_path.lower() and os.path.exists("docs/09_ER_Diagram.png"):
        p_img = add_page()
        p_img.insert_text((MARGIN_LEFT, MARGIN_TOP), "Relational Database Schema (Crow's Foot ER Diagram)", fontsize=11, fontname="Helvetica-Bold", color=NAVY_HEADER)
        img_rect = fitz.Rect(MARGIN_LEFT, MARGIN_TOP + 15, MARGIN_RIGHT, MARGIN_BOTTOM - 20)
        p_img.insert_image(img_rect, filename="docs/09_ER_Diagram.png")

    # Page numbering
    total_pages = len(doc_pdf)
    for idx, p in enumerate(doc_pdf):
        p.draw_line(fitz.Point(MARGIN_LEFT, 800), fitz.Point(MARGIN_RIGHT, 800), color=BORDER_COL, width=0.5)
        p.insert_text((MARGIN_LEFT, 814), f"Employee Performance Rating Analysis | {doc_title}", fontsize=7.5, fontname="helv", color=TEXT_MUTED)
        p.insert_text((MARGIN_RIGHT - 65, 814), f"Page {idx+1} of {total_pages}", fontsize=7.5, fontname="Helvetica-Bold", color=NAVY_HEADER)

    doc_pdf.save(pdf_path)
    print(f"Generated: {pdf_path} ({total_pages} pages)")
    return True


DOCX_DOCUMENTS = [
    ("docs/01_Project_Understanding.docx", "docs/01_Project_Understanding.pdf", "Step 1: Project Understanding"),
    ("docs/02_HMW_Statements.docx", "docs/02_HMW_Statements.pdf", "Step 2: HMW Statements"),
    ("docs/03_Needs_Analysis.docx", "docs/03_Needs_Analysis.pdf", "Step 3: Needs Analysis & RACI"),
    ("docs/05_BRD.docx", "docs/05_BRD.pdf", "Step 5: Business Requirements Document"),
    ("docs/07_TRD.docx", "docs/07_TRD.pdf", "Step 7: Technical Requirements Document"),
    ("docs/09_ER_Diagram.docx", "docs/09_ER_Diagram.pdf", "Step 9: ER Diagram & DB Schema"),
    ("docs/11_User_Stories.docx", "docs/11_User_Stories.pdf", "Step 11: Agile User Stories"),
    ("docs/12_UI_Framework.docx", "docs/12_UI_Framework.pdf", "Step 12: UI Framework & Design Tokens"),
    ("docs/13_Test_Cases.docx", "docs/13_Test_Cases.pdf", "Step 13: Test Cases Specification"),
    ("docs/14_BUILD_PLAN.docx", "docs/14_BUILD_PLAN.pdf", "Step 14: Technical Build Plan"),
    ("docs/2500320140001_Project_Document.docx", "docs/2500320140001_Project_Document.pdf", "Master Project Document")
]


def convert_all():
    print("Starting Batch DOCX to PDF Conversion...")
    for docx_file, pdf_file, title in DOCX_DOCUMENTS:
        if os.path.exists(docx_file):
            convert_docx(docx_file, pdf_file, title)
        else:
            print(f"Skipping {docx_file} (not found)")


if __name__ == '__main__':
    convert_all()
