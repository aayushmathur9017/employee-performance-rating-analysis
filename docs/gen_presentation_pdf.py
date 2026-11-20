"""
Presentation PPTX to Landscape PDF Converter
Project: Employee Performance Rating Analysis (EPRA)
Student: Aayush Mathur | Enrolment No: 2500320140001 | Section: MCA-B
Institution: ABES Engineering College, Ghaziabad
"""

import os
import fitz
from pptx import Presentation

INPUT_PPTX = r"docs/2500320140001_Presentation.pptx"
OUTPUT_PDF = r"docs/2500320140001_Presentation.pdf"

# A4 Landscape dimensions in points
SLIDE_W = 842.0
SLIDE_H = 595.0
MARGIN_LEFT = 50.0
MARGIN_RIGHT = 792.0
USABLE_W = MARGIN_RIGHT - MARGIN_LEFT

NAVY_DARK   = (0.08, 0.16, 0.38)   # #142861
NAVY_LIGHT  = (0.12, 0.23, 0.54)   # #1F3B8A
BLUE_ACCENT = (0.20, 0.45, 0.85)   # #3373D9
GOLD_ACCENT = (0.90, 0.65, 0.15)   # #E6A626
WHITE       = (1.00, 1.00, 1.00)
TEXT_DARK   = (0.15, 0.18, 0.22)
TEXT_MUTED  = (0.45, 0.50, 0.55)
CARD_BG     = (0.95, 0.97, 1.00)
BORDER_COL  = (0.80, 0.85, 0.92)


def wrap_text(text: str, font: str, size: float, max_width: float):
    if not text:
        return ['']
    words = text.split()
    if not words:
        return ['']
    lines, cur = [], []
    for w in words:
        test = ' '.join(cur + [w])
        try:
            w_len = fitz.get_text_length(test, fontname=font, fontsize=size)
        except Exception:
            w_len = len(test) * size * 0.5
        if w_len <= max_width:
            cur.append(w)
        else:
            if cur:
                lines.append(' '.join(cur))
                cur = [w]
            else:
                lines.append(w)
                cur = []
    if cur:
        lines.append(' '.join(cur))
    return lines


def generate_presentation_pdf():
    if not os.path.exists(INPUT_PPTX):
        print(f"Error: {INPUT_PPTX} not found.")
        return False

    prs = Presentation(INPUT_PPTX)
    doc_pdf = fitz.open()

    for slide_idx, slide in enumerate(prs.slides):
        page = doc_pdf.new_page(width=SLIDE_W, height=SLIDE_H)

        # Slide 1 is Title Slide
        if slide_idx == 0:
            # Full Navy Banner
            page.draw_rect(fitz.Rect(0, 0, SLIDE_W, SLIDE_H), fill=NAVY_DARK, color=None)
            page.draw_rect(fitz.Rect(0, 0, SLIDE_W, 10), fill=GOLD_ACCENT, color=None)
            page.draw_rect(fitz.Rect(0, SLIDE_H - 12, SLIDE_W, SLIDE_H), fill=GOLD_ACCENT, color=None)

            page.insert_text((MARGIN_LEFT, 85), "ABES ENGINEERING COLLEGE, GHAZIABAD", fontsize=15, fontname="Helvetica-Bold", color=GOLD_ACCENT)
            page.insert_text((MARGIN_LEFT, 110), "DEPARTMENT OF MASTER OF COMPUTER APPLICATIONS | MCA HCL TRAINING", fontsize=11, fontname="helv", color=(0.8, 0.88, 1))

            page.draw_line(fitz.Point(MARGIN_LEFT, 130), fitz.Point(SLIDE_W - MARGIN_LEFT, 130), color=BLUE_ACCENT, width=1.5)

            page.insert_text((MARGIN_LEFT, 190), "Employee Performance Rating Analysis", fontsize=28, fontname="Helvetica-Bold", color=WHITE)
            page.insert_text((MARGIN_LEFT, 225), "Enterprise Workforce Intelligence, Longitudinal Calibration & Predictive Analytics Platform", fontsize=13, fontname="helv", color=(0.85, 0.92, 1))

            # Attribution Card
            page.draw_rect(fitz.Rect(MARGIN_LEFT, 275, SLIDE_W - MARGIN_LEFT, 490), fill=NAVY_LIGHT, color=BLUE_ACCENT, width=1.0)
            page.insert_text((MARGIN_LEFT + 25, 315), "ACADEMIC CAPSTONE DEFENSE & PROJECT SPECIFICATION", fontsize=12, fontname="Helvetica-Bold", color=GOLD_ACCENT)
            
            meta_items = [
                ("Candidate Name", "Aayush Mathur"),
                ("University Enrolment No", "2500320140001"),
                ("Academic Program & Section", "Master of Computer Applications (MCA) | Section: MCA-B"),
                ("Institution", "ABES Engineering College, Ghaziabad"),
                ("Curriculum Track", "MCA HCL Training / Major Capstone Project (2025-2026)"),
                ("Production URL", "https://employee-performance-rating-analysis.streamlit.app"),
                ("GitHub Codebase", "https://github.com/aayushmathur9017/employee-performance-rating-analysis")
            ]
            
            y_attr = 345
            for label, val in meta_items:
                page.insert_text((MARGIN_LEFT + 25, y_attr), f"{label}:", fontsize=9.5, fontname="Helvetica-Bold", color=(0.75, 0.85, 1))
                page.insert_text((MARGIN_LEFT + 210, y_attr), val, fontsize=9.5, fontname="helv", color=WHITE)
                y_attr += 18

            page.insert_text((MARGIN_LEFT, SLIDE_H - 25), "Confidential - For Academic Evaluation Only", fontsize=8.5, fontname="helv", color=(0.6, 0.7, 0.85))
            page.insert_text((SLIDE_W - MARGIN_LEFT - 100, SLIDE_H - 25), "Slide 1 of 11", fontsize=8.5, fontname="Helvetica-Bold", color=GOLD_ACCENT)
            continue

        # Header for Slides 2 to 11
        page.draw_rect(fitz.Rect(0, 0, SLIDE_W, 65), fill=NAVY_DARK, color=None)
        page.draw_rect(fitz.Rect(0, 65, SLIDE_W, 68), fill=BLUE_ACCENT, color=None)

        page.insert_text((MARGIN_LEFT, 30), "ABES ENGINEERING COLLEGE  |  MCA HCL TRAINING", fontsize=9, fontname="helv", color=(0.8, 0.88, 1))
        page.insert_text((SLIDE_W - MARGIN_LEFT - 180, 30), "Aayush Mathur (2500320140001, MCA-B)", fontsize=9, fontname="helv", color=(0.8, 0.88, 1))

        # Extract shapes
        text_frames = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                t = shape.text_frame.text.strip()
                if t:
                    text_frames.append(t)

        slide_title = text_frames[1] if len(text_frames) > 1 else f"Slide {slide_idx+1}"
        page.insert_text((MARGIN_LEFT, 52), slide_title[:75], fontsize=15, fontname="Helvetica-Bold", color=WHITE)

        # Content Cards Layout
        y_content = 90.0
        remaining_texts = text_frames[3:] if len(text_frames) > 3 else text_frames[2:]

        # Two-column layout if multiple blocks
        if len(remaining_texts) >= 4:
            half_w = (USABLE_W - 20) / 2
            col1 = remaining_texts[:len(remaining_texts)//2 + len(remaining_texts)%2]
            col2 = remaining_texts[len(remaining_texts)//2 + len(remaining_texts)%2:]

            for col_idx, col_blocks in enumerate([col1, col2]):
                x_col = MARGIN_LEFT + col_idx * (half_w + 20)
                y_col = y_content

                for b in col_blocks:
                    b_lines = b.split('\n')
                    header = b_lines[0]
                    body = '\n'.join(b_lines[1:]).strip()

                    card_h = 24.0 + (len(body.split('\n')) * 13.0) + 16.0
                    card_h = min(card_h, 200.0)

                    page.draw_rect(fitz.Rect(x_col, y_col, x_col + half_w, y_col + card_h), fill=CARD_BG, color=BORDER_COL, width=0.8)
                    page.draw_rect(fitz.Rect(x_col, y_col, x_col + half_w, y_col + 20), fill=NAVY_LIGHT, color=None)
                    page.insert_text((x_col + 8, y_col + 14), header[:50], fontsize=8.5, fontname="Helvetica-Bold", color=WHITE)

                    y_sub = y_col + 32
                    for line in body.split('\n'):
                        for wl in wrap_text(line, "helv", 8.0, half_w - 16):
                            page.insert_text((x_col + 8, y_sub), wl, fontsize=8.0, fontname="helv", color=TEXT_DARK)
                            y_sub += 11.5
                    y_col += card_h + 12
        else:
            y_cur = y_content
            for b in remaining_texts:
                b_lines = b.split('\n')
                header = b_lines[0]
                body = '\n'.join(b_lines[1:]).strip()

                card_h = 24.0 + (len(body.split('\n')) * 14.0) + 16.0
                card_h = min(card_h, 150.0)

                page.draw_rect(fitz.Rect(MARGIN_LEFT, y_cur, MARGIN_RIGHT, y_cur + card_h), fill=CARD_BG, color=BORDER_COL, width=0.8)
                page.draw_rect(fitz.Rect(MARGIN_LEFT, y_cur, MARGIN_RIGHT, y_cur + 22), fill=NAVY_LIGHT, color=None)
                page.insert_text((MARGIN_LEFT + 10, y_cur + 15), header[:80], fontsize=9.5, fontname="Helvetica-Bold", color=WHITE)

                y_sub = y_cur + 36
                for line in body.split('\n'):
                    for wl in wrap_text(line, "helv", 8.5, USABLE_W - 20):
                        page.insert_text((MARGIN_LEFT + 10, y_sub), wl, fontsize=8.5, fontname="helv", color=TEXT_DARK)
                        y_sub += 12.5
                y_cur += card_h + 14

        # Slide Footer
        page.draw_line(fitz.Point(MARGIN_LEFT, SLIDE_H - 35), fitz.Point(MARGIN_RIGHT, SLIDE_H - 35), color=BORDER_COL, width=0.5)
        page.insert_text((MARGIN_LEFT, SLIDE_H - 20), "Employee Performance Rating Analysis | MCA HCL Training Capstone Presentation", fontsize=8, fontname="helv", color=TEXT_MUTED)
        page.insert_text((SLIDE_W - MARGIN_LEFT - 70, SLIDE_H - 20), f"Slide {slide_idx+1} of {len(prs.slides)}", fontsize=8, fontname="Helvetica-Bold", color=NAVY_DARK)

    doc_pdf.save(OUTPUT_PDF)
    print(f"Generated: {OUTPUT_PDF} ({len(doc_pdf)} slides)")
    return True


if __name__ == '__main__':
    generate_presentation_pdf()
