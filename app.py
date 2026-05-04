import os
from flask import Flask, request, send_file

import json

from io import BytesIO

from reportlab.lib.pagesizes import A4

from reportlab.lib.colors import HexColor, white

from reportlab.pdfgen import canvas

from reportlab.platypus import Paragraph

from reportlab.lib.styles import ParagraphStyle

from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY

from docx import Document

from docx.shared import Pt, RGBColor, Inches

from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)

# ── Default colour scheme (can be overridden by user) ──

DEFAULT_COLOURS = {

    "executive": {

        "primary": "#1B2A4A",

        "accent": "#C9A96E",

        "text_dark": "#1A1A1A",

        "text_med": "#4A4A4A",

        "text_light": "#7A7A7A",

    },

    "creative": {

        "primary_1": "#6366F1",

        "primary_2": "#8B5CF6",

        "text_dark": "#1A1A2E",

        "text_med": "#4A4A5A",

        "text_light": "#7A7A8A",

    },

    "impact": {

        "header": "#111827",

        "accent": "#0D9488",

        "text_dark": "#111827",

        "text_med": "#4B5563",

        "text_light": "#9CA3AF",

    }

}


def generate_executive_pdf(cv_data, colours=None):

    """Generate Executive template PDF"""

    if colours is None:

        colours = DEFAULT_COLOURS["executive"]

    

    W, H = A4

    SIDEBAR_W = 190

    

    # Parse colours

    NAVY = HexColor(colours.get("primary", "#1B2A4A"))

    ACCENT = HexColor(colours.get("accent", "#C9A96E"))

    TEXT_DARK = HexColor(colours.get("text_dark", "#1A1A1A"))

    TEXT_MED = HexColor(colours.get("text_med", "#4A4A4A"))

    TEXT_LIGHT = HexColor(colours.get("text_light", "#7A7A7A"))

    

    pdf_buffer = BytesIO()

    c = canvas.Canvas(pdf_buffer, pagesize=A4)

    

    # Left sidebar

    c.setFillColor(NAVY)

    c.rect(0, 0, SIDEBAR_W, H, fill=1, stroke=0)

    

    # Photo circle

    cx, cy = SIDEBAR_W / 2, H - 70

    c.setFillColor(HexColor("#2A3F6A"))

    c.circle(cx, cy, 36, fill=1, stroke=0)

    c.setFillColor(white)

    c.setFont("Helvetica-Bold", 18)

    c.drawCentredString(cx, cy - 6, "KN")

    

    # Contact section

    sidebar_x = 18

    y = H - 130

    c.setFont("Helvetica-Bold", 9)

    c.setFillColor(ACCENT)

    c.drawString(sidebar_x, y, "CONTACT")

    c.setStrokeColor(ACCENT)

    c.setLineWidth(0.5)

    c.line(sidebar_x, y - 4, SIDEBAR_W - 18, y - 4)

    y -= 20

    

    c.setFont("Helvetica", 7.5)

    c.setFillColor(HexColor("#D0D0D0"))

    for info in [cv_data.get("email", ""), cv_data.get("phone", ""), 

                 cv_data.get("location", ""), cv_data.get("linkedin", "")]:

        if info:

            c.drawString(sidebar_x, y, info)

            y -= 13

    

    # Skills section

    y -= 12

    c.setFont("Helvetica-Bold", 9)

    c.setFillColor(ACCENT)

    c.drawString(sidebar_x, y, "SKILLS")

    c.setStrokeColor(ACCENT)

    c.line(sidebar_x, y - 4, SIDEBAR_W - 18, y - 4)

    y -= 20

    

    c.setFont("Helvetica", 7.5)

    c.setFillColor(HexColor("#D0D0D0"))

    skills = cv_data.get("skills", [])

    for skill in skills[:12]:  # Limit to 12 for sidebar

        c.drawString(sidebar_x, y, skill)

        y -= 8

        # Progress bar

        bar_w = SIDEBAR_W - 36

        c.setFillColor(HexColor("#2A3F6A"))

        c.roundRect(sidebar_x, y - 1, bar_w, 4, 2, fill=1, stroke=0)

        c.setFillColor(ACCENT)

        c.roundRect(sidebar_x, y - 1, bar_w * 0.8, 4, 2, fill=1, stroke=0)

        y -= 14

        c.setFillColor(HexColor("#D0D0D0"))

    

    # Education section

    y -= 10

    c.setFont("Helvetica-Bold", 9)

    c.setFillColor(ACCENT)

    c.drawString(sidebar_x, y, "EDUCATION")

    c.setStrokeColor(ACCENT)

    c.line(sidebar_x, y - 4, SIDEBAR_W - 18, y - 4)

    y -= 20

    

    for edu in cv_data.get("education", []):

        c.setFont("Helvetica-Bold", 7.5)

        c.setFillColor(white)

        c.drawString(sidebar_x, y, edu.get("degree", ""))

        y -= 12

        c.setFont("Helvetica", 7)

        c.setFillColor(HexColor("#A0A0A0"))

        c.drawString(sidebar_x, y, f"{edu.get('institution', '')} · {edu.get('year', '')}")

        y -= 18

    

    # Right main area

    main_x = SIDEBAR_W + 24

    main_w = W - main_x - 24

    y = H - 45

    

    # Name

    c.setFont("Helvetica-Bold", 24)

    c.setFillColor(NAVY)

    c.drawString(main_x, y, cv_data.get("name", ""))

    y -= 16

    

    # Title

    c.setFont("Helvetica", 9)

    c.setFillColor(TEXT_LIGHT)

    c.drawString(main_x, y, "Digital Product Specialist")

    y -= 24

    

    # Summary

    c.setFont("Helvetica-Bold", 10)

    c.setFillColor(NAVY)

    c.drawString(main_x, y, "SUMMARY")

    c.setStrokeColor(NAVY)

    c.setLineWidth(0.5)

    c.line(main_x, y - 4, W - 24, y - 4)

    y -= 16

    

    style = ParagraphStyle('summary', fontName='Helvetica', fontSize=8.5,

                           leading=13, textColor=TEXT_MED, alignment=TA_JUSTIFY)

    summary_text = cv_data.get("summary", "")

    if summary_text:

        p = Paragraph(summary_text, style)

        pw, ph = p.wrap(main_w, 200)

        p.drawOn(c, main_x, y - ph)

        y -= ph + 16

    

    # Experience

    c.setFont("Helvetica-Bold", 10)

    c.setFillColor(NAVY)

    c.drawString(main_x, y, "EXPERIENCE")

    c.setStrokeColor(NAVY)

    c.line(main_x, y - 4, W - 24, y - 4)

    y -= 18

    

    bullet_style = ParagraphStyle('bullet', fontName='Helvetica', fontSize=8,

                                   leading=11.5, textColor=TEXT_MED, alignment=TA_LEFT,

                                   leftIndent=10, bulletIndent=0)

    

    for job in cv_data.get("experience", []):

        if y < 100:

            c.showPage()

            y = H - 40

            c.setFillColor(NAVY)

            c.rect(0, 0, SIDEBAR_W, H, fill=1, stroke=0)

        

        # Job title & dates

        c.setFont("Helvetica-Bold", 9)

        c.setFillColor(TEXT_DARK)

        c.drawString(main_x, y, job.get("title", ""))

        c.setFont("Helvetica", 7.5)

        c.setFillColor(TEXT_LIGHT)

        tw = c.stringWidth(job.get("dates", ""), "Helvetica", 7.5)

        c.drawString(W - 24 - tw, y, job.get("dates", ""))

        y -= 12

        

        # Company

        c.setFont("Helvetica", 8)

        c.setFillColor(ACCENT)

        c.drawString(main_x, y, job.get("company", ""))

        y -= 14

        

        # Bullets

        for bullet in job.get("bullets", []):

            text = f"<bullet>•</bullet> {bullet}"

            p = Paragraph(text, bullet_style)

            pw, ph = p.wrap(main_w - 10, 200)

            if y - ph < 40:

                c.showPage()

                y = H - 40

                c.setFillColor(NAVY)

                c.rect(0, 0, SIDEBAR_W, H, fill=1, stroke=0)

            p.drawOn(c, main_x, y - ph)

            y -= ph + 3

        

        y -= 8

    

    # Footer

    c.setFont("Helvetica", 6)

    c.setFillColor(TEXT_LIGHT)

    c.drawCentredString(W / 2 + SIDEBAR_W / 2, 15, "Created with RoleAlign")

    

    c.save()

    pdf_buffer.seek(0)

    return pdf_buffer


def generate_creative_pdf(cv_data, colours=None):

    """Generate Creative template PDF"""

    if colours is None:

        colours = DEFAULT_COLOURS["creative"]

    

    W, H = A4

    PURPLE_1 = HexColor(colours.get("primary_1", "#6366F1"))

    PURPLE_2 = HexColor(colours.get("primary_2", "#8B5CF6"))

    TEXT_DARK = HexColor(colours.get("text_dark", "#1A1A2E"))

    TEXT_MED = HexColor(colours.get("text_med", "#4A4A5A"))

    TEXT_LIGHT = HexColor(colours.get("text_light", "#7A7A8A"))

    

    RIGHT_W = 185

    LEFT_W = W - RIGHT_W - 48

    

    pdf_buffer = BytesIO()

    c = canvas.Canvas(pdf_buffer, pagesize=A4)

    

    # Gradient header

    band_h = 90

    steps = 50

    for i in range(steps):

        r1, g1, b1 = 0.388, 0.400, 0.945

        r2, g2, b2 = 0.545, 0.361, 0.965

        t = i / steps

        r = r1 + (r2 - r1) * t

        g = g1 + (g2 - g1) * t

        b = b1 + (b2 - b1) * t

        stripe_h = band_h / steps

        c.setFillColor(HexColor(f"#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}"))

        c.rect(0, H - band_h + i * stripe_h, W, stripe_h + 1, fill=1, stroke=0)

    

    # Header text

    c.setFont("Helvetica-Bold", 28)

    c.setFillColor(white)

    c.drawString(28, H - 50, cv_data.get("name", ""))

    c.setFont("Helvetica", 11)

    c.setFillColor(HexColor("#E0DDFF"))

    c.drawString(28, H - 68, "Digital Product Specialist")

    

    # Right panel

    panel_x = W - RIGHT_W

    c.setFillColor(HexColor("#F0EDFF"))

    c.rect(panel_x, 0, RIGHT_W, H - band_h, fill=1, stroke=0)

    

    # Main content

    left_x = 28

    y = H - band_h - 28

    

    # Summary

    c.setFont("Helvetica-Bold", 10)

    c.setFillColor(PURPLE_1)

    c.drawString(left_x, y, "SUMMARY")

    c.setStrokeColor(PURPLE_1)

    c.setLineWidth(1.5)

    c.line(left_x, y - 4, left_x + 45, y - 4)

    y -= 18

    

    style = ParagraphStyle('summary', fontName='Helvetica', fontSize=8.5,

                           leading=13, textColor=TEXT_MED, alignment=TA_JUSTIFY)

    summary_text = cv_data.get("summary", "")

    if summary_text:

        p = Paragraph(summary_text, style)

        pw, ph = p.wrap(LEFT_W, 200)

        p.drawOn(c, left_x, y - ph)

        y -= ph + 20

    

    # Experience

    c.setFont("Helvetica-Bold", 10)

    c.setFillColor(PURPLE_1)

    c.drawString(left_x, y, "EXPERIENCE")

    c.setStrokeColor(PURPLE_1)

    c.setLineWidth(1.5)

    c.line(left_x, y - 4, left_x + 60, y - 4)

    y -= 18

    

    bullet_style = ParagraphStyle('bullet', fontName='Helvetica', fontSize=8,

                                   leading=11, textColor=TEXT_MED, leftIndent=12, bulletIndent=0)

    

    for job in cv_data.get("experience", []):

        if y < 80:

            c.showPage()

            y = H - 40

            c.setFillColor(HexColor("#F0EDFF"))

            c.rect(panel_x, 0, RIGHT_W, H, fill=1, stroke=0)

        

        c.setFont("Helvetica-Bold", 9)

        c.setFillColor(TEXT_DARK)

        c.drawString(left_x, y, job.get("title", ""))

        c.setFont("Helvetica", 7.5)

        c.setFillColor(TEXT_LIGHT)

        tw = c.stringWidth(job.get("dates", ""), "Helvetica", 7.5)

        c.drawString(left_x + LEFT_W - tw, y, job.get("dates", ""))

        y -= 12

        

        c.setFont("Helvetica", 8)

        c.setFillColor(PURPLE_1)

        c.drawString(left_x, y, job.get("company", ""))

        y -= 14

        

        for bullet in job.get("bullets", []):

            c.setFillColor(PURPLE_2)

            c.circle(left_x + 4, y + 2.5, 2, fill=1, stroke=0)

            

            text = bullet

            p = Paragraph(text, bullet_style)

            pw, ph = p.wrap(LEFT_W - 14, 200)

            if y - ph < 40:

                c.showPage()

                y = H - 40

                c.setFillColor(HexColor("#F0EDFF"))

                c.rect(panel_x, 0, RIGHT_W, H, fill=1, stroke=0)

            p.drawOn(c, left_x, y - ph + 3)

            y -= ph + 3

        

        y -= 10

    

    # Right column skills

    rx = panel_x + 14

    rw = RIGHT_W - 28

    ry = H - band_h - 28

    

    # Photo circle

    photo_cx = panel_x + RIGHT_W / 2

    photo_cy = ry + 5

    c.setFillColor(HexColor("#DDD8FF"))

    c.circle(photo_cx, photo_cy, 32, fill=1, stroke=0)

    c.setFillColor(PURPLE_1)

    c.setFont("Helvetica-Bold", 16)

    c.drawCentredString(photo_cx, photo_cy - 5, "KN")

    ry -= 50

    

    # Skills

    c.setFont("Helvetica-Bold", 9)

    c.setFillColor(PURPLE_1)

    c.drawString(rx, ry, "SKILLS")

    c.setStrokeColor(PURPLE_1)

    c.setLineWidth(1.5)

    c.line(rx, ry - 4, rx + 35, ry - 4)

    ry -= 20

    

    pill_x = rx

    for skill in cv_data.get("skills", []):

        c.setFont("Helvetica", 7)

        tw = c.stringWidth(skill, "Helvetica", 7) + 14

        if pill_x + tw > panel_x + RIGHT_W - 14:

            pill_x = rx

            ry -= 22

        

        c.setFillColor(HexColor("#EDE9FE"))

        c.roundRect(pill_x, ry - 4, tw, 16, 8, fill=1, stroke=0)

        c.setFillColor(PURPLE_1)

        c.drawString(pill_x + 7, ry + 2, skill)

        pill_x += tw + 6

    

    ry -= 36

    

    # Education

    c.setFont("Helvetica-Bold", 9)

    c.setFillColor(PURPLE_1)

    c.drawString(rx, ry, "EDUCATION")

    c.setStrokeColor(PURPLE_1)

    c.setLineWidth(1.5)

    c.line(rx, ry - 4, rx + 58, ry - 4)

    ry -= 20

    

    for edu in cv_data.get("education", []):

        c.setFont("Helvetica-Bold", 8)

        c.setFillColor(TEXT_DARK)

        c.drawString(rx, ry, edu.get("degree", ""))

        ry -= 12

        c.setFont("Helvetica", 7)

        c.setFillColor(TEXT_LIGHT)

        c.drawString(rx, ry, f"{edu.get('institution', '')} · {edu.get('year', '')}")

        ry -= 18

    

    # Footer

    c.setFont("Helvetica", 6)

    c.setFillColor(TEXT_LIGHT)

    c.drawCentredString(W / 2, 12, "Created with RoleAlign")

    

    c.save()

    pdf_buffer.seek(0)

    return pdf_buffer


def generate_impact_pdf(cv_data, colours=None):

    """Generate Impact template PDF"""

    if colours is None:

        colours = DEFAULT_COLOURS["impact"]

    

    W, H = A4

    HEADER_BG = HexColor(colours.get("header", "#111827"))

    TEAL = HexColor(colours.get("accent", "#0D9488"))

    TEXT_DARK = HexColor(colours.get("text_dark", "#111827"))

    TEXT_MED = HexColor(colours.get("text_med", "#4B5563"))

    TEXT_LIGHT = HexColor(colours.get("text_light", "#9CA3AF"))

    

    RIGHT_W = 180

    

    pdf_buffer = BytesIO()

    c = canvas.Canvas(pdf_buffer, pagesize=A4)

    

    # Dark header

    header_h = 100

    c.setFillColor(HEADER_BG)

    c.rect(0, H - header_h, W, header_h, fill=1, stroke=0)

    

    c.setStrokeColor(TEAL)

    c.setLineWidth(3)

    c.line(0, H - header_h, W, H - header_h)

    

    # Name

    c.setFont("Helvetica-Bold", 30)

    c.setFillColor(white)

    c.drawString(28, H - 48, cv_data.get("name", ""))

    

    # Title

    c.setFont("Helvetica", 11)

    c.setFillColor(TEXT_LIGHT)

    c.drawString(28, H - 66, "Digital Product Specialist")

    

    # Right panel

    right_x = W - RIGHT_W - 4

    c.setFillColor(HexColor("#F9FAFB"))

    c.rect(right_x - 10, 0, RIGHT_W + 14, H - header_h + 22, fill=1, stroke=0)

    

    # Main content

    left_x = 28

    left_w = W - RIGHT_W - 52

    body_top = H - header_h - 22

    y = body_top

    

    # Summary

    c.setFont("Helvetica-Bold", 10)

    c.setFillColor(TEXT_DARK)

    c.drawString(left_x, y, "SUMMARY")

    c.setFillColor(TEAL)

    c.rect(left_x, y - 6, 30, 2.5, fill=1, stroke=0)

    y -= 18

    

    style = ParagraphStyle('summary', fontName='Helvetica', fontSize=8.5,

                           leading=13, textColor=TEXT_MED, alignment=TA_JUSTIFY)

    summary_text = cv_data.get("summary", "")

    if summary_text:

        p = Paragraph(summary_text, style)

        pw, ph = p.wrap(left_w, 200)

        p.drawOn(c, left_x, y - ph)

        y -= ph + 20

    

    # Experience

    c.setFont("Helvetic-Bold", 10)

    c.setFillColor(TEXT_DARK)

    c.drawString(left_x, y, "EXPERIENCE")

    c.setFillColor(TEAL)

    c.rect(left_x, y - 6, 50, 2.5, fill=1, stroke=0)

    y -= 18

    

    timeline_x = left_x + 6

    bullet_style = ParagraphStyle('bullet', fontName='Helvetica', fontSize=8,

                                   leading=11, textColor=TEXT_MED, leftIndent=12, bulletIndent=0)

    

    for i, job in enumerate(cv_data.get("experience", [])):

        if y < 80:

            c.showPage()

            y = H - 40

            c.setFillColor(HexColor("#F9FAFB"))

            c.rect(right_x - 10, 0, RIGHT_W + 14, H, fill=1, stroke=0)

        

        # Timeline dot

        c.setFillColor(TEAL)

        c.circle(timeline_x, y + 2, 4, fill=1, stroke=0)

        c.setFillColor(white)

        c.circle(timeline_x, y + 2, 2, fill=1, stroke=0)

        

        # Title

        c.setFont("Helvetica-Bold", 9)

        c.setFillColor(TEXT_DARK)

        c.drawString(left_x + 18, y, job.get("title", ""))

        

        # Dates

        c.setFont("Helvetica", 7.5)

        c.setFillColor(TEXT_LIGHT)

        tw = c.stringWidth(job.get("dates", ""), "Helvetica", 7.5)

        c.drawString(left_x + left_w - tw, y, job.get("dates", ""))

        y -= 12

        

        # Company

        c.setFont("Helvetica", 8)

        c.setFillColor(TEAL)

        c.drawString(left_x + 18, y, job.get("company", ""))

        y -= 14

        

        # Bullets

        for bullet in job.get("bullets", []):

            text = bullet

            p = Paragraph(text, bullet_style)

            pw, ph = p.wrap(left_w - 24, 200)

            if y - ph < 40:

                c.showPage()

                y = H - 40

                c.setFillColor(HexColor("#F9FAFB"))

                c.rect(right_x - 10, 0, RIGHT_W + 14, H, fill=1, stroke=0)

            

            c.setFillColor(HexColor("#D1D5DB"))

            c.circle(left_x + 21, y + 2.5, 1.5, fill=1, stroke=0)

            p.drawOn(c, left_x + 18, y - ph + 3)

            y -= ph + 3

        

        # Timeline line

        if i < len(cv_data.get("experience", [])) - 1:

            c.setStrokeColor(HexColor("#E5E7EB"))

            c.setLineWidth(1)

            c.line(timeline_x, y + 3, timeline_x, y - 6)

        

        y -= 10

    

    # Right column

    ry = body_top

    rx = right_x + 6

    rw = RIGHT_W - 10

    

    # Skills

    c.setFont("Helvetica-Bold", 9)

    c.setFillColor(TEXT_DARK)

    c.drawString(rx, ry, "SKILLS")

    c.setFillColor(TEAL)

    c.rect(rx, ry - 6, 30, 2.5, fill=1, stroke=0)

    ry -= 22

    

    tag_x = rx

    for skill in cv_data.get("skills", []):

        c.setFont("Helvetica", 7)

        tw = c.stringWidth(skill, "Helvetica", 7) + 12

        th = 17

        

        if tag_x + tw > right_x + RIGHT_W - 4:

            tag_x = rx

            ry -= 22

        

        c.setFillColor(HEADER_BG)

        c.roundRect(tag_x, ry - 4, tw, th, 4, fill=1, stroke=0)

        c.setFillColor(white)

        c.drawString(tag_x + 6, ry + 2, skill)

        

        tag_x += tw + 5

    

    ry -= 38

    

    # Education

    c.setFont("Helvetica-Bold", 9)

    c.setFillColor(TEXT_DARK)

    c.drawString(rx, ry, "EDUCATION")

    c.setFillColor(TEAL)

    c.rect(rx, ry - 6, 50, 2.5, fill=1, stroke=0)

    ry -= 22

    

    for edu in cv_data.get("education", []):

        c.setFont("Helvetica-Bold", 8)

        c.setFillColor(TEXT_DARK)

        c.drawString(rx, ry, edu.get("degree", ""))

        ry -= 12

        c.setFont("Helvetica", 7)

        c.setFillColor(TEXT_LIGHT)

        c.drawString(rx, ry, f"{edu.get('institution', '')} · {edu.get('year', '')}")

        ry -= 18

    

    # Footer

    c.setFont("Helvetica", 6)

    c.setFillColor(TEXT_LIGHT)

    c.drawCentredString(W / 2, 12, "Created with RoleAlign")

    

    c.save()

    pdf_buffer.seek(0)

    return pdf_buffer


def generate_docx(cv_data):

    """Generate a clean DOCX file"""

    doc = Document()

    

    # Title

    title = doc.add_paragraph()

    title_run = title.add_run(cv_data.get("name", ""))

    title_run.font.size = Pt(24)

    title_run.font.bold = True

    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    

    # Contact info

    contact = doc.add_paragraph()

    contact_text = " | ".join([cv_data.get("email", ""), cv_data.get("phone", ""), 

                               cv_data.get("location", "")])

    contact_run = contact.add_run(contact_text)

    contact_run.font.size = Pt(9)

    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER

    

    doc.add_paragraph()  # Spacing

    

    # Summary

    summary_heading = doc.add_heading("SUMMARY", level=2)

    summary_heading.style.font.size = Pt(12)

    doc.add_paragraph(cv_data.get("summary", ""))

    

    # Experience

    exp_heading = doc.add_heading("EXPERIENCE", level=2)

    exp_heading.style.font.size = Pt(12)

    

    for job in cv_data.get("experience", []):

        # Job title and company

        job_title = doc.add_paragraph()

        title_run = job_title.add_run(f"{job.get('title', '')} – {job.get('company', '')}")

        title_run.bold = True

        

        # Dates

        date_para = doc.add_paragraph(job.get("dates", ""))

        date_para.style.font.size = Pt(9)

        

        # Bullets

        for bullet in job.get("bullets", []):

            doc.add_paragraph(bullet, style='List Bullet')

    

    # Skills

    skills_heading = doc.add_heading("SKILLS", level=2)

    skills_heading.style.font.size = Pt(12)

    skills_para = doc.add_paragraph()

    skills_para.add_run(", ".join(cv_data.get("skills", [])))

    

    # Education

    edu_heading = doc.add_heading("EDUCATION", level=2)

    edu_heading.style.font.size = Pt(12)

    

    for edu in cv_data.get("education", []):

        edu_para = doc.add_paragraph()

        edu_run = edu_para.add_run(edu.get("degree", ""))

        edu_run.bold = True

        edu_para.add_run(f" – {edu.get('institution', '')} ({edu.get('year', '')})")

    

    docx_buffer = BytesIO()

    doc.save(docx_buffer)

    docx_buffer.seek(0)

    return docx_buffer


@app.route('/', methods=['GET'])

def home():

    return {

        "status": "RoleAlign PDF API is running",

        "endpoints": {

            "POST /generate-pdf": "Generate a CV PDF",

            "POST /generate-docx": "Generate a CV DOCX file"

        }

    }


@app.route('/generate-pdf', methods=['POST'])

def generate_pdf():

    try:

        data = request.get_json()

        

        # Extract request data

        cv_data = data.get('cv_data', {})

        template = data.get('template', 'executive')  # executive, creative, or impact

        colours = data.get('colours', None)

        

        # Generate PDF based on template

        if template == 'executive':

            pdf_buffer = generate_executive_pdf(cv_data, colours)

        elif template == 'creative':

            pdf_buffer = generate_creative_pdf(cv_data, colours)

        elif template == 'impact':

            pdf_buffer = generate_impact_pdf(cv_data, colours)

        else:

            return {"error": "Invalid template. Use 'executive', 'creative', or 'impact'"}, 400

        

        # Return PDF

        return send_file(

            pdf_buffer,

            mimetype='application/pdf',

            as_attachment=True,

            download_name=f'CV_{template}_{cv_data.get("name", "CV").replace(" ", "_")}.pdf'

        )

    

    except Exception as e:

        return {"error": str(e)}, 500


@app.route('/generate-docx', methods=['POST'])

def generate_docx_route():

    try:

        data = request.get_json()

        cv_data = data.get('cv_data', {})

        

        docx_buffer = generate_docx(cv_data)

        

        return send_file(

            docx_buffer,

            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',

            as_attachment=True,

            download_name=f'CV_{cv_data.get("name", "CV").replace(" ", "_")}.docx'

        )

    

    except Exception as e:

        return {"error": str(e)}, 500


if __name__ == '__main__':

    app.run(debug=False, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))

Wait — I need to add one line at the top. Open app.py and add this to the very beginning:

import os

So the imports section should look like:

import os

from flask import Flask, request, send_file
