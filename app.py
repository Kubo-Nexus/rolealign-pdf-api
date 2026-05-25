import os
import re
import base64
import urllib.request
from io import BytesIO
from xml.sax.saxutils import escape

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor, white
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
CORS(app)

W, H = A4
API_VERSION = "1.2.4"

ACRONYMS = {
    "sap": "SAP", "sd": "SD", "mm": "MM", "wh": "WH", "fi": "FI",
    "sop": "SOP", "sops": "SOPs", "kpi": "KPI", "kpis": "KPIs",
    "ats": "ATS", "ai": "AI", "cv": "CV", "api": "API", "apis": "APIs",
    "crm": "CRM", "sql": "SQL", "hr": "HR", "it": "IT", "qa": "QA",
}


def safe_hex(colour_str, fallback):
    try:
        if colour_str and isinstance(colour_str, str) and colour_str.startswith("#") and len(colour_str) in (4, 7):
            return HexColor(colour_str)
    except Exception:
        pass
    return HexColor(fallback)


def as_colour(value, fallback="#4A4A4A"):
    if hasattr(value, "red"):
        return value
    return safe_hex(value, fallback)


def clean_text(value):
    if value is None:
        return ""
    text = str(value).strip()
    text = text.replace("—", "–")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\s+-\s*$", "", text)
    text = re.sub(r"\s+\(\s*\)\s*$", "", text)

    text = re.sub(
        r"SAP\s+Super\s+User\s*\(\s*SD\s*,?\s*MM\s*,?\s*WH\s*&\s*FI\s*\)",
        "SAP Super User (SD, MM, WH & FI)",
        text,
        flags=re.IGNORECASE,
    )
    text = text.replace("Sd, Mm, Wh & Fi", "SD, MM, WH & FI")
    text = text.replace("Sd Mm Wh & Fi", "SD, MM, WH & FI")
    text = text.replace("SD MM WH & FI", "SD, MM, WH & FI")

    for wrong, right in ACRONYMS.items():
        text = re.sub(rf"\b{wrong}\b", right, text, flags=re.IGNORECASE)

    return text.strip()


def clean_join(parts, sep=" - "):
    cleaned = [clean_text(p) for p in parts if clean_text(p)]
    return sep.join(cleaned)


def normalise_cv(cv):
    cv = cv or {}
    name = clean_text(cv.get("name") or cv.get("candidate_name") or cv.get("full_name") or "Candidate")

    clean_cv = {
        "name": name,
        "email": clean_text(cv.get("email")),
        "phone": clean_text(cv.get("phone")),
        "location": clean_text(cv.get("location")),
        "linkedin": clean_text(cv.get("linkedin")),
        "summary": clean_text(cv.get("summary")),
        "photo": cv.get("photo") or cv.get("profile_photo") or cv.get("photo_url") or cv.get("profileImageUrl"),
        "is_premium": bool(cv.get("is_premium", True)),
        "render_version": cv.get("render_version") or API_VERSION,
    }

    skills = cv.get("skills") or []
    if isinstance(skills, str):
        skills = re.split(r",|\n|•|·", skills)

    clean_skills = []
    seen = set()
    for skill in skills:
        s = clean_text(skill)
        if s and s.lower() not in seen:
            clean_skills.append(s)
            seen.add(s.lower())

    clean_cv["skills"] = clean_skills
    clean_cv["skills_csv"] = ", ".join(clean_skills)

    experience = []
    for job in cv.get("experience", []) or []:
        if not isinstance(job, dict):
            continue

        title = clean_text(job.get("title") or job.get("role"))
        company = clean_text(job.get("company") or job.get("employer"))
        dates = clean_text(job.get("dates") or job.get("date") or job.get("period"))
        bullets = [clean_text(b) for b in (job.get("bullets") or job.get("responsibilities") or []) if clean_text(b)]

        header = clean_text(job.get("header"))
        if not header:
            header = clean_join([clean_join([title, company], " - "), dates], " | ")

        if title or company or dates or bullets:
            experience.append({
                "title": title,
                "company": company,
                "dates": dates,
                "header": header,
                "bullets": bullets,
            })

    clean_cv["experience"] = experience

    education = []
    for edu in cv.get("education", []) or []:
        if isinstance(edu, str):
            degree = clean_text(edu)
            institution = ""
            year = ""
        elif isinstance(edu, dict):
            degree = clean_text(edu.get("degree") or edu.get("qualification") or edu.get("name"))
            institution = clean_text(edu.get("institution") or edu.get("school") or edu.get("provider"))
            year = clean_text(edu.get("year") or edu.get("dates") or edu.get("date"))
        else:
            continue

        line = clean_join([degree, institution], " - ")
        if year:
            line = f"{line} ({year})" if line else year

        if line and line.lower() not in {"professional certifications", "professional certification"}:
            education.append({
                "degree": degree,
                "institution": institution,
                "year": year,
                "line": line,
            })

    clean_cv["education"] = education

    certifications = []
    for cert in cv.get("certifications", []) or []:
        cert_text = clean_text(cert.get("name") if isinstance(cert, dict) else cert)
        if cert_text and cert_text.lower() not in {"professional certifications", "professional certification"}:
            certifications.append(cert_text)

    clean_cv["certifications"] = certifications
    return clean_cv


def decode_photo(photo_data):
    if not photo_data:
        return None

    try:
        if isinstance(photo_data, str) and photo_data.startswith("http"):
            req = urllib.request.Request(photo_data, headers={"User-Agent": "RoleAlignPDF/1.2.4"})
            with urllib.request.urlopen(req, timeout=8) as response:
                return ImageReader(BytesIO(response.read()))

        data = str(photo_data)
        if "," in data:
            data = data.split(",", 1)[1]

        return ImageReader(BytesIO(base64.b64decode(data)))
    except Exception as e:
        print(f"[photo_error] {e}")
        return None


def draw_circular_photo(c, img_reader, cx, cy, radius):
    if img_reader is None:
        return

    c.saveState()
    path = c.beginPath()
    path.circle(cx, cy, radius)
    path.close()
    c.clipPath(path, stroke=0)
    c.drawImage(
        img_reader,
        cx - radius,
        cy - radius,
        radius * 2,
        radius * 2,
        preserveAspectRatio=True,
        mask="auto",
    )
    c.restoreState()


def draw_rolealign_watermark(c):
    c.saveState()
    try:
        c.setFillColor(HexColor("#D9D9D9"))
        c.setFillAlpha(0.18)
    except Exception:
        c.setFillColor(HexColor("#E6E6E6"))

    c.setFont("Helvetica-Bold", 58)
    c.translate(W / 2, H / 2)
    c.rotate(35)
    c.drawCentredString(0, 0, "RoleAlign")
    c.drawCentredString(0, 110, "RoleAlign")
    c.drawCentredString(0, -110, "RoleAlign")
    c.restoreState()


def footer_brand(c, is_premium=True):
    if not is_premium:
        draw_rolealign_watermark(c)
        c.setFont("Helvetica", 6)
        c.setFillColor(HexColor("#888888"))
        c.drawCentredString(W / 2, 12, "Created with RoleAlign")


def section_heading(c, x, y, label, colour, width=70):
    c.setFont("Helvetica-Bold", 10)
    c.setFillColor(as_colour(colour))
    c.drawString(x, y, label)
    c.setStrokeColor(as_colour(colour))
    c.setLineWidth(0.8)
    c.line(x, y - 4, x + width, y - 4)


def paragraph_height(text, width, font="Helvetica", size=8, leading=11, colour="#4A4A4A", bold=False, alignment=TA_LEFT):
    style = ParagraphStyle(
        "p",
        fontName="Helvetica-Bold" if bold else font,
        fontSize=size,
        leading=leading,
        textColor=as_colour(colour),
        alignment=alignment,
    )
    p = Paragraph(escape(clean_text(text)), style)
    _, h = p.wrap(width, 800)
    return p, h


def draw_wrapped(c, text, x, y, width, font="Helvetica", size=8, leading=11, colour="#4A4A4A", bold=False, alignment=TA_LEFT):
    p, h = paragraph_height(text, width, font, size, leading, colour, bold, alignment)
    p.drawOn(c, x, y - h)
    return y - h


def draw_list_lines(c, items, x, y, width, colour="#4A4A4A", size=7.4, leading=9.4, gap=4, bold=False, max_y=35):
    for item in items:
        item = clean_text(item)
        if not item:
            continue

        p, h = paragraph_height(item, width, size=size, leading=leading, colour=colour, bold=bold)
        if y - h < max_y:
            break

        p.drawOn(c, x, y - h)
        y -= h + gap

    return y


def draw_role(c, job, x, y, width, accent, text_dark="#1A1A1A", text_med="#4A4A4A", text_light="#777777", bullet=True):
    title = clean_text(job.get("title"))
    company = clean_text(job.get("company"))
    dates = clean_text(job.get("dates"))

    date_w = 82
    title_w = max(120, width - date_w - 10)

    title_y = draw_wrapped(
        c,
        title,
        x,
        y,
        title_w,
        size=8.8,
        leading=10.7,
        colour=text_dark,
        bold=True,
    )

    if dates:
        c.setFont("Helvetica", 7.5)
        c.setFillColor(as_colour(text_light))
        tw = c.stringWidth(dates, "Helvetica", 7.5)
        c.drawString(x + width - tw, y, dates)

    y = min(title_y, y - 12)

    if company:
        c.setFont("Helvetica", 8)
        c.setFillColor(as_colour(accent))
        c.drawString(x, y, company)
        y -= 13

    bstyle = ParagraphStyle(
        "bullet",
        fontName="Helvetica",
        fontSize=8,
        leading=11,
        textColor=as_colour(text_med),
        leftIndent=10,
        bulletIndent=0,
    )

    for item in job.get("bullets", []) or []:
        item = clean_text(item)
        if not item:
            continue

        txt = f"<bullet>&#8226;</bullet> {escape(item)}" if bullet else escape(item)
        p = Paragraph(txt, bstyle)
        _, h = p.wrap(width - 4, 500)
        p.drawOn(c, x, y - h)
        y -= h + 3

    return y - 12


def draw_education(c, education, x, y, width, title_colour="#1A1A1A", size=7.2, leading=9, max_y=35):
    for edu in education:
        line = clean_text(edu.get("line"))
        if not line:
            continue

        p, h = paragraph_height(line, width, size=size, leading=leading, colour=title_colour)
        if y - h < max_y:
            break

        p.drawOn(c, x, y - h)
        y -= h + 7

    return y


def generate_starter_pdf(cv, colours):
    cv = normalise_cv(cv)
    cv["is_premium"] = False

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    x = 42
    y = H - 45
    width = W - 84
    accent = safe_hex(colours.get("accent"), "#111827")

    footer_brand(c, False)

    c.setFont("Helvetica-Bold", 22)
    c.setFillColor(HexColor("#111827"))
    c.drawString(x, y, cv["name"])
    y -= 16

    c.setFont("Helvetica", 8)
    c.setFillColor(HexColor("#4B5563"))
    c.drawString(x, y, clean_join([cv.get("email"), cv.get("phone"), cv.get("location")], " | "))
    y -= 28

    section_heading(c, x, y, "SUMMARY", accent, 65)
    y -= 16
    y = draw_wrapped(c, cv.get("summary"), x, y, width, size=8.5, leading=12, colour="#374151", alignment=TA_JUSTIFY) - 14

    section_heading(c, x, y, "EXPERIENCE", accent, 80)
    y -= 18

    for job in cv.get("experience", []):
        if y < 95:
            c.showPage()
            footer_brand(c, False)
            y = H - 45

        y = draw_role(c, job, x, y, width, accent, bullet=False)

    if y < 100:
        c.showPage()
        footer_brand(c, False)
        y = H - 45

    section_heading(c, x, y, "SKILLS", accent, 45)
    y -= 16
    y = draw_wrapped(c, cv.get("skills_csv"), x, y, width, size=8, leading=11, colour="#374151") - 14

    if y < 100:
        c.showPage()
        footer_brand(c, False)
        y = H - 45

    section_heading(c, x, y, "EDUCATION", accent, 70)
    y -= 16
    draw_education(c, cv.get("education", []), x, y, width)

    c.save()
    buf.seek(0)
    return buf


def generate_executive_pdf(cv, colours):
    cv = normalise_cv(cv)

    SIDEBAR_W = 190
    NAVY = safe_hex(colours.get("primary"), "#1B2A4A")
    ACCENT = safe_hex(colours.get("accent"), "#C9A96E")

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    def draw_sidebar():
        c.setFillColor(NAVY)
        c.rect(0, 0, SIDEBAR_W, H, fill=1, stroke=0)
        footer_brand(c, cv.get("is_premium", True))

    draw_sidebar()

    sx = 18
    cx = SIDEBAR_W / 2
    cy = H - 70
    photo_img = decode_photo(cv.get("photo"))

    if photo_img:
        c.setFillColor(white)
        c.circle(cx, cy, 38, fill=1, stroke=0)
        draw_circular_photo(c, photo_img, cx, cy, 36)
    else:
        c.setFillColor(HexColor("#2A3F6A"))
        c.circle(cx, cy, 36, fill=1, stroke=0)
        c.setFillColor(white)
        c.setFont("Helvetica-Bold", 18)
        initials = "".join([w[0] for w in cv["name"].split()[:2]]).upper()
        c.drawCentredString(cx, cy - 6, initials)

    y = H - 130

    section_heading(c, sx, y, "CONTACT", ACCENT, SIDEBAR_W - 36)
    y -= 20
    y = draw_list_lines(
        c,
        [cv.get("email"), cv.get("phone"), cv.get("location"), cv.get("linkedin")],
        sx,
        y,
        SIDEBAR_W - 36,
        "#D0D0D0",
        7.1,
        9,
        3,
    )

    y -= 10

    section_heading(c, sx, y, "SKILLS", ACCENT, SIDEBAR_W - 36)
    y -= 20
    y = draw_list_lines(c, cv.get("skills", []), sx, y, SIDEBAR_W - 36, "#D0D0D0", 6.8, 8.2, 3)

    y -= 8

    section_heading(c, sx, y, "EDUCATION", ACCENT, SIDEBAR_W - 36)
    y -= 18
    draw_education(c, cv.get("education", []), sx, y, SIDEBAR_W - 36, "#FFFFFF", 6.7, 8.2)

    mx = SIDEBAR_W + 24
    mw = W - SIDEBAR_W - 48
    y = H - 45

    c.setFont("Helvetica-Bold", 24)
    c.setFillColor(NAVY)
    c.drawString(mx, y, cv["name"])
    y -= 26

    section_heading(c, mx, y, "SUMMARY", NAVY, mw)
    y -= 16
    y = draw_wrapped(c, cv.get("summary"), mx, y, mw, size=8.5, leading=13, colour="#4A4A4A", alignment=TA_JUSTIFY) - 16

    section_heading(c, mx, y, "EXPERIENCE", NAVY, mw)
    y -= 18

    for job in cv.get("experience", []):
        if y < 90:
            c.showPage()
            draw_sidebar()
            y = H - 45

        y = draw_role(c, job, mx, y, mw, ACCENT, "#1A1A1A", "#4A4A4A", "#7A7A7A")

    c.save()
    buf.seek(0)
    return buf


def draw_creative_page1_sidebar(c, cv, panel_x, right_w, top_y, p1):
    rx = panel_x + 14
    panel_w = right_w - 28
    y = top_y

    photo_img = decode_photo(cv.get("photo"))
    if photo_img:
        c.setFillColor(white)
        c.circle(panel_x + right_w / 2, y - 8, 34, fill=1, stroke=0)
        draw_circular_photo(c, photo_img, panel_x + right_w / 2, y - 8, 32)
        y -= 62

    section_heading(c, rx, y, "SKILLS", p1, 35)
    y -= 18

    y = draw_list_lines(
        c,
        cv.get("skills", []),
        rx,
        y,
        panel_w,
        "#4A4A5A",
        6.8,
        8.2,
        3,
        max_y=95,
    )

    y -= 8

    if y > 125:
        section_heading(c, rx, y, "EDUCATION", p1, 58)
        y -= 18
        draw_education(c, cv.get("education", []), rx, y, panel_w, "#1A1A2E", 6.6, 8, max_y=35)


def generate_creative_pdf(cv, colours):
    cv = normalise_cv(cv)

    P1 = safe_hex(colours.get("primary_1", colours.get("primary")), "#6366F1")
    RIGHT_W = 185
    LEFT_W = W - RIGHT_W - 62
    PANEL_BG = HexColor("#F0EDFF")

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    band_h = 90
    panel_x = W - RIGHT_W

    def draw_page_bg(page1=False):
        footer_brand(c, cv.get("is_premium", True))

        if page1:
            c.setFillColor(P1)
            c.rect(0, H - band_h, W, band_h, fill=1, stroke=0)

            c.setFont("Helvetica-Bold", 28)
            c.setFillColor(white)
            c.drawString(28, H - 50, cv["name"])

            c.setFont("Helvetica", 7.5)
            c.setFillColor(HexColor("#D0CDFF"))
            c.drawString(28, H - 83, clean_join([cv.get("email"), cv.get("phone"), cv.get("location")], " | "))

            c.setFillColor(PANEL_BG)
            c.rect(panel_x, 0, RIGHT_W, H - band_h, fill=1, stroke=0)

            draw_creative_page1_sidebar(c, cv, panel_x, RIGHT_W, H - band_h - 28, P1)
        else:
            c.setFillColor(PANEL_BG)
            c.rect(panel_x, 0, RIGHT_W, H, fill=1, stroke=0)

    draw_page_bg(page1=True)

    lx = 28
    y = H - band_h - 28

    section_heading(c, lx, y, "SUMMARY", P1, 45)
    y -= 18
    y = draw_wrapped(c, cv.get("summary"), lx, y, LEFT_W, size=8.5, leading=13, colour="#4A4A5A", alignment=TA_JUSTIFY) - 20

    section_heading(c, lx, y, "EXPERIENCE", P1, 60)
    y -= 18

    for job in cv.get("experience", []):
        if y < 90:
            c.showPage()
            draw_page_bg(page1=False)
            y = H - 40

        y = draw_role(c, job, lx, y, LEFT_W, P1, "#1A1A2E", "#4A4A5A", "#7A7A8A")

    c.save()
    buf.seek(0)
    return buf


def draw_impact_page1_sidebar(c, cv, right_x, right_w, top_y, teal):
    rx = right_x + 8
    panel_w = right_w - 22
    y = top_y

    section_heading(c, rx, y, "SKILLS", teal, 35)
    y -= 18

    y = draw_list_lines(
        c,
        cv.get("skills", []),
        rx,
        y,
        panel_w,
        "#374151",
        6.8,
        8.2,
        3,
        max_y=95,
    )

    y -= 8

    if y > 125:
        section_heading(c, rx, y, "EDUCATION", teal, 58)
        y -= 18
        draw_education(c, cv.get("education", []), rx, y, panel_w, "#111827", 6.6, 8, max_y=35)


def generate_impact_pdf(cv, colours):
    cv = normalise_cv(cv)

    HEADER_BG = safe_hex(colours.get("primary"), "#111827")
    TEAL = safe_hex(colours.get("accent"), "#0D9488")
    RIGHT_W = 180

    right_x = W - RIGHT_W - 4
    left_x = 28
    left_w = W - RIGHT_W - 58
    header_h = 100
    body_top = H - header_h - 22

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    def draw_page_bg(page1=False):
        footer_brand(c, cv.get("is_premium", True))

        if page1:
            c.setFillColor(HEADER_BG)
            c.rect(0, H - header_h, W, header_h, fill=1, stroke=0)

            c.setStrokeColor(TEAL)
            c.setLineWidth(3)
            c.line(0, H - header_h, W, H - header_h)

            c.setFont("Helvetica-Bold", 30)
            c.setFillColor(white)
            c.drawString(28, H - 48, cv["name"])

            c.setFont("Helvetica", 7.5)
            c.setFillColor(HexColor("#9CA3AF"))
            c.drawString(28, H - 85, clean_join([cv.get("email"), cv.get("phone"), cv.get("location")], " | "))

            photo_img = decode_photo(cv.get("photo"))
            if photo_img:
                c.setFillColor(white)
                c.circle(W - 65, H - header_h / 2, 34, fill=1, stroke=0)
                draw_circular_photo(c, photo_img, W - 65, H - header_h / 2, 31)

            c.setFillColor(HexColor("#F9FAFB"))
            c.rect(right_x - 10, 0, RIGHT_W + 14, body_top + 22, fill=1, stroke=0)

            draw_impact_page1_sidebar(c, cv, right_x, RIGHT_W, body_top, TEAL)
        else:
            c.setFillColor(HexColor("#F9FAFB"))
            c.rect(right_x - 10, 0, RIGHT_W + 14, H, fill=1, stroke=0)

    draw_page_bg(page1=True)

    y = body_top

    section_heading(c, left_x, y, "SUMMARY", TEAL, 45)
    y -= 18
    y = draw_wrapped(c, cv.get("summary"), left_x, y, left_w, size=8.5, leading=13, colour="#4B5563", alignment=TA_JUSTIFY) - 20

    section_heading(c, left_x, y, "EXPERIENCE", TEAL, 70)
    y -= 18

    for job in cv.get("experience", []):
        if y < 90:
            c.showPage()
            draw_page_bg(page1=False)
            y = H - 40

        y = draw_role(c, job, left_x + 18, y, left_w - 18, TEAL, "#111827", "#4B5563", "#9CA3AF")

    c.save()
    buf.seek(0)
    return buf


def generate_docx(cv):
    cv = normalise_cv(cv)

    doc = Document()

    title = doc.add_paragraph()
    run = title.add_run(cv["name"])
    run.font.size = Pt(24)
    run.bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    contact = doc.add_paragraph()
    crun = contact.add_run(clean_join([cv.get("email"), cv.get("phone"), cv.get("location")], " | "))
    crun.font.size = Pt(9)
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("SUMMARY", level=2)
    doc.add_paragraph(cv.get("summary", ""))

    doc.add_heading("EXPERIENCE", level=2)
    for job in cv.get("experience", []):
        jp = doc.add_paragraph()
        jr = jp.add_run(clean_join([job.get("title"), job.get("company")], " - "))
        jr.bold = True

        if job.get("dates"):
            dp = doc.add_paragraph(job.get("dates"))
            if dp.runs:
                dp.runs[0].font.size = Pt(9)

        for bullet in job.get("bullets", []):
            doc.add_paragraph(clean_text(bullet), style="List Bullet")

    doc.add_heading("SKILLS", level=2)
    doc.add_paragraph(cv.get("skills_csv", ""))

    doc.add_heading("EDUCATION", level=2)
    for edu in cv.get("education", []):
        if edu.get("line"):
            doc.add_paragraph(edu.get("line"))

    if cv.get("certifications"):
        doc.add_heading("CERTIFICATIONS", level=2)
        for cert in cv.get("certifications", []):
            doc.add_paragraph(cert)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "ok": True,
        "service": "rolealign-pdf-api",
        "version": API_VERSION,
        "endpoints": ["POST /generate-pdf", "POST /generate-docx", "GET /health"],
    })


@app.route("/health", methods=["GET"])
def health():
    return jsonify({
        "ok": True,
        "service": "rolealign-pdf-api",
        "version": API_VERSION,
    })


@app.route("/generate-pdf", methods=["POST"])
def gen_pdf():
    try:
        data = request.get_json() or {}
        cv = data.get("cv_data") or data.get("cv") or {}
        template = (data.get("template") or "executive").lower().strip()
        colours = data.get("colours") or {}

        render_version = data.get("render_version") or cv.get("render_version") or API_VERSION
        cv = normalise_cv({**cv, "render_version": render_version})

        print({
            "event": "generate_pdf",
            "template": template,
            "api_version": API_VERSION,
            "render_version": render_version,
            "name": cv.get("name"),
            "skills_count": len(cv.get("skills", [])),
            "experience_count": len(cv.get("experience", [])),
            "has_photo": bool(cv.get("photo")),
        })

        if template == "starter":
            cv["is_premium"] = False
            buf = generate_starter_pdf(cv, colours)
        elif template == "executive":
            buf = generate_executive_pdf(cv, colours)
        elif template == "creative":
            buf = generate_creative_pdf(cv, colours)
        elif template == "impact":
            buf = generate_impact_pdf(cv, colours)
        else:
            return jsonify({"error": f"Invalid template: {template}"}), 400

        name = re.sub(r"[^A-Za-z0-9_\-]+", "_", cv.get("name", "CV").replace(" ", "_"))

        return send_file(
            buf,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=f"CV_{template}_{name}.pdf",
        )
    except Exception as e:
        print(f"[generate_pdf_error] {e}")
        return jsonify({"error": str(e), "version": API_VERSION}), 500


@app.route("/generate-docx", methods=["POST"])
def gen_docx():
    try:
        data = request.get_json() or {}
        cv = normalise_cv(data.get("cv_data") or data.get("cv") or {})

        print({
            "event": "generate_docx",
            "api_version": API_VERSION,
            "name": cv.get("name"),
            "skills_count": len(cv.get("skills", [])),
            "experience_count": len(cv.get("experience", [])),
        })

        buf = generate_docx(cv)
        name = re.sub(r"[^A-Za-z0-9_\-]+", "_", cv.get("name", "CV").replace(" ", "_"))

        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=f"CV_{name}.docx",
        )
    except Exception as e:
        print(f"[generate_docx_error] {e}")
        return jsonify({"error": str(e), "version": API_VERSION}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port)
