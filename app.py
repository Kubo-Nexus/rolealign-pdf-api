import os
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor, white, black
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
CORS(app)

W, H = A4


@app.route('/', methods=['GET'])
def home():
    return jsonify({
        "status": "RoleAlign PDF API is running",
        "version": "1.0.0",
        "endpoints": {
            "POST /generate-pdf": "Generate a styled CV PDF",
            "POST /generate-docx": "Generate an editable CV DOCX"
        }
    })


def safe_hex(colour_str, fallback):
    try:
        if colour_str and colour_str.startswith('#') and len(colour_str) in (4, 7):
            return HexColor(colour_str)
    except Exception:
        pass
    return HexColor(fallback)


def generate_executive_pdf(cv, colours):
    SIDEBAR_W = 190
    NAVY = safe_hex(colours.get('primary'), '#1B2A4A')
    ACCENT = safe_hex(colours.get('accent'), '#C9A96E')
    TEXT_DARK = HexColor('#1A1A1A')
    TEXT_MED = HexColor('#4A4A4A')
    TEXT_LIGHT = HexColor('#7A7A7A')
    NAVY_LIGHT = HexColor('#2A3F6A')

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    def draw_sidebar(c):
        c.setFillColor(NAVY)
        c.rect(0, 0, SIDEBAR_W, H, fill=1, stroke=0)

    draw_sidebar(c)

    cx, cy = SIDEBAR_W / 2, H - 70
    c.setFillColor(NAVY_LIGHT)
    c.circle(cx, cy, 36, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont('Helvetica-Bold', 18)
    initials = ''.join([w[0] for w in cv.get('name', 'CV').split()[:2]]).upper()
    c.drawCentredString(cx, cy - 6, initials)

    sx = 18
    y = H - 130
    c.setFont('Helvetica-Bold', 9)
    c.setFillColor(ACCENT)
    c.drawString(sx, y, 'CONTACT')
    c.setStrokeColor(ACCENT)
    c.setLineWidth(0.5)
    c.line(sx, y - 4, SIDEBAR_W - 18, y - 4)
    y -= 20

    c.setFont('Helvetica', 7.5)
    c.setFillColor(HexColor('#D0D0D0'))
    for info in [cv.get('email', ''), cv.get('phone', ''), cv.get('location', ''), cv.get('linkedin', '')]:
        if info:
            c.drawString(sx, y, str(info))
            y -= 13

    y -= 12
    c.setFont('Helvetica-Bold', 9)
    c.setFillColor(ACCENT)
    c.drawString(sx, y, 'SKILLS')
    c.setStrokeColor(ACCENT)
    c.line(sx, y - 4, SIDEBAR_W - 18, y - 4)
    y -= 20

    c.setFont('Helvetica', 7.5)
    c.setFillColor(HexColor('#D0D0D0'))
    for skill in cv.get('skills', [])[:12]:
        c.drawString(sx, y, str(skill))
        y -= 8
        bar_w = SIDEBAR_W - 36
        c.setFillColor(HexColor('#2A3F6A'))
        c.roundRect(sx, y - 1, bar_w, 4, 2, fill=1, stroke=0)
        c.setFillColor(ACCENT)
        c.roundRect(sx, y - 1, bar_w * 0.8, 4, 2, fill=1, stroke=0)
        y -= 14
        c.setFillColor(HexColor('#D0D0D0'))

    y -= 10
    c.setFont('Helvetica-Bold', 9)
    c.setFillColor(ACCENT)
    c.drawString(sx, y, 'EDUCATION')
    c.setStrokeColor(ACCENT)
    c.line(sx, y - 4, SIDEBAR_W - 18, y - 4)
    y -= 20

    for edu in cv.get('education', []):
        c.setFont('Helvetica-Bold', 7.5)
        c.setFillColor(white)
        c.drawString(sx, y, str(edu.get('degree', '')))
        y -= 12
        c.setFont('Helvetica', 7)
        c.setFillColor(HexColor('#A0A0A0'))
        inst = edu.get('institution', '')
        yr = edu.get('year', '')
        c.drawString(sx, y, '{} - {}'.format(inst, yr))
        y -= 18

    mx = SIDEBAR_W + 24
    mw = W - mx - 24
    y = H - 45

    c.setFont('Helvetica-Bold', 24)
    c.setFillColor(NAVY)
    c.drawString(mx, y, str(cv.get('name', '')))
    y -= 24

    c.setFont('Helvetica-Bold', 10)
    c.setFillColor(NAVY)
    c.drawString(mx, y, 'SUMMARY')
    c.setStrokeColor(NAVY)
    c.setLineWidth(0.5)
    c.line(mx, y - 4, W - 24, y - 4)
    y -= 16

    style = ParagraphStyle('summary', fontName='Helvetica', fontSize=8.5,
                           leading=13, textColor=TEXT_MED, alignment=TA_JUSTIFY)
    summary = cv.get('summary', '')
    if summary:
        p = Paragraph(str(summary), style)
        pw, ph = p.wrap(mw, 200)
        p.drawOn(c, mx, y - ph)
        y -= ph + 16

    c.setFont('Helvetica-Bold', 10)
    c.setFillColor(NAVY)
    c.drawString(mx, y, 'EXPERIENCE')
    c.setStrokeColor(NAVY)
    c.line(mx, y - 4, W - 24, y - 4)
    y -= 18

    bstyle = ParagraphStyle('bullet', fontName='Helvetica', fontSize=8,
                            leading=11.5, textColor=TEXT_MED, alignment=TA_LEFT,
                            leftIndent=10, bulletIndent=0)

    for job in cv.get('experience', []):
        if y < 100:
            c.showPage()
            draw_sidebar(c)
            y = H - 40

        c.setFont('Helvetica-Bold', 9)
        c.setFillColor(TEXT_DARK)
        c.drawString(mx, y, str(job.get('title', '')))
        c.setFont('Helvetica', 7.5)
        c.setFillColor(TEXT_LIGHT)
        dates = str(job.get('dates', ''))
        tw = c.stringWidth(dates, 'Helvetica', 7.5)
        c.drawString(W - 24 - tw, y, dates)
        y -= 12

        c.setFont('Helvetica', 8)
        c.setFillColor(ACCENT)
        c.drawString(mx, y, str(job.get('company', '')))
        y -= 14

        for bullet in job.get('bullets', []):
            txt = '<bullet>&#8226;</bullet> {}'.format(str(bullet))
            p = Paragraph(txt, bstyle)
            pw, ph = p.wrap(mw - 10, 200)
            if y - ph < 40:
                c.showPage()
                draw_sidebar(c)
                y = H - 40
            p.drawOn(c, mx, y - ph)
            y -= ph + 3
        y -= 8

    c.setFont('Helvetica', 6)
    c.setFillColor(TEXT_LIGHT)
    c.drawCentredString(W / 2 + SIDEBAR_W / 2, 15, 'Created with RoleAlign')

    c.save()
    buf.seek(0)
    return buf


def generate_creative_pdf(cv, colours):
    PURPLE_1 = safe_hex(colours.get('primary_1', colours.get('primary')), '#6366F1')
    PURPLE_2 = safe_hex(colours.get('primary_2', colours.get('accent')), '#8B5CF6')
    TEXT_DARK = HexColor('#1A1A2E')
    TEXT_MED = HexColor('#4A4A5A')
    TEXT_LIGHT = HexColor('#7A7A8A')
    RIGHT_W = 185
    LEFT_W = W - RIGHT_W - 48
    PANEL_BG = HexColor('#F0EDFF')

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    band_h = 90
    steps = 50
    for i in range(steps):
        r1, g1, b1 = 0.388, 0.400, 0.945
        r2, g2, b2 = 0.545, 0.361, 0.965
        t = i / float(steps)
        r = r1 + (r2 - r1) * t
        g = g1 + (g2 - g1) * t
        b = b1 + (b2 - b1) * t
        sh = band_h / float(steps)
        c.setFillColor(HexColor('#{:02x}{:02x}{:02x}'.format(int(r*255), int(g*255), int(b*255))))
        c.rect(0, H - band_h + i * sh, W, sh + 1, fill=1, stroke=0)

    c.setFont('Helvetica-Bold', 28)
    c.setFillColor(white)
    c.drawString(28, H - 50, str(cv.get('name', '')))
    c.setFont('Helvetica', 7.5)
    c.setFillColor(HexColor('#D0CDFF'))
    contact = '{}  |  {}  |  {}'.format(cv.get('email', ''), cv.get('phone', ''), cv.get('location', ''))
    c.drawString(28, H - 83, contact)

    panel_x = W - RIGHT_W
    c.setFillColor(PANEL_BG)
    c.rect(panel_x, 0, RIGHT_W, H - band_h, fill=1, stroke=0)

    lx = 28
    y = H - band_h - 28

    c.setFont('Helvetica-Bold', 10)
    c.setFillColor(PURPLE_1)
    c.drawString(lx, y, 'SUMMARY')
    c.setStrokeColor(PURPLE_1)
    c.setLineWidth(1.5)
    c.line(lx, y - 4, lx + 45, y - 4)
    y -= 18

    style = ParagraphStyle('summary', fontName='Helvetica', fontSize=8.5,
                           leading=13, textColor=TEXT_MED, alignment=TA_JUSTIFY)
    summary = cv.get('summary', '')
    if summary:
        p = Paragraph(str(summary), style)
        pw, ph = p.wrap(LEFT_W, 200)
        p.drawOn(c, lx, y - ph)
        y -= ph + 20

    c.setFont('Helvetica-Bold', 10)
    c.setFillColor(PURPLE_1)
    c.drawString(lx, y, 'EXPERIENCE')
    c.setStrokeColor(PURPLE_1)
    c.setLineWidth(1.5)
    c.line(lx, y - 4, lx + 60, y - 4)
    y -= 18

    bstyle = ParagraphStyle('bullet', fontName='Helvetica', fontSize=8,
                            leading=11, textColor=TEXT_MED, leftIndent=12, bulletIndent=0)

    for job in cv.get('experience', []):
        if y < 80:
            c.showPage()
            y = H - 40
            c.setFillColor(PANEL_BG)
            c.rect(panel_x, 0, RIGHT_W, H, fill=1, stroke=0)

        c.setFont('Helvetica-Bold', 9)
        c.setFillColor(TEXT_DARK)
        c.drawString(lx, y, str(job.get('title', '')))
        c.setFont('Helvetica', 7.5)
        c.setFillColor(TEXT_LIGHT)
        dates = str(job.get('dates', ''))
        tw = c.stringWidth(dates, 'Helvetica', 7.5)
        c.drawString(lx + LEFT_W - tw, y, dates)
        y -= 12

        c.setFont('Helvetica', 8)
        c.setFillColor(PURPLE_1)
        c.drawString(lx, y, str(job.get('company', '')))
        y -= 14

        for bullet in job.get('bullets', []):
            c.setFillColor(PURPLE_2)
            c.circle(lx + 4, y + 2.5, 2, fill=1, stroke=0)
            p = Paragraph(str(bullet), bstyle)
            pw, ph = p.wrap(LEFT_W - 14, 200)
            if y - ph < 40:
                c.showPage()
                y = H - 40
                c.setFillColor(PANEL_BG)
                c.rect(panel_x, 0, RIGHT_W, H, fill=1, stroke=0)
            p.drawOn(c, lx, y - ph + 3)
            y -= ph + 3
        y -= 10

    rx = panel_x + 14
    ry = H - band_h - 28

    photo_cx = panel_x + RIGHT_W / 2
    photo_cy = ry + 5
    c.setFillColor(HexColor('#DDD8FF'))
    c.circle(photo_cx, photo_cy, 32, fill=1, stroke=0)
    c.setFillColor(PURPLE_1)
    c.setFont('Helvetica-Bold', 16)
    initials = ''.join([w[0] for w in cv.get('name', 'CV').split()[:2]]).upper()
    c.drawCentredString(photo_cx, photo_cy - 5, initials)
    ry -= 50

    c.setFont('Helvetica-Bold', 9)
    c.setFillColor(PURPLE_1)
    c.drawString(rx, ry, 'SKILLS')
    c.setStrokeColor(PURPLE_1)
    c.setLineWidth(1.5)
    c.line(rx, ry - 4, rx + 35, ry - 4)
    ry -= 20

    pill_x = rx
    for skill in cv.get('skills', []):
        c.setFont('Helvetica', 7)
        tw = c.stringWidth(str(skill), 'Helvetica', 7) + 14
        if pill_x + tw > panel_x + RIGHT_W - 14:
            pill_x = rx
            ry -= 22
        c.setFillColor(HexColor('#EDE9FE'))
        c.roundRect(pill_x, ry - 4, tw, 16, 8, fill=1, stroke=0)
        c.setFillColor(PURPLE_1)
        c.drawString(pill_x + 7, ry + 2, str(skill))
        pill_x += tw + 6

    ry -= 36

    c.setFont('Helvetica-Bold', 9)
    c.setFillColor(PURPLE_1)
    c.drawString(rx, ry, 'EDUCATION')
    c.setStrokeColor(PURPLE_1)
    c.setLineWidth(1.5)
    c.line(rx, ry - 4, rx + 58, ry - 4)
    ry -= 20

    for edu in cv.get('education', []):
        c.setFont('Helvetica-Bold', 8)
        c.setFillColor(TEXT_DARK)
        c.drawString(rx, ry, str(edu.get('degree', '')))
        ry -= 12
        c.setFont('Helvetica', 7)
        c.setFillColor(TEXT_LIGHT)
        c.drawString(rx, ry, '{} - {}'.format(edu.get('institution', ''), edu.get('year', '')))
        ry -= 18

    c.setFont('Helvetica', 6)
    c.setFillColor(TEXT_LIGHT)
    c.drawCentredString(W / 2, 12, 'Created with RoleAlign')

    c.save()
    buf.seek(0)
    return buf


def generate_impact_pdf(cv, colours):
    HEADER_BG = safe_hex(colours.get('primary'), '#111827')
    TEAL = safe_hex(colours.get('accent'), '#0D9488')
    TEXT_DARK = HexColor('#111827')
    TEXT_MED = HexColor('#4B5563')
    TEXT_LIGHT = HexColor('#9CA3AF')
    PANEL_BG = HexColor('#F9FAFB')
    RIGHT_W = 180

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    header_h = 100
    c.setFillColor(HEADER_BG)
    c.rect(0, H - header_h, W, header_h, fill=1, stroke=0)
    c.setStrokeColor(TEAL)
    c.setLineWidth(3)
    c.line(0, H - header_h, W, H - header_h)

    c.setFont('Helvetica-Bold', 30)
    c.setFillColor(white)
    c.drawString(28, H - 48, str(cv.get('name', '')))
    c.setFont('Helvetica', 7.5)
    c.setFillColor(HexColor('#9CA3AF'))
    contact = '{}  |  {}  |  {}'.format(cv.get('email', ''), cv.get('phone', ''), cv.get('location', ''))
    c.drawString(28, H - 85, contact)

    photo_cx = W - 65
    photo_cy = H - header_h / 2
    c.setFillColor(white)
    c.circle(photo_cx, photo_cy, 34, fill=1, stroke=0)
    c.setFillColor(HexColor('#374151'))
    c.circle(photo_cx, photo_cy, 31, fill=1, stroke=0)
    c.setFont('Helvetica-Bold', 16)
    c.setFillColor(white)
    initials = ''.join([w[0] for w in cv.get('name', 'CV').split()[:2]]).upper()
    c.drawCentredString(photo_cx, photo_cy - 5, initials)

    right_x = W - RIGHT_W - 4
    left_x = 28
    left_w = W - RIGHT_W - 52
    body_top = H - header_h - 22

    c.setFillColor(PANEL_BG)
    c.rect(right_x - 10, 0, RIGHT_W + 14, body_top + 22, fill=1, stroke=0)

    y = body_top

    c.setFont('Helvetica-Bold', 10)
    c.setFillColor(TEXT_DARK)
    c.drawString(left_x, y, 'SUMMARY')
    c.setFillColor(TEAL)
    c.rect(left_x, y - 6, 30, 2.5, fill=1, stroke=0)
    y -= 18

    style = ParagraphStyle('summary', fontName='Helvetica', fontSize=8.5,
                           leading=13, textColor=TEXT_MED, alignment=TA_JUSTIFY)
    summary = cv.get('summary', '')
    if summary:
        p = Paragraph(str(summary), style)
        pw, ph = p.wrap(left_w, 200)
        p.drawOn(c, left_x, y - ph)
        y -= ph + 20

    c.setFont('Helvetica-Bold', 10)
    c.setFillColor(TEXT_DARK)
    c.drawString(left_x, y, 'EXPERIENCE')
    c.setFillColor(TEAL)
    c.rect(left_x, y - 6, 50, 2.5, fill=1, stroke=0)
    y -= 18

    timeline_x = left_x + 6
    bstyle = ParagraphStyle('bullet', fontName='Helvetica', fontSize=8,
                            leading=11, textColor=TEXT_MED, leftIndent=12, bulletIndent=0)

    exp_list = cv.get('experience', [])
    for i, job in enumerate(exp_list):
        if y < 80:
            c.showPage()
            y = H - 40
            c.setFillColor(PANEL_BG)
            c.rect(right_x - 10, 0, RIGHT_W + 14, H, fill=1, stroke=0)

        c.setFillColor(TEAL)
        c.circle(timeline_x, y + 2, 4, fill=1, stroke=0)
        c.setFillColor(white)
        c.circle(timeline_x, y + 2, 2, fill=1, stroke=0)

        c.setFont('Helvetica-Bold', 9)
        c.setFillColor(TEXT_DARK)
        c.drawString(left_x + 18, y, str(job.get('title', '')))

        c.setFont('Helvetica', 7.5)
        c.setFillColor(TEXT_LIGHT)
        dates = str(job.get('dates', ''))
        tw = c.stringWidth(dates, 'Helvetica', 7.5)
        c.drawString(left_x + left_w - tw, y, dates)
        y -= 12

        c.setFont('Helvetica', 8)
        c.setFillColor(TEAL)
        c.drawString(left_x + 18, y, str(job.get('company', '')))
        y -= 14

        for bullet in job.get('bullets', []):
            c.setFillColor(HexColor('#D1D5DB'))
            c.circle(left_x + 21, y + 2.5, 1.5, fill=1, stroke=0)
            p = Paragraph(str(bullet), bstyle)
            pw, ph = p.wrap(left_w - 24, 200)
            if y - ph < 40:
                c.showPage()
                y = H - 40
                c.setFillColor(PANEL_BG)
                c.rect(right_x - 10, 0, RIGHT_W + 14, H, fill=1, stroke=0)
            p.drawOn(c, left_x + 18, y - ph + 3)
            y -= ph + 3

        if i < len(exp_list) - 1:
            c.setStrokeColor(HexColor('#E5E7EB'))
            c.setLineWidth(1)
            c.line(timeline_x, y + 3, timeline_x, y - 6)
        y -= 10

    ry = body_top
    rx = right_x + 6

    c.setFont('Helvetica-Bold', 9)
    c.setFillColor(TEXT_DARK)
    c.drawString(rx, ry, 'SKILLS')
    c.setFillColor(TEAL)
    c.rect(rx, ry - 6, 30, 2.5, fill=1, stroke=0)
    ry -= 22

    tag_x = rx
    for skill in cv.get('skills', []):
        c.setFont('Helvetica', 7)
        tw = c.stringWidth(str(skill), 'Helvetica', 7) + 12
        th = 17
        if tag_x + tw > right_x + RIGHT_W - 4:
            tag_x = rx
            ry -= 22
        c.setFillColor(HEADER_BG)
        c.roundRect(tag_x, ry - 4, tw, th, 4, fill=1, stroke=0)
        c.setFillColor(white)
        c.drawString(tag_x + 6, ry + 2, str(skill))
        tag_x += tw + 5

    ry -= 38

    c.setFont('Helvetica-Bold', 9)
    c.setFillColor(TEXT_DARK)
    c.drawString(rx, ry, 'EDUCATION')
    c.setFillColor(TEAL)
    c.rect(rx, ry - 6, 50, 2.5, fill=1, stroke=0)
    ry -= 22

    for edu in cv.get('education', []):
        c.setFont('Helvetica-Bold', 8)
        c.setFillColor(TEXT_DARK)
        c.drawString(rx, ry, str(edu.get('degree', '')))
        ry -= 12
        c.setFont('Helvetica', 7)
        c.setFillColor(TEXT_LIGHT)
        c.drawString(rx, ry, '{} - {}'.format(edu.get('institution', ''), edu.get('year', '')))
        ry -= 18

    c.setFont('Helvetica', 6)
    c.setFillColor(TEXT_LIGHT)
    c.drawCentredString(W / 2, 12, 'Created with RoleAlign')

    c.save()
    buf.seek(0)
    return buf


def generate_docx(cv):
    doc = Document()

    title = doc.add_paragraph()
    run = title.add_run(str(cv.get('name', '')))
    run.font.size = Pt(24)
    run.bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    contact = doc.add_paragraph()
    ct = ' | '.join(filter(None, [cv.get('email', ''), cv.get('phone', ''), cv.get('location', '')]))
    crun = contact.add_run(ct)
    crun.font.size = Pt(9)
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    doc.add_heading('SUMMARY', level=2)
    doc.add_paragraph(str(cv.get('summary', '')))

    doc.add_heading('EXPERIENCE', level=2)
    for job in cv.get('experience', []):
        jp = doc.add_paragraph()
        jr = jp.add_run('{} - {}'.format(job.get('title', ''), job.get('company', '')))
        jr.bold = True
        dp = doc.add_paragraph(str(job.get('dates', '')))
        dp.runs[0].font.size = Pt(9)
        for bullet in job.get('bullets', []):
            doc.add_paragraph(str(bullet), style='List Bullet')

    doc.add_heading('SKILLS', level=2)
    skills = ', '.join([str(s) for s in cv.get('skills', [])])
    doc.add_paragraph(skills)

    doc.add_heading('EDUCATION', level=2)
    for edu in cv.get('education', []):
        ep = doc.add_paragraph()
        er = ep.add_run(str(edu.get('degree', '')))
        er.bold = True
        ep.add_run(' - {} ({})'.format(edu.get('institution', ''), edu.get('year', '')))

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


@app.route('/generate-pdf', methods=['POST'])
def gen_pdf():
    try:
        data = request.get_json()
        cv = data.get('cv_data', {})
        template = data.get('template', 'executive')
        colours = data.get('colours') or {}

        if template == 'executive':
            buf = generate_executive_pdf(cv, colours)
        elif template == 'creative':
            buf = generate_creative_pdf(cv, colours)
        elif template == 'impact':
            buf = generate_impact_pdf(cv, colours)
        else:
            return jsonify({"error": "Invalid template"}), 400

        name = str(cv.get('name', 'CV')).replace(' ', '_')
        return send_file(
            buf,
            mimetype='application/pdf',
            as_attachment=True,
            download_name='CV_{}_{}.pdf'.format(template, name)
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/generate-docx', methods=['POST'])
def gen_docx():
    try:
        data = request.get_json()
        cv = data.get('cv_data', {})

        buf = generate_docx(cv)

        name = str(cv.get('name', 'CV')).replace(' ', '_')
        return send_file(
            buf,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='CV_{}.docx'.format(name)
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
