import os
from flask import Flask, request, send_file
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor, white
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY

app = Flask(__name__)

@app.route('/', methods=['GET'])
def home():
    return {
        "status": "RoleAlign PDF API is running",
        "message": "API is ready for deployment"
    }

@app.route('/generate-pdf', methods=['POST'])
def generate_pdf():
    try:
        data = request.get_json()
        return {
            "status": "success",
            "template": data.get('template', 'executive'),
            "message": "PDF generation working"
        }
    except Exception as e:
        return {"error": str(e)}, 500

@app.route('/generate-docx', methods=['POST'])
def generate_docx_route():
    try:
        data = request.get_json()
        return {
            "status": "success",
            "message": "DOCX generation working"
        }
    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
