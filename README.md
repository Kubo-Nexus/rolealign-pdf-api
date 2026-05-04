# RoleAlign PDF API

Generates styled PDF and DOCX CVs with customizable colours.

## Endpoints

### POST /generate-pdf

Generates a PDF CV in the specified template style.

**Request:**

```json

{

  "template": "executive",

  "cv_data": {

    "name": "Kuben Naidoo",

    "email": "k.naidoo1206@gmail.com",

    "phone": "072 488 6362",

    "location": "Cape Town, South Africa",

    "linkedin": "linkedin.com/in/kuben-naidoo",

    "summary": "...",

    "experience": [...],

    "skills": [...],

    "education": [...]

  },

  "colours": {

    "primary": "#1B2A4A",

    "accent": "#C9A96E"

  }

}

Response: PDF file
POST /generate-docx
Generates an editable DOCX file.

Request: Same as /generate-pdf (colours ignored)

Response: DOCX file
Templates
executive - Navy sidebar, gold accents
creative - Purple gradient header, modern layout
impact - Dark header, timeline experience, tag badges
