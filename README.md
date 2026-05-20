# PDF → DOCX Converter

A modern web-based PDF to DOCX conversion application built with Node.js, Express, and the `docx` library. The application allows users to upload PDF files through a drag-and-drop interface, extract embedded images, reconstruct formatted DOCX documents, and download the converted Word file.

The project focuses on preserving:

- Images
- Tables
- Document formatting
- Structured layouts
- Headings and sections

# Features

- Drag-and-drop PDF upload interface
- PDF to DOCX conversion pipeline
- Embedded image extraction using `pdfimages`
- Automatic image background cleanup
- DOCX document reconstruction
- Styled headings, paragraphs, tables, and bullet lists
- Real-time conversion progress UI
- Error handling and validation
- Automatic cleanup of temporary files
- Responsive frontend UI

# Tech Stack

## Backend

- Node.js
- Express.js
- Multer
- docx
- poppler-utils (`pdfimages`)
- Python + Pillow (for image cleanup)

## Frontend

- HTML5
- CSS3
- Vanilla JavaScript

---

## Project Structure

```
pdf2docx-app/
├── server.js
├── public/
│   └── index.html
├── package.json
└── README.md
```

# How the Application Works

## Conversion Pipeline

```text
User Uploads PDF
        ↓
PDF Validation
        ↓
Extract Embedded Images
        ↓
Clean Image Backgrounds
        ↓
Rebuild DOCX Structure
        ↓
Generate DOCX Buffer
        ↓
Download DOCX File
```

---

# Backend Workflow

## 1. File Upload

The frontend uploads a PDF file using a multipart form request.

Handled using:

- `multer`
- Express route `/convert`

Example:

```javascript
app.post("/convert", upload.single("pdf"), async (req, res) => {
```

---

## 2. Image Extraction

The application extracts embedded PDF images using:

```bash
pdfimages -png input.pdf output_prefix
```

This requires:

```bash
poppler-utils
```

Installed on Ubuntu/Debian:

```bash
sudo apt install poppler-utils
```

---

## 3. Image Cleaning

The project uses Python and Pillow to:

- remove black backgrounds
- apply masks
- clean extracted images

Example cleanup logic:

```python
from PIL import Image
```

---

## 4. DOCX Reconstruction

The project uses the `docx` library to dynamically generate:

- headings
- paragraphs
- tables
- bullet lists
- images
- styled sections

Main objects used:

```javascript
Document;
Paragraph;
TextRun;
ImageRun;
Table;
TableRow;
TableCell;
```

---

## 5. File Download

The generated DOCX buffer is returned directly to the browser.

Content type:

```text
application/vnd.openxmlformats-officedocument.wordprocessingml.document
```

---

# Frontend Overview

The frontend provides:

- Drag-and-drop upload
- File validation
- Progress animation
- Error handling
- DOCX download

Main frontend features:

## Upload Area

```html
<input type="file" id="fileInput" accept=".pdf" />
```

---

## Conversion Progress

The UI simulates conversion stages:

```javascript
const steps = [
  "Inspecting PDF…",
  "Extracting embedded images…",
  "Cleaning image backgrounds…",
  "Building document structure…",
  "Packing DOCX…",
];
```

---

# Installation

## 1. Clone Repository

```bash
git clone https://github.com/aaronGeb/pdf2docx-app.git
cd pdf2docx-app
```

---

## 2. Install Node.js Dependencies

```bash
npm install
```

---

## 3. Install Python Dependencies

Install Pillow:

```bash
pip install pillow numpy
```

---

## 4. Install poppler-utils

### Ubuntu/Debian

```bash
sudo apt install poppler-utils
```

### macOS

Using Homebrew:

```bash
brew install poppler
```

# Running the Application

## Development Mode

```bash
npm run dev
```

---

## Production Mode

```bash
npm start
```

---

# Access the Application

Open:

```text
http://localhost:3000
```

---

# API Reference

## POST /convert

Converts uploaded PDF into DOCX.

### Request

Multipart form data:

| Field | Type |
| ----- | ---- |
| pdf   | File |

---

### Response

Returns:

- `.docx` file buffer

---

### Example Using cURL

```bash
curl -X POST http://localhost:3000/convert \
  -F "pdf=@sample.pdf" \
  --output output.docx
```

---

# package.json Scripts

```json
"scripts": {
  "start": "node server.js",
  "dev": "node --watch server.js"
}
```

---

# Key Components

## server.js

Main backend application.

Responsibilities:

- file uploads
- image extraction
- image cleanup
- DOCX generation
- API handling
- temporary file cleanup

---

## index.html

Frontend interface.

Responsibilities:

- drag-and-drop UI
- upload handling
- conversion progress
- displaying errors
- download handling

---

# DOCX Generation Features

The generated DOCX supports:

- Styled headings
- Cover pages
- Warning sections
- Tables
- Bullet lists
- Embedded images
- Custom spacing and alignment

---

# Error Handling

The application validates:

- Missing files
- Invalid file types
- Failed image extraction
- Missing Poppler installation
- DOCX generation failures

Example:

```javascript
if (!req.file) return res.status(400).json({ error: "No PDF uploaded." });
```

---

# Temporary File Cleanup

After conversion:

- temporary directories are deleted
- uploaded PDFs are removed

Example:

```javascript
fs.rmSync(workDir, { recursive: true });
```

---

# Current Limitations

This project currently works best with:

- structured PDFs
- manuals
- PDFs with embedded images
- predictable layouts

Potential limitations:

- scanned PDFs
- OCR-heavy documents
- highly complex layouts
- unsupported fonts

---

# Future Improvements

Potential enhancements:

- OCR support using Tesseract
- Multi-page document parsing
- Better table reconstruction
- AI-assisted layout detection
- Authentication system
- Queue-based processing
- Docker deployment
- Cloud storage integration
- Drag-and-drop multiple files
- Real progress tracking

---

# Docker Deployment (Optional)

Example Dockerfile:

```dockerfile
FROM node:20

WORKDIR /app

COPY package*.json ./
RUN npm install

COPY . .

EXPOSE 3000

CMD ["npm", "start"]
```

---

# Recommended Production Improvements

For production deployment:

- Add rate limiting
- Add file size limits
- Use async queues
- Add logging and monitoring
- Store uploads in object storage
- Add security validation
- Containerize with Docker

---

# Example Use Cases

- Technical manual conversion
- Business document processing
- PDF archival workflows
- Automated document generation
- Enterprise document pipelines

---

# Dependencies

## Node.js Dependencies

```json
{
  "docx": "^9.6.1",
  "express": "^4.18.2",
  "multer": "^1.4.5-lts.1"
}
```

---

# Environment Requirements

- Node.js >= 18
- Python 3.11
- poppler-utils
- Pillow

---

# Security Considerations

Recommended improvements:

- Validate MIME types
- Add antivirus scanning
- Restrict upload size
- Use sandboxed processing
- Add authentication

---

# Performance Notes

Performance depends on:

- PDF size
- Number of embedded images
- DOCX complexity
- Server CPU and memory

---

# Troubleshooting

## Error: pdfimages failed

Install Poppler:

```bash
sudo apt install poppler-utils
```

---

## Error: No images extracted from PDF

Possible reasons:

- scanned PDF
- unsupported PDF structure
- encrypted PDF

---

## Error: Python image cleanup failed

Install dependencies:

```bash
pip install pillow numpy
```

---

# License

MIT License

---

# Acknowledgements

Libraries and tools used:

- Express.js
- docx
- Pillow
- Poppler
- Multer
- Node.js

---
