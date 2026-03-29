# PDF Editor Tool

Simple local tool for editing text blocks inside PDF files.

## Features
- Upload PDF
- Detect editable text blocks
- Edit text in browser
- Export edited PDF

## Tech
- Backend: FastAPI + PyMuPDF
- Frontend: HTML + CSS + JavaScript

## Run
1. Create virtual environment:
   python -m venv .venv
2. Activate:
   .venv\Scripts\activate
3. Install deps:
   pip install -r requirements.txt
4. Start app:
   uvicorn app:app --reload --port 8080
5. Open:
   http://127.0.0.1:8080

## Notes
- Best results for digitally-generated PDFs (not scanned images).
- For scanned PDFs, OCR integration is needed (next step).
