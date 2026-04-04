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
Preferred local startup with the bundled virtual environment:

1. Activate the virtual environment:
   `.venv\Scripts\activate`
2. Install or refresh dependencies:
   `python -m pip install -r requirements.txt`
3. Start the app:
   `python run_server.py`
4. Open:
   `http://127.0.0.1:8000`

Alternative direct command:

`python -m uvicorn app:app --host 127.0.0.1 --port 8000`

Windows helper launchers already included:
- `run-server.cmd`
- `run-server.ps1`
- `start-pdf-editor.bat`

## Notes
- The app listens on port `8000`.
- Best results for digitally-generated PDFs (not scanned images).
- For scanned PDFs, OCR integration is needed (next step).
