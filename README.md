# PDF/RPT Bank Statement to Excel Converter

A lightweight Flask app that converts bank statement PDFs or RPT/TXT exports into a structured Excel file. The parser uses heuristics to detect dates, references, withdrawals, deposits, balances, and balance types (Dr/Cr), while merging multi-line narrations.

## Features
- PDF and RPT/TXT upload
- Heuristic, multi-bank parsing
- Clean Excel output with key columns
- Optional frontend-only deployment for Vercel

## Getting started (backend)

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open `http://localhost:5000` in your browser.

## Vercel frontend deployment

1. Deploy the `frontend/` directory as a static site in Vercel.
2. In the deployed UI, enter your backend API URL (e.g. `https://api.yourdomain.com`) before uploading.
   - If the frontend and backend share the same domain, leave the field blank.

## Notes
- PDF parsing uses `pdfplumber`; scanned images may require OCR before upload.
- Excel output relies on `openpyxl`.
