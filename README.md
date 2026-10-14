# Incuxai PDF-to-Excel AI Agent

A student-built prototype that extracts tables and text from PDFs and exports an Excel workbook with an extraction log.

## Features
- PDF upload, up to 50 pages and 25 MB
- Table extraction with PDFPlumber
- Text extraction when a page contains no detected table
- OCR fallback for scanned pages using PyMuPDF + Tesseract
- Excel workbook with separate data blocks and an extraction log
- Basic output validation flags and preview
- CSV download of the extraction log

## Important limitations
- OCR/table extraction can make mistakes. Review every important value against the source PDF.
- Scanned table column detection is heuristic; complex layouts may need manual correction.
- This prototype does not use a hosted LLM. It combines document extraction libraries and OCR. Do not describe it as a trained AI model.
- Uploaded files are processed in the running app session; do not upload confidential documents.

## Run locally
1. Install Python 3.10+.
2. Install Tesseract OCR on your operating system.
3. Open a terminal in this folder and run:
   `pip install -r requirements.txt`
4. Run:
   `streamlit run app.py`

## Deploy using GitHub + Streamlit Community Cloud (phone-friendly)
1. Create a GitHub repository named `incuxai-pdf-to-excel`.
2. Add `app.py`, `requirements.txt`, and `packages.txt` to the repository root.
3. Open https://share.streamlit.io and sign in using GitHub.
4. Choose **Create app**, select the repository, branch `main`, and file `app.py`.
5. Deploy and wait for the build to finish.
6. Open the live URL and test with the included `sample_files/sample_tables.pdf`.

Streamlit's deployment guide: https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy

## Project structure
- `app.py` — application
- `requirements.txt` — Python dependencies
- `packages.txt` — system package for OCR
- `sample_files/sample_tables.pdf` — sample test PDF
- `sample_files/expected_sample.xlsx` — expected values for the sample
- `DEMO_SCRIPT.md` — short demonstration script
- `VALIDATION_REPORT_TEMPLATE.md` — record measured results
- `SUBMISSION_CHECKLIST.md` — final submission checklist
