# VTU Result Parser

~~A desktop tool to parse scanned VTU marksheet PDFs and export clean Excel sheets for result analysis.~~

**[NEW]** A desktop tool to parse VTU marksheet PDFs using selectable text extraction (no OCR pipeline) and export clean Excel sheets for result analysis.

This project includes:
- A drag-and-drop GUI app (`app.py`)
- ~~OCR + PDF parsing engine (`pdfparser.py`)~~
- **[NEW]** Text-first PDF parsing engine (`pdfparser.py`) using `pikepdf` + `pdfplumber`
- Excel export with 3 sheets:
  - `Result Sheet` (before revaluation layout)
  - `Credit Sheet` (TOT/GP/CP + SGPA)
  - `Raw Data` (source rows)

## Features

- Bulk parse all PDFs in a selected folder.
- ~~OCR cleanup for common scan issues (code misreads, noisy tokens, split lines).~~
- **[NEW]** Fast text extraction from unprotected PDFs via `pikepdf` + `pdfplumber`.
- **[NEW]** Optional parallel parsing (up to 3 worker threads) for faster bulk runs.
- Subject-wise marks extraction: Internal, External, Total, Result.
- **[NEW]** Fail counting logic is result-driven: any non-`P` result is treated as fail.
- **[NEW]** Student names are normalized to uppercase in exported sheets.
- One row per student in report sheets.
- Bottom summary tables (subject pass/fail stats + overall result).
- Credit-system sheet with formulas:
- Credits are asked once per run using a popup (mandatory before export).
- ~~Works with bundled local `tesseract/` and `poppler/` folders for portable distribution.~~
- **[NEW]** Lightweight packaging: no bundled OCR binaries required.


## Requirements (Source Run)

- Windows
- Python 3.10+
- ~~Tesseract OCR and Poppler binaries~~
- **[NEW]** No external OCR binaries required

Python packages:
- ~~`pytesseract`~~
- ~~`pdf2image`~~
- ~~`opencv-python`~~
- ~~`numpy`~~
- **[NEW]** `pikepdf`
- **[NEW]** `pdfplumber`
- `pandas`
- `openpyxl`
- `tkinterdnd2`

## Run From Source

```powershell
.\.venv\Scripts\Activate.ps1
python .\app.py
```

## How To Use

1. Open the app.
2. Drag and drop a folder containing VTU marksheet PDFs (or click `Browse Folder`).
3. Click `Parse & Export to Excel`.
4. Enter subject credits in the popup (required).
5. **[NEW]** Enter subject sort positions in the same popup (required and editable).
6. **[NEW]** Set sheet config in app UI before export (College, Department, Year Period, Revaluation status, Semester).
7. Exported workbook is saved in the same selected folder.

## Bundled Binary Paths

~~`pdfparser.py` resolves binaries in this order:~~

~~1. Local bundled paths (preferred):~~
~~   - `./tesseract/tesseract.exe`~~
~~   - `./poppler/bin`~~
~~2. System fallback for Tesseract:~~
~~   - `C:\Program Files\Tesseract-OCR\tesseract.exe`~~

~~This also supports PyInstaller runtime extraction (`sys._MEIPASS`).~~

**[NEW]** Current parser path flow:
- Opens PDF with `pikepdf` (strips protection when possible)
- Extracts selectable text with `pdfplumber`
- Parses subject rows directly from text (no image conversion/OCR)


## Notes

- Keep Excel file closed while exporting, otherwise write can fail due to file lock.
- Credit popup is mandatory by design: no credits means no export.
- **[NEW]** Column sizing policy in exports: `A:C` are fixed; `D` onward use fixed width (`10`) to avoid accidental hidden numeric content.
