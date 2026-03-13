# VTU Result Parser

A desktop tool to parse scanned VTU marksheet PDFs and export clean Excel sheets for result analysis.

This project includes:
- A drag-and-drop GUI app (`app.py`)
- OCR + PDF parsing engine (`pdfparser.py`)
- Excel export with 3 sheets:
  - `Result Sheet` (before revaluation layout)
  - `Credit Sheet` (TOT/GP/CP + SGPA)
  - `Raw Data` (source rows)

## Features

- Bulk parse all PDFs in a selected folder.
- OCR cleanup for common scan issues (code misreads, noisy tokens, split lines).
- Subject-wise marks extraction: Internal, External, Total, Result.
- One row per student in report sheets.
- Bottom summary tables (subject pass/fail stats + overall result).
- Credit-system sheet with formulas:
- Credits are asked once per run using a popup (mandatory before export).
- Works with bundled local `tesseract/` and `poppler/` folders for portable distribution.


## Requirements (Source Run)

- Windows
- Python 3.10+
- Tesseract OCR and Poppler binaries

Python packages:
- `pytesseract`
- `pdf2image`
- `opencv-python`
- `numpy`
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
5. Exported workbook is saved in the same selected folder.

## Bundled Binary Paths

`pdfparser.py` resolves binaries in this order:

1. Local bundled paths (preferred):
   - `./tesseract/tesseract.exe`
   - `./poppler/bin`
2. System fallback for Tesseract:
   - `C:\Program Files\Tesseract-OCR\tesseract.exe`

This also supports PyInstaller runtime extraction (`sys._MEIPASS`).


## Notes

- Keep Excel file closed while exporting, otherwise write can fail due to file lock.
- Credit popup is mandatory by design: no credits means no export.
