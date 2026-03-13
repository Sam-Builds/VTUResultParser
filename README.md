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
  - `GP` from total marks scale
  - `CP = GP * credit`
  - `Total CP`, `SGPA`, `Remarks`, `No. of Backlogs`, `Percentage`
- Credits are asked once per run using a popup (mandatory before export).
- Works with bundled local `tesseract/` and `poppler/` folders for portable distribution.

## Project Structure

```text
app.py
pdfparser.py
tesseract/
poppler/
VTU_Parser.spec
dist/
release/
```

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

Install dependencies:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install pytesseract pdf2image opencv-python numpy pandas openpyxl tkinterdnd2 pyinstaller
```

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

## Output Workbook

### 1) Result Sheet
- Institution and semester heading rows.
- Subjects laid out horizontally with `INT`, `EXT`, `TOT`.
- Row-wise `Total Marks`, `%`, and `Total no. of Fail`.
- Bottom summary:
  - No. of students taking/pass/fail per subject
  - Result in percentage per subject
  - Overall pass/fail count and overall percentage

### 2) Credit Sheet
- Subjects laid out as `TOT`, `GP`, `CP`.
- GP formula:

```excel
=IF(D6>=90,10,IF(D6>=80,9,IF(D6>=70,8,IF(D6>=60,7,IF(D6>=55,6,IF(D6>=50,5,IF(D6>=40,4,0)))))))
```

- CP formula example (if credit for subject is 4):

```excel
=E6*4
```

- Row totals:
  - `TOTAL CP`
  - `SGPA = TOTAL CP / (sum of all entered subject credits)`
  - `REMARKS` (`PASS` if no backlogs, otherwise `FAIL`)
  - `No. of Backlogs`
  - `Percentage`
- Includes the same bottom summary format.

### 3) Raw Data
- One row per parsed subject entry.
- Useful for verification/debugging.

## Bundled Binary Paths

`pdfparser.py` resolves binaries in this order:

1. Local bundled paths (preferred):
   - `./tesseract/tesseract.exe`
   - `./poppler/bin`
2. System fallback for Tesseract:
   - `C:\Program Files\Tesseract-OCR\tesseract.exe`

This also supports PyInstaller runtime extraction (`sys._MEIPASS`).

## Build EXE

```powershell
.\.venv\Scripts\Activate.ps1
python -m PyInstaller --noconfirm --clean --windowed --onefile --name VTU_Parser app.py --add-data "tesseract;tesseract" --add-data "poppler;poppler"
```

Output:
- `dist\VTU_Parser.exe`

## Create Release ZIP

```powershell
New-Item -ItemType Directory -Force -Path ".\release\VTU_Parser_Package_Lite" | Out-Null
Copy-Item ".\dist\VTU_Parser.exe" ".\release\VTU_Parser_Package_Lite\" -Force
Copy-Item ".\app.py" ".\release\VTU_Parser_Package_Lite\" -Force
Copy-Item ".\pdfparser.py" ".\release\VTU_Parser_Package_Lite\" -Force
Copy-Item ".\tesseract" ".\release\VTU_Parser_Package_Lite\tesseract" -Recurse -Force
Copy-Item ".\poppler" ".\release\VTU_Parser_Package_Lite\poppler" -Recurse -Force
Compress-Archive -Path ".\release\VTU_Parser_Package_Lite\*" -DestinationPath ".\release\VTU_Parser_Package_Lite.zip" -Force
```

## Notes

- Keep Excel file closed while exporting, otherwise write can fail due to file lock.
- OCR quality depends on scan clarity. Better scans improve extraction reliability.
- Credit popup is mandatory by design: no credits means no export.

## License

Add your preferred license (for example MIT) before publishing publicly.
