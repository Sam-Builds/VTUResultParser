import re
import sys
from datetime import datetime
from collections import OrderedDict, defaultdict
from pathlib import Path

import cv2
import numpy as np
import pandas as pd
import pytesseract
from pdf2image import convert_from_path

_BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
_LOCAL_TESS_EXE = _BASE_DIR / "tesseract" / "tesseract.exe"
_LOCAL_POPPLER_BIN = _BASE_DIR / "poppler" / "bin"

if _LOCAL_TESS_EXE.exists():
    pytesseract.pytesseract.tesseract_cmd = str(_LOCAL_TESS_EXE)
else:
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

POPPLER_BIN = str(_LOCAL_POPPLER_BIN) if _LOCAL_POPPLER_BIN.exists() else None

DATE_RE = re.compile(r"\b\d{4}-\d{2}-\d{2}\b")
CODE_RE = re.compile(r"(?<![A-Z0-9])[1I][A-Z]{4,6}\d{3,}[A-Z]?(?![A-Z0-9])", re.IGNORECASE)
RESULT_RE = re.compile(r"\b(P|F|A|W|X|NE)\b", re.IGNORECASE)


def normalize_code(raw_code):
    cleaned = re.sub(r"[^A-Za-z0-9]", "", raw_code.upper())
    cleaned = re.sub(r"^1I(?=B)", "1", cleaned)
    if cleaned.startswith("I"):
        cleaned = "1" + cleaned[1:]
    cleaned = cleaned.replace("BATA", "BAIA")
    cleaned = cleaned.replace("BAIJA", "BAIA")
    cleaned = cleaned.replace("BAJA", "BAIA")

    if cleaned.startswith("1B") and len(cleaned) >= 6:
        tail = cleaned[2:]
        m = re.match(r"^([A-Z0-9]+?)(\d{3})([A-Z]?)$", tail)
        if m:
            dept, digits, trailing = m.group(1), m.group(2), m.group(3)
            dept = (
                dept.replace("0", "O")
                    .replace("5", "S")
                    .replace("8", "B")
                    .replace("1", "I")
            )
            digits = (
                digits.replace("O", "0")
                      .replace("I", "1")
                      .replace("S", "5")
                      .replace("B", "8")
            )
            cleaned = "1B" + dept + digits + trailing
    return cleaned


_OCR_FP_MAP = str.maketrans({
    "O": "0",
    "I": "1",
    "l": "1",
    "S": "5",
    "B": "8",
    "Z": "2",
    "G": "0",
    "J": "I",
})


def _code_fingerprint(code: str) -> str:
    """Collapse OCR-confusable characters so variant OCR readings hash identically."""
    return code.upper().translate(_OCR_FP_MAP)


def _code_letter_quality(code: str) -> int:
    """Higher score = fewer digit artifacts in the department section (better display choice)."""
    m = re.search(r"^1B([A-Z0-9]+?)\d{3}[A-Z]?$", code)
    if m:
        return -sum(1 for c in m.group(1) if c.isdigit())
    return 0


def _subject_name_fingerprint(name: str) -> str:
    cleaned = re.sub(r"[^A-Z]", "", name.upper())
    cleaned = cleaned.replace("COMMUNICATIONSKILLS", "COMMUNICATIONSKILLS")
    return cleaned


def preprocess_image(page_image):
    img = np.array(page_image)
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    return cv2.adaptiveThreshold(
        gray,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        35,
        11,
    )


def extract_meta(text_blob):
    usn = re.search(r"(?:Seat\s+)?Number\s*[:>]?\s*([A-Z0-9]+)", text_blob, re.IGNORECASE)
    name = re.search(r"Student\s+Name\s*[:>]?\s*([^\n]+)", text_blob, re.IGNORECASE)
    return (
        usn.group(1).strip() if usn else "Unknown",
        name.group(1).strip() if name else "Unknown",
    )


def rows_from_plain_text(text):
    rows = []
    current = []
    for raw_line in text.splitlines():
        line = " ".join(raw_line.split())
        if not line:
            continue
        has_code = bool(CODE_RE.search(line))
        if has_code and current:
            rows.append(" ".join(current))
            current = []
        if has_code or current:
            current.append(line)
            if DATE_RE.search(line):
                rows.append(" ".join(current))
                current = []
    if current:
        rows.append(" ".join(current))
    return rows


def rows_from_data(preprocessed_image):
    data = pytesseract.image_to_data(
        preprocessed_image,
        config="--oem 3 --psm 6",
        output_type=pytesseract.Output.DICT,
    )
    grouped = defaultdict(list)
    for i, tok in enumerate(data["text"]):
        token = tok.strip()
        if not token:
            continue
        key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        grouped[key].append((data["left"][i], token))

    lines = [
        " ".join(token for _, token in sorted(tokens))
        for _, tokens in sorted(grouped.items())
    ]
    return rows_from_plain_text("\n".join(lines))


def parse_number_candidates(text):
    text = re.sub(r"(\d{1,3})\.\d+", r"\1", text)

    ocr_digit_map = str.maketrans(
        {
            "O": "0",
            "o": "0",
            "I": "1",
            "l": "1",
            "S": "5",
            "B": "8",
            "Z": "2",
            "g": "9",
            "q": "9",
        }
    )
    nums = []
    for token in re.findall(r"[A-Za-z0-9]+", text):
        fixed = token.translate(ocr_digit_map)
        digits = "".join(ch for ch in fixed if ch.isdigit())
        if not digits:
            continue
        if len(digits) == 4:
            left = int(digits[:2])
            right = int(digits[2:])
            if left <= 60 and right <= 120:
                nums.extend([left, right])
                continue
        if 1 <= len(digits) <= 3:
            value = int(digits)
            if 0 < value <= 200:
                nums.append(value)
    return nums


def infer_marks(nums):
    if not nums:
        return "", "", ""

    window = nums[-6:]
    best = None
    for i in range(len(window)):
        for j in range(i + 1, len(window)):
            for k in range(j + 1, len(window)):
                a, b, c = window[i], window[j], window[k]
                if a <= 60 and b <= 120 and c <= 200:
                    err = abs((a + b) - c)
                    score = (err, -k)
                    if best is None or score < best[0]:
                        best = (score, (a, b, c))

    if best and best[0][0] <= 5:
        a, b, c = best[1]
        return str(a), str(b), str(c)

    if len(window) >= 3:
        a, b, c = window[-3], window[-2], window[-1]
        return str(a), str(b), str(c)

    if len(window) == 2:
        a, b = window
        if a == b and a >= 35:
            return "", "", str(b)
        if a <= 60 and b <= 60:
            return str(a), str(b), str(a + b)
        if b >= a:
            return str(a), "", str(b)
        return "", str(a), str(b)

    return "", "", str(window[-1])


def recover_missing_marks(internal, external, total):
    """Fill likely missing fields when OCR captured only part of marks columns."""
    try:
        i = int(internal) if internal else None
        e = int(external) if external else None
        t = int(total) if total else None
    except ValueError:
        return internal, external, total

    if i is not None and e is not None and t is None:
        t = i + e

    if i is not None and e is None and t is not None and t >= i:
        e = t - i
    if e is not None and i is None and t is not None and t >= e:
        i = t - e

    if i is None and e is None and t is not None:
        i = t
        e = 0

    return (
        str(i) if i is not None else "",
        str(e) if e is not None else "",
        str(t) if t is not None else "",
    )


def parse_row(row_text):
    code_match = CODE_RE.search(row_text)
    if not code_match:
        return None

    raw_code = code_match.group(0)
    subject_code = normalize_code(raw_code)
    date_match = DATE_RE.search(row_text)
    result_match = RESULT_RE.search(row_text)

    cleaned = row_text
    cleaned = cleaned.replace(raw_code, " ")
    if date_match:
        cleaned = cleaned.replace(date_match.group(0), " ")

    result = result_match.group(1).upper() if result_match else ""
    if result_match:
        cleaned = re.sub(rf"\b{re.escape(result_match.group(1))}\b", " ", cleaned, flags=re.IGNORECASE)

    first_num = re.search(r"\d", cleaned)
    if first_num:
        subject_name = cleaned[: first_num.start()]
        marks_region = cleaned[first_num.start() :]
    else:
        subject_name = cleaned
        marks_region = ""

    subject_name = re.sub(r"\s+", " ", subject_name).replace("|", " ").strip(" -_:")
    nums = parse_number_candidates(marks_region)
    internal, external, total = infer_marks(nums)
    internal, external, total = recover_missing_marks(internal, external, total)

    if len(nums) == 1 and internal and total and not external:
        external = "0"

    return {
        "Subject Code": subject_code,
        "Subject Name": subject_name,
        "Internal": internal,
        "External": external,
        "Total": total,
        "Result": result,
        "_ObservedNumCount": len(nums),
    }


def candidate_score(parsed):
    score = 0
    if parsed["Subject Name"]:
        score += 2
    if parsed["Result"]:
        score += 2
    if parsed["Internal"]:
        score += 1
    if parsed["External"]:
        score += 1
    if parsed["Total"]:
        score += 1
    if parsed["Subject Code"].startswith("1B"):
        score += 1

    observed_num_count = parsed.get("_ObservedNumCount", 0)
    if observed_num_count >= 3:
        score += 4
    elif observed_num_count == 2:
        score += 1
    elif observed_num_count == 1:
        score -= 1

    try:
        i = int(parsed["Internal"]) if parsed["Internal"] else None
        e = int(parsed["External"]) if parsed["External"] else None
        t = int(parsed["Total"]) if parsed["Total"] else None
    except ValueError:
        return score

    if i is not None and e is not None and t is not None:
        err = abs((i + e) - t)
        score += max(0, 6 - err)
        if t < max(i, e):
            score -= 3
        if i <= 60 and e <= 60 and t == i + e:
            score += 3
    elif t is not None:
        score += 1

    if t is not None:
        if t >= 35:
            score += 2
        else:
            score -= 2

    return score


def parse_scanned_vtu(pdf_path):
    convert_kwargs = {"dpi": 300}
    if POPPLER_BIN:
        convert_kwargs["poppler_path"] = POPPLER_BIN
    pages = convert_from_path(pdf_path, **convert_kwargs)
    raw_text_pages = []
    all_rows = []

    for page in pages:
        raw_text = pytesseract.image_to_string(page, config="--oem 3 --psm 6")
        pre = preprocess_image(page)
        pre_text = pytesseract.image_to_string(pre, config="--oem 3 --psm 6")

        raw_text_pages.append(raw_text)
        raw_text_pages.append(pre_text)

        all_rows.extend(rows_from_plain_text(raw_text))
        all_rows.extend(rows_from_plain_text(pre_text))
        all_rows.extend(rows_from_data(pre))

    usn_val, name_val = extract_meta("\n".join(raw_text_pages))

    best_by_code = OrderedDict()
    for row in all_rows:
        parsed = parse_row(row)
        if not parsed:
            continue
        code = parsed["Subject Code"]
        quality = candidate_score(parsed)
        if code not in best_by_code or quality > best_by_code[code][0]:
            best_by_code[code] = (quality, parsed)

    fp_score: dict = {}
    fp_display_code: dict = {}
    for code, (score, parsed) in best_by_code.items():
        fp = (_code_fingerprint(code), _subject_name_fingerprint(parsed["Subject Name"]))
        if fp not in fp_score or score > fp_score[fp][0]:
            fp_score[fp] = (score, parsed)
        lq = _code_letter_quality(code)
        if fp not in fp_display_code or lq > fp_display_code[fp][0]:
            fp_display_code[fp] = (lq, code)

    data = []
    for fp, (score, parsed) in fp_score.items():
        parsed["Subject Code"] = fp_display_code[fp][1]
        parsed["USN"] = usn_val
        parsed["Name"] = name_val
        data.append(parsed)
    cols = ["USN", "Name", "Subject Code", "Subject Name", "Internal", "External", "Total", "Result"]
    normalized = []
    for row in data:
        normalized.append({col: row.get(col, "") for col in cols})
    return normalized


if __name__ == "__main__":
    file_to_read = "001.pdf"
    all_data = parse_scanned_vtu(file_to_read)

    if all_data:
        df = pd.DataFrame(all_data)
        output_file = "Final_Results.xlsx"
        try:
            df.to_excel(output_file, index=False)
        except PermissionError:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"Final_Results_{stamp}.xlsx"
            df.to_excel(output_file, index=False)
            print("Final_Results.xlsx is open/locked, wrote to fallback file instead.")
        print(f"Success! Extracted {len(df)} subjects from {file_to_read} -> {output_file}")
    else:
        print("No subjects found. Verify the scan quality and Tesseract path.")