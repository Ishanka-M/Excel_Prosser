"""Core logic for Excel Cleaner & Quantity Scaler (UI-independent, testable)."""
import io
import re
import unicodedata
from datetime import datetime, date

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

TARGET_HEADERS = {"QUANTITY", "ACTUAL QTY"}
HEADER_SCAN_ROWS = 30

# ----------------------------- Helpers -----------------------------

def normalize_header(text) -> str:
    """Header compare කරන්න: whitespace collapse + uppercase."""
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text)).strip().upper()


def clean_text(v) -> str:
    """Cell value එකක් clean text එකකට convert කරනවා.

    - hidden / control / zero-width chars remove
    - non-breaking space -> normal space
    - whitespace collapse + strip
    - numbers/dates string වලට convert
    """
    if v is None:
        return ""
    if isinstance(v, bool):
        s = str(v)
    elif isinstance(v, int):
        s = str(v)
    elif isinstance(v, float):
        s = str(int(v)) if v.is_integer() else ("%f" % v).rstrip("0").rstrip(".")
    elif isinstance(v, (datetime, date)):
        s = v.isoformat()
    else:
        s = str(v)

    # non-breaking space සහ zero-width chars
    s = s.replace("\xa0", " ").replace("\u200b", "").replace("\ufeff", "")
    # control chars (category C*) remove — tab/space පමණක් තියාගන්නවා
    s = "".join(ch for ch in s if ch in ("\t", " ") or not unicodedata.category(ch).startswith("C"))
    # whitespace collapse + strip
    s = re.sub(r"\s+", " ", s).strip()
    return s


def to_number(v):
    """Value එකක් number එකක් නම් float return කරනවා, නැත්නම් None."""
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def num_to_text(n) -> str:
    """Number එකක් clean text එකකට (trailing .0 නැතුව)."""
    n = float(n)
    if n.is_integer():
        return str(int(round(n)))
    return ("%f" % n).rstrip("0").rstrip(".")


def count_int_digits(n) -> int:
    """Integer part එකේ digit ගණන."""
    return len(str(abs(int(round(float(n))))))


# ----------------------------- Core processing -----------------------------

def process_sheet(rows, multiplier, scale=True):
    """එක sheet එකක rows (list of tuples) process කරනවා.

    scale=False නම් clean පමණයි (QUANTITY/Actual Qty scale කරන්නේ නැහැ) — unmarked sheets වලට.

    Returns: (cleaned_grid, header_idx, target_cols, warnings, scaled_count, target_found)
    """
    # 1) header row එක හොයනවා (QUANTITY / Actual Qty තියෙන පළවෙනි row)
    header_idx = None
    target_cols = {}  # col_index -> header label
    if scale:
        for i, row in enumerate(rows[:HEADER_SCAN_ROWS]):
            norm = [normalize_header(c) for c in row]
            found = {j: norm[j] for j in range(len(norm)) if norm[j] in TARGET_HEADERS}
            if found:
                header_idx = i
                target_cols = found
                break

    # 2) හැම cell එකක්ම clean text එකට
    cleaned = [[clean_text(c) for c in row] for row in rows]

    warnings = []
    scaled_count = 0

    if scale and header_idx is not None:
        for r in range(header_idx + 1, len(rows)):
            for col_idx, label in target_cols.items():
                if col_idx >= len(rows[r]):
                    continue
                orig = rows[r][col_idx]
                num = to_number(orig)
                if num is None:
                    continue  # හිස් / number නොවන cell skip

                scaled = num * multiplier

                # ---- validation (final scaled value එකට) ----
                is_decimal = abs(scaled - round(scaled)) > 1e-9
                digits = count_int_digits(scaled)
                issue = None
                if is_decimal:
                    issue = "Decimal value (multiplier මදි / decimal clear වෙලා නැහැ)"
                elif digits <= 3:
                    issue = f"Digit {digits}ක් පමණයි — 3ට වඩා වැඩි විය යුතුයි"

                if issue:
                    warnings.append({
                        "cell": f"{get_column_letter(col_idx + 1)}{r + 1}",
                        "header": label,
                        "original": num_to_text(num),
                        "scaled": num_to_text(scaled),
                        "issue": issue,
                    })

                # scaled value එක text එකක් විදිහට in-place replace
                cleaned[r][col_idx] = num_to_text(scaled)
                scaled_count += 1

    return cleaned, header_idx, target_cols, warnings, scaled_count, bool(target_cols)


def build_workbook(processed_sheets):
    """processed_sheets: dict[sheet_name] = cleaned_grid -> xlsx bytes (all cells TEXT)."""
    wb = Workbook()
    wb.remove(wb.active)
    for sheet_name, grid in processed_sheets.items():
        ws = wb.create_sheet(title=sheet_name[:31])  # Excel sheet name limit 31
        for r, row in enumerate(grid, start=1):
            for c, val in enumerate(row, start=1):
                cell = ws.cell(row=r, column=c, value=val)
                cell.number_format = "@"  # TEXT format
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


