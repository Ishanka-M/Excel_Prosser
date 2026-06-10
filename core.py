"""Core logic — ALL cells -> TEXT, displayed value EXACTLY preserved.

System sheet  : හැම cell එකක්ම text ('@'). Displayed value එකම text ලෙස
                (dates -> number_format එකට අනුව, numbers -> format respect,
                 strings -> hidden chars විතරක් අයින්). QUANTITY/Actual Qty -> scale.
Physical sheet: original verbatim copy (value + format + style).
"""
import io
import re
from copy import copy
from datetime import datetime, date, time

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

TARGET_HEADERS = {"QUANTITY", "ACTUAL QTY"}
HEADER_SCAN_ROWS = 30


# --------------------------- utils ---------------------------

def normalize_header(text) -> str:
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text)).strip().upper()


def clean_str(s: str) -> str:
    """Hidden/zero-width/control chars විතරක් අයින් — visible value නොවෙනස්ව."""
    s = s.replace("\xa0", " ").replace("\u200b", "").replace("\ufeff", "")
    s = "".join(ch for ch in s if ch in ("\t", " ") or ord(ch) >= 32)
    return s.strip()


def to_number(v):
    if v is None or v == "" or isinstance(v, bool):
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
    n = float(n)
    if n.is_integer():
        return str(int(round(n)))
    return ("%f" % n).rstrip("0").rstrip(".")


def count_int_digits(n) -> int:
    return len(str(abs(int(round(float(n))))))


# --------------------------- faithful number rendering ---------------------------

def render_number(v, fmt) -> str:
    f0 = (fmt or "General").split(";")[0]
    core = re.sub(r"\[[^\]]*\]", "", f0).replace('"', "").replace("\\", "")
    has_comma = "," in core
    if "." in core and any(ch in core.split(".", 1)[1] for ch in "0#"):
        dec = sum(1 for ch in core.split(".", 1)[1] if ch in "0#")
        return f"{v:,.{dec}f}" if has_comma else f"{v:.{dec}f}"
    if float(v).is_integer():
        iv = int(round(v))
        zeros = core.count("0")
        if zeros > 1 and set(core) <= set("0,#"):
            sign = "-" if iv < 0 else ""
            return f"{sign}{abs(iv):0{zeros}d}"
        return f"{iv:,}" if has_comma else str(iv)
    return ("%f" % v).rstrip("0").rstrip(".")


# --------------------------- faithful date rendering ---------------------------

MONTHS = ["", "January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]
WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


def _tokenize_dt_format(fmt):
    fmt = re.sub(r"\[[^\]]*\]", "", fmt.split(";")[0])
    tokens, i, n = [], 0, len(fmt)
    while i < n:
        ch = fmt[i]
        if ch == '"':
            j = fmt.find('"', i + 1)
            j = n if j == -1 else j
            tokens.append(("lit", fmt[i + 1:j])); i = j + 1; continue
        if ch == "\\":
            if i + 1 < n:
                tokens.append(("lit", fmt[i + 1])); i += 2
            else:
                i += 1
            continue
        up = fmt[i:i + 5].upper()
        if up == "AM/PM":
            tokens.append(("ap", 5)); i += 5; continue
        if fmt[i:i + 3].upper() == "A/P":
            tokens.append(("ap", 3)); i += 3; continue
        cl = ch.lower()
        if cl in "ymdhs":
            j = i
            while j < n and fmt[j].lower() == cl:
                j += 1
            tokens.append((cl, j - i)); i = j; continue
        tokens.append(("lit", ch)); i += 1
    return tokens


def _m_is_minute(tokens, idx):
    for k in range(idx - 1, -1, -1):
        t = tokens[k][0]
        if t == "h":
            return True
        if t in ("y", "d", "m", "s"):
            return False
    for k in range(idx + 1, len(tokens)):
        t = tokens[k][0]
        if t == "s":
            return True
        if t in ("y", "d", "m", "h"):
            return False
    return False


def render_excel_date(dt, fmt) -> str:
    """datetime/date/time එක Excel number_format එකට අනුව faithful string."""
    if not fmt or fmt.lower() == "general":
        if isinstance(dt, datetime) and (dt.hour or dt.minute or dt.second):
            return dt.strftime("%m/%d/%Y %H:%M:%S")
        if isinstance(dt, (datetime, date)):
            return dt.strftime("%m/%d/%Y")
        return dt.strftime("%H:%M:%S")
    tokens = _tokenize_dt_format(fmt)
    has_ap = any(t == "ap" for t, _ in tokens)
    hh, mm, ss = getattr(dt, "hour", 0), getattr(dt, "minute", 0), getattr(dt, "second", 0)
    yy, mo, dd = getattr(dt, "year", 1900), getattr(dt, "month", 1), getattr(dt, "day", 1)
    wd = dt.weekday() if isinstance(dt, (datetime, date)) else 0
    out = []
    for idx, (t, cnt) in enumerate(tokens):
        if t == "lit":
            out.append(cnt if isinstance(cnt, str) else "")
        elif t == "y":
            out.append(f"{yy:04d}" if cnt >= 4 else f"{yy % 100:02d}")
        elif t == "d":
            if cnt >= 4:
                out.append(WEEKDAYS[wd])
            elif cnt == 3:
                out.append(WEEKDAYS[wd][:3])
            elif cnt == 2:
                out.append(f"{dd:02d}")
            else:
                out.append(str(dd))
        elif t == "h":
            h = hh % 12 if has_ap else hh
            if has_ap and h == 0:
                h = 12
            out.append(f"{h:02d}" if cnt >= 2 else str(h))
        elif t == "s":
            out.append(f"{ss:02d}" if cnt >= 2 else str(ss))
        elif t == "ap":
            out.append(("AM" if hh < 12 else "PM") if cnt == 5
                       else ("A" if hh < 12 else "P"))
        elif t == "m":
            if _m_is_minute(tokens, idx):
                out.append(f"{mm:02d}" if cnt >= 2 else str(mm))
            else:
                if cnt >= 4:
                    out.append(MONTHS[mo])
                elif cnt == 3:
                    out.append(MONTHS[mo][:3])
                elif cnt == 2:
                    out.append(f"{mo:02d}")
                else:
                    out.append(str(mo))
    return "".join(out)


# --------------------------- faithful text (any value) ---------------------------

def faithful_text(value, number_format) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, str):
        return clean_str(value)
    if isinstance(value, (datetime, date, time)):
        return render_excel_date(value, number_format)
    if isinstance(value, (int, float)):
        return render_number(float(value), number_format)
    return clean_str(str(value))


# --------------------------- sheet naming ---------------------------

def safe_title(title, used):
    for ch in "[]:*?/\\":
        title = title.replace(ch, " ")
    title = re.sub(r"\s+", " ", title).strip()[:31] or "Sheet"
    base, i = title, 2
    while title.lower() in used:
        suffix = f" ({i})"
        title = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(title.lower())
    return title


# --------------------------- copies ---------------------------

def _copy_dims_merges(dst, src):
    for mc in src.merged_cells.ranges:
        try:
            dst.merge_cells(str(mc))
        except Exception:
            pass
    for col, dim in src.column_dimensions.items():
        if dim.width:
            dst.column_dimensions[col].width = dim.width


def copy_sheet_exact(dst, src):
    """Physical: verbatim (value + number_format + style)."""
    for row in src.iter_rows():
        for cell in row:
            if cell.value is None and not cell.has_style:
                continue
            nc = dst.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                nc.font = copy(cell.font)
                nc.fill = copy(cell.fill)
                nc.border = copy(cell.border)
                nc.alignment = copy(cell.alignment)
                nc.protection = copy(cell.protection)
                nc.number_format = cell.number_format
    _copy_dims_merges(dst, src)


def fill_text_sheet(dst, src, multiplier, name, scale):
    """System / copied sheet: හැම cell එකක්ම faithful TEXT; scale නම් qty columns scale."""
    rows = list(src.iter_rows())
    header_idx, target_cols = None, {}
    if scale:
        for i, row in enumerate(rows[:HEADER_SCAN_ROWS]):
            found = {c.column: normalize_header(c.value) for c in row
                     if normalize_header(c.value) in TARGET_HEADERS}
            if found:
                header_idx, target_cols = i, found
                break

    warnings, scaled = [], 0
    for i, row in enumerate(rows):
        for cell in row:
            is_qty = (scale and header_idx is not None
                      and cell.column in target_cols and i > header_idx)
            if is_qty and to_number(cell.value) is not None:
                num = to_number(cell.value)
                val = num * multiplier
                txt = num_to_text(val)
                is_dec = abs(val - round(val)) > 1e-9
                digits = count_int_digits(val)
                issue = None
                if is_dec:
                    issue = "Decimal (multiplier මදි)"
                elif digits <= 3:
                    issue = f"Digit {digits}ක් — 3ට වඩා වැඩි විය යුතුයි"
                if issue:
                    warnings.append({
                        "sheet": name,
                        "cell": f"{get_column_letter(cell.column)}{cell.row}",
                        "header": target_cols[cell.column],
                        "original": num_to_text(num), "scaled": txt, "issue": issue,
                    })
                scaled += 1
            else:
                txt = faithful_text(cell.value, cell.number_format)

            if txt != "":
                nc = dst.cell(row=cell.row, column=cell.column, value=txt)
                nc.number_format = "@"
    _copy_dims_merges(dst, src)
    return warnings, scaled, bool(target_cols), target_cols


# --------------------------- orchestration ---------------------------

def build_output(file_bytes, selected, multiplier, include_unmarked):
    src = load_workbook(io.BytesIO(file_bytes), data_only=True)
    selected_set = set(selected)
    out = Workbook()
    out.remove(out.active)
    used = set()
    summary, warnings, physical_specs = [], [], []

    for name in src.sheetnames:
        marked = name in selected_set
        if not marked and not include_unmarked:
            continue
        ws = src[name]
        if marked:
            sys_ws = out.create_sheet(title=safe_title(f"System {name}", used))
            w, scaled, found, tcols = fill_text_sheet(sys_ws, ws, multiplier, name, scale=True)
            warnings.extend(w)
            physical_specs.append((ws, name))
            qty_status = ("✅ " + ", ".join(tcols.values())) if found else "⚠️ Not found"
            summary.append({"Sheet": name, "Qty column": qty_status,
                            "Scaled": scaled, "⚠": len(w)})
        else:
            cl_ws = out.create_sheet(title=safe_title(name, used))
            fill_text_sheet(cl_ws, ws, multiplier, name, scale=False)
            summary.append({"Sheet": name, "Qty column": "— clean-only (text)",
                            "Scaled": 0, "⚠": 0})

    for ws, name in physical_specs:
        ph_ws = out.create_sheet(title=safe_title(f"Physical {name}", used))
        copy_sheet_exact(ph_ws, ws)

    src.close()
    if not out.sheetnames:
        out.create_sheet(title="Sheet1")
    bio = io.BytesIO()
    out.save(bio)
    bio.seek(0)
    total_scaled = sum(r["Scaled"] for r in summary)
    return summary, warnings, bio.getvalue(), total_scaled, out.sheetnames
