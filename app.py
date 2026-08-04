"""
Excel Cleaner & Quantity Scaler — single-file Streamlit app (self-contained).
Run:  python -m streamlit run app.py

STABILITY UPDATE v2 (logic 100% same — output Excel එක නොවෙනස්):
  * Output workbook එක RAM එකේ cache වෙන්නේ නෑ — disk temp file එකකට save (OOM #1 හේතුව)
  * Cache keys = file digest (MB ගණන් bytes හැම rerun එකකම hash වෙන එක නවතී)
  * Workbook close + gc.collect() → RAM ආපහු release
  * st.fragment / st.container(border) version-safe → පරණ Streamlit එකකත් "Oh no" නෑ
  * Processing try/except → error එකකින් app crash වෙන්නේ නෑ
  * Presence registry capped + heartbeat 5s → 10s (websocket load අඩුයි)

v4 (UI + multi-user layer) — processing logic byte-for-byte same:
  * Professional UI (masthead, numbered steps, refined metric cards)
  * User management: name, online/busy list, job queue (semaphore)
  * JSON process log (data/process_log.json) — line counts per run/sheet
  * Temp files per-user + auto-prune, dataframe fallback, render try-except
"""
import io
import os
import re
import gc
import time
import uuid
import atexit
import hashlib
import tempfile
import threading
from copy import copy
from datetime import datetime, date, time as _time

import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# ===================== CORE LOGIC (inlined) =====================
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
    if isinstance(value, (datetime, date, _time)):
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


# professional table theme
_HDR_FILL = PatternFill("solid", fgColor="1F4E78")
_HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
_BAND_FILL = PatternFill("solid", fgColor="EEF3FA")
_THIN = Side(style="thin", color="D9D9D9")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_AL_CENTER = Alignment(horizontal="center", vertical="center")
_AL_LEFT = Alignment(horizontal="left", vertical="center")
_AL_RIGHT = Alignment(horizontal="right", vertical="center")


def build_physical_pro(dst, src):
    """Physical sheet — original values + number_format තියාගෙන, professional table styling.

    styled header + borders + banded rows + auto width + freeze + autofilter.
    """
    rows = list(src.iter_rows())
    cells_with_val = [c for row in rows for c in row if c.value is not None]
    if not cells_with_val:
        copy_sheet_exact(dst, src)
        return

    max_col = max(c.column for c in cells_with_val)
    last_row = max(c.row for c in cells_with_val)

    # header row = මුල් dense row එක (>=3 non-empty), නැත්නම් මුල් non-empty row
    header_row, first_nonempty = None, None
    for r_idx, row in enumerate(rows, start=1):
        ne = sum(1 for c in row if c.value not in (None, ""))
        if ne and first_nonempty is None:
            first_nonempty = r_idx
        if ne >= 3:
            header_row = r_idx
            break
    header_row = header_row or first_nonempty or 1

    # 1) values + number_format + width tracking
    col_w = {}
    for cell in cells_with_val:
        nc = dst.cell(row=cell.row, column=cell.column, value=cell.value)
        nc.number_format = cell.number_format
        txt = faithful_text(cell.value, cell.number_format)
        if txt:
            col_w[cell.column] = max(col_w.get(cell.column, 0), len(txt))

    # 2) style the table rectangle [header_row..last_row] x [1..max_col]
    for r in range(header_row, last_row + 1):
        for c in range(1, max_col + 1):
            nc = dst.cell(row=r, column=c)
            nc.border = _BORDER
            if r == header_row:
                nc.fill = _HDR_FILL
                nc.font = _HDR_FONT
                nc.alignment = _AL_CENTER
            else:
                if (r - header_row) % 2 == 0:
                    nc.fill = _BAND_FILL
                nc.alignment = _AL_RIGHT if to_number(nc.value) is not None else _AL_LEFT

    # 3) widths / heights
    for col, ln in col_w.items():
        dst.column_dimensions[get_column_letter(col)].width = min(max(10, ln + 2), 45)
    dst.row_dimensions[header_row].height = 20

    # 4) freeze + autofilter
    dst.freeze_panes = dst.cell(row=header_row + 1, column=1)
    dst.auto_filter.ref = f"A{header_row}:{get_column_letter(max_col)}{last_row}"

    # 5) merges (title rows etc.)
    for mc in src.merged_cells.ranges:
        try:
            dst.merge_cells(str(mc))
        except Exception:
            pass

    # memory: loop එක ඉවර නම් references අත්හරිනවා
    rows.clear()
    cells_with_val.clear()


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
            is_qty = (scale and multiplier is not None and header_idx is not None
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
    rows.clear()
    return warnings, scaled, bool(target_cols), target_cols


# --------------------------- orchestration ---------------------------

def build_output(file_bytes, selected, multiplier, include_unmarked, out_path):
    src = load_workbook(io.BytesIO(file_bytes), data_only=True, keep_links=False)
    out = Workbook()
    try:
        selected_set = set(selected)
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
            build_physical_pro(ph_ws, ws)

        physical_specs.clear()

        if not out.sheetnames:
            out.create_sheet(title="Sheet1")
        # RAM එකට bytes ගන්නේ නෑ — කෙලින්ම disk එකට save (memory spike එක අඩකින් අඩුයි)
        out.save(out_path)
        total_scaled = sum(r["Scaled"] for r in summary)
        sheet_order = list(out.sheetnames)
        return summary, warnings, out_path, total_scaled, sheet_order
    finally:
        # RAM release — මේක නැත්නම් file කීපයකින් පස්සේ process එක OOM වෙලා down වෙනවා
        try:
            src.close()
        except Exception:
            pass
        try:
            out.close()
        except Exception:
            pass
        del src, out
        gc.collect()


# ===================== STREAMLIT APP =====================
import json

st.set_page_config(page_title="Excel Cleaner & Qty Scaler", page_icon="📊",
                   layout="centered", initial_sidebar_state="expanded")

CSS = """
<style>
.block-container{max-width:1000px;padding-top:1.2rem;padding-bottom:3.5rem;}
#MainMenu, footer{visibility:hidden;}

/* ---------- masthead ---------- */
.masthead{
  background:#0F172A;border:1px solid #1E293B;border-left:4px solid #3B82F6;
  border-radius:12px;padding:20px 24px;margin-bottom:20px;
}
.masthead .kicker{
  color:#60A5FA;font-size:.68rem;font-weight:700;letter-spacing:.16em;
  text-transform:uppercase;margin:0 0 6px;
}
.masthead h1{color:#F8FAFC;font-size:1.35rem;font-weight:650;margin:0;letter-spacing:-.01em;}
.masthead p{color:#94A3B8;font-size:.85rem;margin:.35rem 0 0;line-height:1.5;}

/* ---------- step headers ---------- */
.step{display:flex;align-items:center;gap:.6rem;margin:0 0 .85rem;}
.step .n{
  width:22px;height:22px;border-radius:6px;background:#3B82F6;color:#fff;
  font-size:.72rem;font-weight:700;display:inline-flex;align-items:center;
  justify-content:center;flex:0 0 auto;
}
.step .t{font-size:.76rem;font-weight:700;letter-spacing:.1em;
  text-transform:uppercase;opacity:.75;}
.step .hint{margin-left:auto;font-size:.74rem;opacity:.55;font-weight:500;}

/* ---------- metrics ---------- */
[data-testid="stMetric"]{
  background:rgba(148,163,184,.08);border:1px solid rgba(148,163,184,.2);
  border-radius:10px;padding:12px 16px;
}
[data-testid="stMetricValue"]{font-size:1.4rem;font-weight:650;letter-spacing:-.02em;}
[data-testid="stMetricLabel"]{opacity:.65;font-size:.74rem;font-weight:600;
  letter-spacing:.04em;text-transform:uppercase;}

/* ---------- controls ---------- */
.stButton>button,.stDownloadButton>button{border-radius:8px;font-weight:600;font-size:.87rem;}
.stDownloadButton>button{width:100%;padding:.6rem;}
div[data-testid="stFileUploader"] section{border-radius:10px;border-style:dashed;}
section[data-testid="stSidebar"] [data-testid="stMetric"]{text-align:center;}
.stCheckbox label p{font-size:.86rem;}
hr{margin:.9rem 0;}

/* ---------- footnote ---------- */
.foot{font-size:.75rem;opacity:.5;text-align:center;margin-top:1.6rem;}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

PRESENCE_WINDOW = 30      # තත්පර — මේ ඇතුළත heartbeat ආපු session = online
PRESENCE_BEAT = 10        # heartbeat interval (websocket traffic අඩුයි)
MAX_SESSIONS = 500        # registry unbounded වෙලා memory කන එක නවත්වන්න
BIG_FILE_MB = 20          # මීට වඩා ලොකු file එකකදී warning
CACHE_TTL = 1800          # තත්පර 30min
TMP_MAX_AGE = 7200        # පැය 2කට වඩා පරණ temp output files අයින්
MAX_CONCURRENT_JOBS = 2   # එකවර process වෙන්න දෙන jobs — RAM එක බේරෙනවා
QUEUE_WAIT_SEC = 300      # queue එකේ බලාගෙන ඉන්න max තත්පර
LOG_KEEP_RUNS = 200       # JSON log එකේ තියාගන්න runs ගාණ


# ----------------------- small UI helpers (version-safe) -----------------------

def step(n, title, hint=""):
    st.markdown(
        f'<div class="step"><span class="n">{n}</span><span class="t">{title}</span>'
        f'<span class="hint">{hint}</span></div>',
        unsafe_allow_html=True,
    )


def box():
    """st.container(border=...) පරණ version වල නෑ — safe wrapper."""
    try:
        return st.container(border=True)
    except TypeError:
        return st.container()


def safe_table(rows, column_config=None):
    """st.column_config පරණ Streamlit වල නෑ — error එකකින් page එක වැටෙන්නේ නෑ."""
    try:
        st.dataframe(rows, use_container_width=True, hide_index=True,
                     column_config=column_config)
    except Exception:
        try:
            st.dataframe(rows, use_container_width=True)
        except Exception:
            st.table(rows)


def notify(msg):
    if hasattr(st, "toast"):
        st.toast(msg)
    else:
        st.success(msg)


# ----------------------- Online presence (shared across users) -----------------------

@st.cache_resource
def _presence_registry():
    """හැම user session එකකම පොදු (shared) registry එක — process එකට එකයි."""
    return {"sessions": {}, "lock": threading.Lock()}


@st.cache_resource
def _job_registry():
    """Concurrency control — එකවර process වෙන jobs ගාණ සීමා කරනවා.

    Users 5ක් එකවර ලොකු file process කරොත් RAM ඉවර වෙලා container එක මැරෙනවා.
    Semaphore එකෙන් එකවර MAX_CONCURRENT_JOBS ගාණක් විතරක් යනවා, ඉතුරු අය queue එකේ.
    """
    return {
        "sem": threading.BoundedSemaphore(MAX_CONCURRENT_JOBS),
        "active": {},                    # sid -> {"name", "file", "started"}
        "lock": threading.Lock(),
        "log_lock": threading.Lock(),
    }


def my_sid():
    return st.session_state.setdefault("_sid", str(uuid.uuid4()))


def my_name():
    return st.session_state.get("_user_name") or f"User-{my_sid()[:4]}"


def heartbeat(window=PRESENCE_WINDOW):
    """Session එක alive කියලා ලකුණු කරලා, දැන් online අය ලැයිස්තුව දෙනවා."""
    try:
        reg = _presence_registry()
        jobs = _job_registry()
        sid, now = my_sid(), time.time()
        with reg["lock"]:
            reg["sessions"][sid] = {"t": now, "name": my_name()}
            stale = [k for k, v in reg["sessions"].items() if now - v["t"] > window]
            for k in stale:
                reg["sessions"].pop(k, None)
            # safety cap — කවදාවත් unbounded වෙන්නේ නෑ
            if len(reg["sessions"]) > MAX_SESSIONS:
                ordered = sorted(reg["sessions"].items(), key=lambda kv: kv[1]["t"])
                for k, _ in ordered[:-MAX_SESSIONS]:
                    reg["sessions"].pop(k, None)
            people = [(v["name"], k) for k, v in reg["sessions"].items()]
        with jobs["lock"]:
            busy = set(jobs["active"].keys())
        return [{"name": nm, "busy": k in busy, "me": k == sid}
                for nm, k in sorted(people)]
    except Exception:
        return [{"name": "—", "busy": False, "me": True}]


def _online_badge_impl():
    try:
        people = heartbeat()
        working = sum(1 for p in people if p["busy"])
        c1, c2 = st.columns(2)
        c1.metric("Online", len(people))
        c2.metric("Processing", working)
        lines = []
        for p in people[:8]:
            dot = "🟠" if p["busy"] else "🟢"
            lines.append(f"{dot} {p['name']}" + (" *(you)*" if p["me"] else ""))
        if len(people) > 8:
            lines.append(f"…+{len(people) - 8}")
        st.caption("\n\n".join(lines))
    except Exception:
        st.metric("Online", "—")


# Version-safe: පරණ Streamlit එකක st.fragment නැති නිසා import-time එකේම
# AttributeError → "Oh no. Error running app" වෙනවා. දැන් fallback එකක් තියෙනවා.
if hasattr(st, "fragment"):
    online_badge = st.fragment(run_every=PRESENCE_BEAT)(_online_badge_impl)
elif hasattr(st, "experimental_fragment"):
    online_badge = st.experimental_fragment(run_every=PRESENCE_BEAT)(_online_badge_impl)
else:
    online_badge = _online_badge_impl


# ----------------------- Cached read + disk-backed processing -----------------------
# cache key එකට digest එක විතරයි (bytes underscore-prefixed → hash වෙන්නේ නෑ).
# මේකෙන් හැම rerun එකකදීම MB ගණන් bytes hash කරන CPU/RAM spike එක නවතිනවා.

@st.cache_data(show_spinner=False, max_entries=3, ttl=CACHE_TTL)
def read_sheet_names(digest: str, _file_bytes: bytes):
    wb = load_workbook(io.BytesIO(_file_bytes), read_only=True, data_only=True, keep_links=False)
    try:
        return list(wb.sheetnames)
    finally:
        try:
            wb.close()
        except Exception:
            pass
        gc.collect()


_TMP_DIR = os.path.join(tempfile.gettempdir(), "xl_cleaner")
try:
    os.makedirs(_TMP_DIR, exist_ok=True)
except Exception:
    _TMP_DIR = tempfile.gettempdir()


def _drop_temp(path):
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


def _prune_temp(max_age=TMP_MAX_AGE):
    """පරණ output files අයින් — disk එක පිරිලා app එක වැටෙන්නේ නෑ."""
    try:
        now = time.time()
        for f in os.listdir(_TMP_DIR):
            p = os.path.join(_TMP_DIR, f)
            if os.path.isfile(p) and now - os.path.getmtime(p) > max_age:
                _drop_temp(p)
    except Exception:
        pass


@atexit.register
def _cleanup_temp():
    try:
        for f in os.listdir(_TMP_DIR):
            _drop_temp(os.path.join(_TMP_DIR, f))
    except Exception:
        pass


# ----------------------- Line counting + JSON process log -----------------------

@st.cache_data(show_spinner=False, max_entries=6, ttl=CACHE_TTL)
def count_lines(digest: str, _file_bytes: bytes, sheets: tuple):
    """Process වෙන sheets වල data lines (හිස් නොවන rows) ගාණ.

    read_only + values_only — වේගවත්, memory අඩු. Output එකට කිසිම බලපෑමක් නෑ.
    """
    counts = {}
    wb = None
    try:
        wb = load_workbook(io.BytesIO(_file_bytes), read_only=True,
                           data_only=True, keep_links=False)
        for nm in sheets:
            if nm not in wb.sheetnames:
                continue
            n = 0
            for row in wb[nm].iter_rows(values_only=True):
                if any(v is not None and str(v).strip() != "" for v in row):
                    n += 1
            counts[nm] = n
    except Exception:
        pass
    finally:
        try:
            if wb is not None:
                wb.close()
        except Exception:
            pass
        gc.collect()
    return counts


def _resolve_data_dir():
    """Log file එක තියෙන තැන — app folder එක write කරන්න බැරි නම් temp එකට."""
    here = os.path.dirname(os.path.abspath(globals().get("__file__", "."))) or "."
    for base in (here, tempfile.gettempdir()):
        try:
            d = os.path.join(base, "data")
            os.makedirs(d, exist_ok=True)
            probe = os.path.join(d, ".w")
            with open(probe, "w") as f:
                f.write("1")
            os.remove(probe)
            return d
        except Exception:
            continue
    return tempfile.gettempdir()


_DATA_DIR = _resolve_data_dir()
LOG_PATH = os.path.join(_DATA_DIR, "process_log.json")


def _empty_log():
    return {"app": "Excel Cleaner & Quantity Scaler", "updated": None,
            "totals": {"runs": 0, "lines": 0, "sheets": 0, "scaled_cells": 0, "users": 0},
            "runs": []}


def load_log():
    try:
        with open(LOG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and "runs" in data:
            return data
    except Exception:
        pass
    return _empty_log()


def save_log(data):
    """Atomic write — users කීපදෙනෙක් එකවර ලිව්වත් file එක corrupt වෙන්නේ නෑ."""
    tmp = f"{LOG_PATH}.{uuid.uuid4().hex[:6]}.tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, LOG_PATH)
        return True
    except Exception:
        _drop_temp(tmp)
        return False


def append_run(record):
    """Run එකක් log එකට — lock එකක් යටතේ (multi-user safe)."""
    jobs = _job_registry()
    try:
        with jobs["log_lock"]:
            data = load_log()
            data["runs"].insert(0, record)
            del data["runs"][LOG_KEEP_RUNS:]
            t = {"runs": len(data["runs"]), "lines": 0, "sheets": 0,
                 "scaled_cells": 0, "users": 0}
            users = set()
            for r in data["runs"]:
                t["lines"] += r.get("total_lines", 0)
                t["sheets"] += len(r.get("sheets", []))
                t["scaled_cells"] += r.get("total_scaled", 0)
                users.add(r.get("user", "?"))
            t["users"] = len(users)
            data["totals"] = t
            data["updated"] = datetime.now().isoformat(timespec="seconds")
            save_log(data)
            return data
    except Exception:
        return load_log()


def run_processing(digest: str, file_bytes: bytes, selected: tuple,
                   multiplier, include_unmarked: bool):
    """Output එක RAM එකේ cache කරන්නේ නෑ — disk temp file එකකට save කරනවා.

    Marked sheet එකකට output එකේ:
      - "System <name>"   : faithful TEXT + QUANTITY/Actual Qty scale, original position
      - "Physical <name>" : original sheet එක verbatim (value+format+style), workbook අන්තිමට
    """
    _prune_temp()
    # per-user file name — user දෙන්නෙක් එකවර process කරාම එකිනෙකාගේ output එක overwrite වෙන්නේ නෑ
    out_path = os.path.join(
        _TMP_DIR, f"{my_sid()[:6]}_{digest[:8]}_{uuid.uuid4().hex[:6]}.xlsx")
    return build_output(file_bytes, selected, multiplier, include_unmarked, out_path)


def clear_last_result():
    old = st.session_state.pop("_last_result", None)
    if old:
        _drop_temp(old[2])
    st.session_state.pop("_last_meta", None)


# ----------------------------- Sidebar -----------------------------

with st.sidebar:
    st.markdown("##### Your name")
    st.text_input("Name", key="_user_name", placeholder=f"User-{my_sid()[:4]}",
                  label_visibility="collapsed",
                  help="Process log එකේ සහ online list එකේ පේන නම.")
    st.divider()
    st.markdown("##### Status")
    online_badge()
    st.caption(f"එකවර process වෙන්නේ jobs {MAX_CONCURRENT_JOBS}යි — ඉතුරු අය queue එකේ.")
    st.divider()
    st.markdown("##### How it works")
    st.caption(
        "1. Excel file එක upload කරන්න\n\n"
        "2. Process කරන sheets mark කරන්න\n\n"
        "3. Multiplier (×100 / ×1000) තෝරන්න\n\n"
        "4. Clean Excel download කරන්න"
    )
    st.divider()
    st.markdown("##### Output structure")
    st.caption(
        "**System &lt;name&gt;** — text convert + QTY scale\n\n"
        "**Physical &lt;name&gt;** — original values, styled table\n\n"
        "Marked sheet එකකට මේ දෙකම හැදෙනවා."
    )
    st.divider()
    if st.button("Clear cache / free memory", use_container_width=True):
        try:
            st.cache_data.clear()
        except Exception:
            pass
        clear_last_result()
        _prune_temp(0)
        gc.collect()
        notify("Memory නිදහස් කළා")


# ----------------------------- Main -----------------------------

st.markdown(
    '<div class="masthead">'
    '<p class="kicker">Warehouse Data Tools</p>'
    '<h1>Excel Cleaner &amp; Quantity Scaler</h1>'
    '<p>Upload → sheets mark කරන්න → QUANTITY / Actual Qty clean &amp; scale → download. '
    'Mark කරන හැම sheet එකකටම System (text) සහ Physical (original) sheet දෙකක් හැදෙනවා.</p>'
    '</div>',
    unsafe_allow_html=True,
)

# ---- Step 1: Upload ----
with box():
    step(1, "Upload workbook", ".xlsx / .xlsm")
    uploaded = st.file_uploader(
        "Excel file", type=["xlsx", "xlsm"], label_visibility="collapsed"
    )

if uploaded is None:
    st.info("පටන් ගන්න Excel file එකක් upload කරන්න.")
    st.markdown('<div class="foot">Multi-user ready · disk-backed output · no data stored</div>',
                unsafe_allow_html=True)
    st.stop()

try:
    file_bytes = uploaded.getvalue()
    digest = hashlib.md5(file_bytes).hexdigest()      # cache key (ලාබයි, එක පාරයි)
    size_mb = len(file_bytes) / (1024 * 1024)
except Exception as e:
    st.error(f"File එක කියවන්න බැරි වුණා: {e}")
    st.stop()

# අලුත් file එකක් නම් පරණ result එක අත්හරිනවා (memory එකේ රැඳෙන්නේ නෑ)
if st.session_state.get("_last_digest") != digest:
    clear_last_result()
    st.session_state["_last_digest"] = digest
    gc.collect()

if size_mb > BIG_FILE_MB:
    st.warning(
        f"File එක {size_mb:.1f} MB — ලොකුයි. Process කරද්දී RAM ගොඩක් යනවා, "
        "එකපාරට sheets ටිකක් විතරක් mark කරන එක safe."
    )

try:
    sheet_names = read_sheet_names(digest, file_bytes)
except MemoryError:
    st.error("Memory මදි වුණා — file එක ලොකු වැඩියි. Sheets බෙදලා try කරන්න.")
    st.stop()
except Exception as e:
    st.error(f"File එක කියවන්න බැරි වුණා: {e}")
    st.stop()

if not sheet_names:
    st.error("මේ workbook එකේ sheets නෑ.")
    st.stop()


def skey(i):
    """Checkbox key එක file එකට bind — අලුත් file එකකට පරණ selection ඇදෙන්නේ නෑ."""
    return f"sheet_{digest[:8]}_{i}"


# ---- Step 2: Sheets ----
with box():
    h1, h2 = st.columns([3, 1.4])
    with h1:
        step(2, "Select sheets", f"{len(sheet_names)} sheets")
    with h2:
        b1, b2 = st.columns(2)
        if b1.button("All", use_container_width=True):
            for i in range(len(sheet_names)):
                st.session_state[skey(i)] = True
        if b2.button("Clear", use_container_width=True):
            for i in range(len(sheet_names)):
                st.session_state[skey(i)] = False

    selected = []
    cols = st.columns(min(3, len(sheet_names)) or 1)
    for idx, name in enumerate(sheet_names):
        with cols[idx % len(cols)]:
            if st.checkbox(name, key=skey(idx)):
                selected.append(name)

# ---- Step 3: Settings ----
with box():
    step(3, "Settings")
    sc1, sc2 = st.columns([1, 1.6])
    with sc1:
        mult_choice = st.radio("Multiply QTY by", options=["None", 100, 1000],
                               horizontal=True, index=1)
        multiplier = None if mult_choice == "None" else mult_choice
    with sc2:
        include_unmarked = st.checkbox(
            "Unmarked sheets ද include කරන්න (clean-only)", value=False
        )

run = st.button("Process & generate", type="primary",
                use_container_width=True, disabled=not selected)
if not selected:
    st.caption("අඩුම තරමේ එක sheet එකක්වත් mark කරන්න.")

# ---- Run (error එකකින් app එක crash වෙන්නේ නෑ) ----
def _set_busy(on, file_name=""):
    """Online list එකට 🟠 එකයි, queue එකට visibility එකයි."""
    try:
        jobs = _job_registry()
        with jobs["lock"]:
            if on:
                jobs["active"][my_sid()] = {"name": my_name(), "file": file_name,
                                            "started": time.time()}
            else:
                jobs["active"].pop(my_sid(), None)
    except Exception:
        pass


if run and selected:
    clear_last_result()
    gc.collect()
    jobs = _job_registry()
    slot = jobs["sem"].acquire(blocking=False)
    if not slot:
        with jobs["lock"]:
            others = ", ".join(v["name"] for v in jobs["active"].values()) or "වෙන user කෙනෙක්"
        with st.spinner(f"Queue එකේ — දැන් process කරන්නේ: {others}"):
            slot = jobs["sem"].acquire(timeout=QUEUE_WAIT_SEC)

    if not slot:
        st.error("Server එක busy — ටික වෙලාවකින් ආපහු try කරන්න.")
    else:
        _set_busy(True, uploaded.name)
        started = time.time()
        try:
            with st.spinner("Processing..."):
                result = run_processing(
                    digest, file_bytes, tuple(sorted(selected)), multiplier, include_unmarked
                )
            lines = count_lines(digest, file_bytes, tuple(sorted(selected)))
            summary_rows = result[0]
            scaled_by_sheet = {r["Sheet"]: r.get("Scaled", 0) for r in summary_rows}
            warn_by_sheet = {r["Sheet"]: r.get("⚠", 0) for r in summary_rows}
            record = {
                "run_id": uuid.uuid4().hex[:10],
                "time": datetime.now().isoformat(timespec="seconds"),
                "user": my_name(),
                "session": my_sid()[:8],
                "file": uploaded.name,
                "file_mb": round(size_mb, 2),
                "multiplier": multiplier,
                "include_unmarked": bool(include_unmarked),
                "sheets": [
                    {"name": nm, "lines": lines.get(nm, 0),
                     "scaled_cells": scaled_by_sheet.get(nm, 0),
                     "warnings": warn_by_sheet.get(nm, 0)}
                    for nm in sorted(selected)
                ],
                "total_lines": sum(lines.get(nm, 0) for nm in selected),
                "total_scaled": result[3],
                "total_warnings": len(result[1]),
                "output_sheets": len(result[4]),
                "duration_sec": round(time.time() - started, 2),
            }
            append_run(record)
            st.session_state["_last_result"] = result
            st.session_state["_last_meta"] = {"name": uploaded.name,
                                              "multiplier": multiplier,
                                              "record": record}
        except MemoryError:
            st.session_state.pop("_last_result", None)
            st.error(
                "Memory මදි වුණා — sheets ගණන අඩු කරලා නැත්නම් file එක කොටස් වලට කඩලා try කරන්න."
            )
        except Exception as e:
            st.session_state.pop("_last_result", None)
            st.error(f"Process කරද්දී error එකක්: {type(e).__name__} — {e}")
        finally:
            _set_busy(False)
            try:
                jobs["sem"].release()
            except Exception:
                pass
            gc.collect()

# ---- Results (session_state එකේ තියෙනවා → download click කරාම නැති වෙන්නේ නෑ) ----
if st.session_state.get("_last_result"):
    try:
        summary, all_warnings, out_path, total_scaled, out_order = st.session_state["_last_result"]
        meta = st.session_state.get("_last_meta", {})
        res_mult = meta.get("multiplier")

        with box():
            step(4, "Result", f"{len(out_order)} sheets")

            rec = meta.get("record", {})
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Lines processed", f"{rec.get('total_lines', 0):,}")
            m2.metric("Output sheets", len(out_order))
            m3.metric("Cells scaled" + (f" ×{res_mult}" if res_mult else ""), total_scaled)
            m4.metric("Warnings", len(all_warnings))

            if rec.get("sheets"):
                st.caption(
                    "Lines per sheet — "
                    + " · ".join(f"{s['name']}: {s['lines']:,}" for s in rec["sheets"])
                    + f"  ·  {rec.get('duration_sec', 0)}s"
                )

            safe_table(summary, {
                "Sheet": st.column_config.TextColumn(width="medium"),
                "Qty column": st.column_config.TextColumn(width="medium"),
                "Scaled": st.column_config.NumberColumn(width="small"),
                "⚠": st.column_config.NumberColumn(width="small"),
            } if hasattr(st, "column_config") else None)

            base = meta.get("name", uploaded.name).rsplit(".", 1)[0]
            try:
                if os.path.exists(out_path):
                    with open(out_path, "rb") as fh:
                        st.download_button(
                            "Download cleaned Excel",
                            data=fh,
                            file_name=f"{base}_cleaned.xlsx",
                            mime=("application/vnd.openxmlformats-officedocument"
                                  ".spreadsheetml.sheet"),
                            type="primary",
                        )
                else:
                    st.warning("Output file එක තව නෑ (app එක restart වෙලා). ආපහු Process කරන්න.")
            except Exception:
                st.warning("Output file එක කියවන්න බැරි වුණා. ආපහු Process කරන්න.")

            with st.expander(f"Output sheet order · {len(out_order)} sheets"):
                st.write("  →  ".join(out_order))

            if all_warnings:
                with st.expander(f"Warnings ({len(all_warnings)})", expanded=False):
                    st.caption("Scaled value එක decimal, නැත්නම් digit 3ට වඩා වැඩි නැහැ.")
                    safe_table([{
                        "Sheet": w["sheet"], "Cell": w["cell"], "Header": w["header"],
                        "Original": w["original"], f"×{res_mult}": w["scaled"],
                        "Issue": w["issue"],
                    } for w in all_warnings])
    except Exception as e:
        st.session_state.pop("_last_result", None)
        st.error(f"Result එක පෙන්නද්දී error එකක්: {type(e).__name__} — {e}")

# ---- Step 5: Process log (JSON) ----
with box():
    step(5, "Process log", os.path.basename(LOG_PATH))
    log_data = load_log()
    tot = log_data.get("totals", {})

    l1, l2, l3, l4 = st.columns(4)
    l1.metric("Total runs", f"{tot.get('runs', 0):,}")
    l2.metric("Total lines", f"{tot.get('lines', 0):,}")
    l3.metric("Sheets done", f"{tot.get('sheets', 0):,}")
    l4.metric("Users", f"{tot.get('users', 0):,}")

    runs = log_data.get("runs", [])
    if runs:
        safe_table([{
            "Time": r.get("time", "").replace("T", " "),
            "User": r.get("user", ""),
            "File": r.get("file", ""),
            "Sheets": len(r.get("sheets", [])),
            "Lines": r.get("total_lines", 0),
            "Scaled": r.get("total_scaled", 0),
            "⚠": r.get("total_warnings", 0),
            "Sec": r.get("duration_sec", 0),
        } for r in runs[:15]])

        d1, d2 = st.columns([3, 1])
        with d1:
            st.download_button(
                "Download process_log.json",
                data=json.dumps(log_data, ensure_ascii=False, indent=2).encode("utf-8"),
                file_name="process_log.json",
                mime="application/json",
            )
        with d2:
            if st.button("Reset log", use_container_width=True):
                if st.session_state.get("_confirm_reset"):
                    save_log(_empty_log())
                    st.session_state["_confirm_reset"] = False
                    notify("Log එක reset කළා")
                    st.rerun()
                else:
                    st.session_state["_confirm_reset"] = True
                    st.warning("ආපහු click කරොත් log එක මකෙනවා.")
    else:
        st.caption("තාම runs නෑ — process කරාම මෙතන එනවා.")

st.markdown('<div class="foot">Multi-user ready · disk-backed output · JSON process log</div>',
            unsafe_allow_html=True)
