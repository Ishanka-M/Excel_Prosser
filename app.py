"""
Excel Cleaner & Quantity Scaler — single-file Streamlit app (self-contained).
Run:  python -m streamlit run app.py
"""
import io
import re
import time
import uuid
import threading
from copy import copy
from datetime import datetime, date, time as _time

import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

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


# ===================== STREAMLIT APP =====================

st.set_page_config(page_title="Excel Cleaner & Qty Scaler", page_icon="📊", layout="centered")

CSS = """
<style>
.block-container {max-width: 920px; padding-top: 1.4rem; padding-bottom: 3rem;}
#MainMenu, footer {visibility: hidden;}
.app-header{
  background:linear-gradient(135deg,#4f46e5 0%,#0ea5e9 100%);
  padding:20px 24px;border-radius:16px;margin-bottom:18px;
  box-shadow:0 6px 24px rgba(79,70,229,.25);
}
.app-header h1{color:#fff;font-size:1.45rem;font-weight:700;margin:0;letter-spacing:.2px;}
.app-header p{color:#e0e7ff;font-size:.86rem;margin:.3rem 0 0;}
.step-label{font-size:.78rem;font-weight:700;letter-spacing:.6px;
  text-transform:uppercase;color:#6366f1;margin:.2rem 0 .5rem;}
[data-testid="stMetric"]{
  background:rgba(125,125,125,.07);border:1px solid rgba(125,125,125,.16);
  border-radius:14px;padding:12px 16px;
}
[data-testid="stMetricValue"]{font-size:1.5rem;font-weight:700;}
[data-testid="stMetricLabel"]{opacity:.7;font-size:.78rem;}
.stButton>button,.stDownloadButton>button{border-radius:10px;font-weight:600;}
.stDownloadButton>button{width:100%;padding:.55rem;}
section[data-testid="stSidebar"] [data-testid="stMetric"]{text-align:center;}
hr{margin:1rem 0;}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

PRESENCE_WINDOW = 20  # තත්පර — මේ ඇතුළත heartbeat ආපු session = online


# ----------------------- Online presence (shared across users) -----------------------

@st.cache_resource
def _presence_registry():
    """හැම user session එකකම පොදු (shared) registry එක — process එකට එකයි."""
    return {"sessions": {}, "lock": threading.Lock()}


def heartbeat_and_count(window=PRESENCE_WINDOW):
    reg = _presence_registry()
    sid = st.session_state.setdefault("_sid", str(uuid.uuid4()))
    now = time.time()
    with reg["lock"]:
        reg["sessions"][sid] = now
        stale = [k for k, v in reg["sessions"].items() if now - v > window]
        for k in stale:
            reg["sessions"].pop(k, None)
        return len(reg["sessions"])


@st.fragment(run_every=5)
def online_badge():
    """තත්පර 5කට වරක් rerun වෙලා heartbeat update + count පෙන්නනවා (app එක rerun නොකර)."""
    n = heartbeat_and_count()
    st.metric("🟢 Online users", n)


# ----------------------- Cached heavy work (speed + shared across users) -----------------------

@st.cache_data(show_spinner=False, max_entries=30)
def read_sheet_names(file_bytes: bytes):
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    names = list(wb.sheetnames)
    wb.close()
    return names


@st.cache_data(show_spinner=False, max_entries=30)
def run_processing(file_bytes: bytes, selected: tuple, multiplier: int, include_unmarked: bool):
    """Cached wrapper — එකම file+settings නැවත දාම instant (multi-user share).

    Marked sheet එකකට output එකේ:
      - "System <name>"   : faithful TEXT + QUANTITY/Actual Qty scale, original position
      - "Physical <name>" : original sheet එක verbatim (value+format+style), workbook අන්තිමට
    """
    return build_output(file_bytes, selected, multiplier, include_unmarked)


# ----------------------------- Sidebar -----------------------------

with st.sidebar:
    st.markdown("### ⚙️ Status")
    online_badge()
    st.divider()
    st.markdown("**How it works**")
    st.caption(
        "1. Excel upload කරන්න\n\n"
        "2. Process කරන sheets mark කරන්න\n\n"
        "3. Multiplier (×100 / ×1000) තෝරන්න\n\n"
        "4. Clean Excel download කරන්න"
    )
    st.divider()
    st.caption(
        "Marked sheet එකකට **System** (scaled) + **Physical** (original) "
        "sheet දෙකක් හැදෙනවා.\n\nMulti-user ready · cached for speed."
    )


# ----------------------------- Main -----------------------------

st.markdown(
    '<div class="app-header"><h1>📊 Excel Cleaner &amp; Quantity Scaler</h1>'
    '<p>Upload → mark sheets → clean &amp; scale QUANTITY / Actual Qty → download</p></div>',
    unsafe_allow_html=True,
)

# ---- Step 1: Upload ----
with st.container(border=True):
    st.markdown('<div class="step-label">Step 1 · Upload</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Excel file (.xlsx / .xlsm)", type=["xlsx", "xlsm"], label_visibility="collapsed"
    )

if uploaded is None:
    st.info("පටන් ගන්න Excel file එකක් upload කරන්න.")
    st.stop()

file_bytes = uploaded.getvalue()
try:
    sheet_names = read_sheet_names(file_bytes)
except Exception as e:
    st.error(f"File එක කියවන්න බැරි වුණා: {e}")
    st.stop()

# ---- Step 2: Sheets ----
with st.container(border=True):
    h1, h2 = st.columns([3, 1.4])
    with h1:
        st.markdown('<div class="step-label">Step 2 · Select sheets</div>', unsafe_allow_html=True)
    with h2:
        b1, b2 = st.columns(2)
        if b1.button("All", use_container_width=True):
            for i in range(len(sheet_names)):
                st.session_state[f"sheet_{i}"] = True
        if b2.button("Clear", use_container_width=True):
            for i in range(len(sheet_names)):
                st.session_state[f"sheet_{i}"] = False

    selected = []
    cols = st.columns(min(3, len(sheet_names)) or 1)
    for idx, name in enumerate(sheet_names):
        with cols[idx % len(cols)]:
            if st.checkbox(name, key=f"sheet_{idx}"):
                selected.append(name)

# ---- Step 3: Settings ----
with st.container(border=True):
    st.markdown('<div class="step-label">Step 3 · Settings</div>', unsafe_allow_html=True)
    sc1, sc2 = st.columns([1, 1.6])
    with sc1:
        mult_choice = st.radio("Multiply QTY by", options=["None", 100, 1000],
                               horizontal=True, index=1)
        multiplier = None if mult_choice == "None" else mult_choice
    with sc2:
        include_unmarked = st.checkbox(
            "Unmarked sheets ද include කරන්න (clean-only)", value=False
        )

run = st.button("▶️  Process & generate", type="primary", use_container_width=True, disabled=not selected)
if not selected:
    st.caption("අඩුම තරමේ එක sheet එකක්වත් mark කරන්න.")

# ---- Results ----
if run and selected:
    with st.spinner("Processing..."):
        summary, all_warnings, out_bytes, total_scaled, out_order = run_processing(
            file_bytes, tuple(sorted(selected)), multiplier, include_unmarked
        )

    with st.container(border=True):
        st.markdown('<div class="step-label">Result</div>', unsafe_allow_html=True)

        m1, m2, m3 = st.columns(3)
        m1.metric("Output sheets", len(out_order))
        m2.metric("Cells scaled" + (f" ×{multiplier}" if multiplier else " (none)"), total_scaled)
        m3.metric("Warnings", len(all_warnings))

        st.dataframe(
            summary, use_container_width=True, hide_index=True,
            column_config={
                "Sheet": st.column_config.TextColumn(width="medium"),
                "Qty column": st.column_config.TextColumn(width="medium"),
                "Scaled": st.column_config.NumberColumn(width="small"),
                "⚠": st.column_config.NumberColumn(width="small"),
            },
        )

        base = uploaded.name.rsplit(".", 1)[0]
        st.download_button(
            "⬇️  Download cleaned Excel",
            data=out_bytes,
            file_name=f"{base}_cleaned.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        with st.expander(f"Output sheet order  ·  {len(out_order)} sheets"):
            st.write("  →  ".join(out_order))

        if all_warnings:
            with st.expander(f"⚠️  Warnings ({len(all_warnings)})", expanded=False):
                st.caption("Scaled value එක decimal, නැත්නම් digit 3ට වඩා වැඩි නැහැ.")
                st.dataframe(
                    [{
                        "Sheet": w["sheet"], "Cell": w["cell"], "Header": w["header"],
                        "Original": w["original"], f"×{multiplier}": w["scaled"],
                        "Issue": w["issue"],
                    } for w in all_warnings],
                    use_container_width=True, hide_index=True,
                )
