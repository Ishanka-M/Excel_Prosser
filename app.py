"""
Excel Cleaner & Quantity Scaler  —  Streamlit UI
------------------------------------------------
Multi-user ready + cached (speed) + live online-user count.

- Excel (.xlsx / .xlsm) upload -> sheet names view + mark
- Mark කරපු sheet වල: value paste, clean (space/hidden symbols remove),
  හැම cell එකක්ම TEXT, QUANTITY/Actual Qty x100/x1000 in-place replace.
- Scaled value decimal නම් / digit 3ට වඩා වැඩි නැත්නම් -> notify.
- Unmarked sheets optionally include (clean-only).
- Summary + clean Excel download.
"""

import io
import time
import uuid
import threading

import streamlit as st
from openpyxl import load_workbook

from core import build_output

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
        multiplier = st.radio("Multiply QTY by", options=[100, 1000], horizontal=True)
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
        m2.metric(f"Cells scaled ×{multiplier}", total_scaled)
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
