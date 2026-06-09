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

from core import process_sheet, build_workbook

st.set_page_config(page_title="Excel Cleaner & Qty Scaler", page_icon="📊", layout="wide")

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
    """සම්පූර්ණ processing — bytes+params එකම නම් cache එකෙන් instant (multi-user share)."""
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    selected_set = set(selected)

    processed = {}      # ordered: original sheet order
    summary = []
    warnings = []

    for name in wb.sheetnames:
        is_marked = name in selected_set
        if not is_marked and not include_unmarked:
            continue
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        grid, header_idx, target_cols, w, scaled, found = process_sheet(
            rows, multiplier, scale=is_marked
        )
        processed[name] = grid
        for x in w:
            x["sheet"] = name
        warnings.extend(w)

        if not is_marked:
            qty_status = "— (clean-only)"
        elif not found:
            qty_status = "⚠️ Qty header හමු වුණේ නැහැ"
        else:
            qty_status = "✅ " + ", ".join(target_cols.values())

        summary.append({
            "Sheet": name,
            "Mode": "Scaled" if is_marked else "Clean-only",
            "Rows": len(rows),
            "Qty column": qty_status,
            "Scaled cells": scaled,
            "Warnings": len(w),
        })

    wb.close()
    out_bytes = build_workbook(processed)
    total_scaled = sum(r["Scaled cells"] for r in summary)
    return summary, warnings, out_bytes, total_scaled


# ----------------------------- Sidebar -----------------------------

with st.sidebar:
    st.header("ℹ️ Status")
    online_badge()
    st.divider()
    st.caption(
        "හැම user session එකකම වෙන වෙනම state — එකම වෙලාවට කිහිප දෙනෙක්ට පාවිච්චි කරන්න පුළුවන්.\n\n"
        "එකම file + settings නැවත process කරාම cache එකෙන් instant."
    )


# ----------------------------- Main -----------------------------

st.title("📊 Excel Cleaner & Quantity Scaler")
st.caption("Excel upload → sheets mark → clean + QUANTITY/Actual Qty scale → download.")

uploaded = st.file_uploader("Excel file එක upload කරන්න (.xlsx / .xlsm)", type=["xlsx", "xlsm"])

if uploaded is None:
    st.info("පටන් ගන්න Excel file එකක් upload කරන්න.")
    st.stop()

file_bytes = uploaded.getvalue()

try:
    sheet_names = read_sheet_names(file_bytes)
except Exception as e:
    st.error(f"File එක කියවන්න බැරි වුණා: {e}")
    st.stop()

st.subheader("1️⃣ Sheets mark කරන්න")
mc1, mc2 = st.columns([3, 1])
with mc2:
    if st.button("✅ Select all"):
        for i in range(len(sheet_names)):
            st.session_state[f"sheet_{i}"] = True
    if st.button("✖️ Clear all"):
        for i in range(len(sheet_names)):
            st.session_state[f"sheet_{i}"] = False

selected = []
cols = st.columns(min(3, len(sheet_names)) or 1)
for idx, name in enumerate(sheet_names):
    with cols[idx % len(cols)]:
        if st.checkbox(name, key=f"sheet_{idx}"):
            selected.append(name)

st.subheader("2️⃣ Settings")
sc1, sc2 = st.columns(2)
with sc1:
    multiplier = st.radio(
        "QUANTITY / Actual Qty value එක මෙයින් වැඩි කරනවා:",
        options=[100, 1000], horizontal=True,
    )
with sc2:
    include_unmarked = st.checkbox(
        "Unmarked sheets ද output එකට include කරන්න (clean-only, scale කරන්නේ නැහැ)",
        value=False,
    )

st.divider()
run = st.button("▶️ Process කරන්න", type="primary", disabled=not selected)

if not selected:
    st.warning("අඩුම තරමේ එක sheet එකක්වත් mark කරන්න.")

if run and selected:
    with st.spinner("Processing..."):
        summary, all_warnings, out_bytes, total_scaled = run_processing(
            file_bytes, tuple(sorted(selected)), multiplier, include_unmarked
        )

    st.subheader("3️⃣ Summary")
    st.dataframe(summary, use_container_width=True, hide_index=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("Sheets in output", len(summary))
    c2.metric(f"Cells scaled (×{multiplier})", total_scaled)
    c3.metric("Warnings", len(all_warnings))

    if all_warnings:
        st.subheader("⚠️ Warnings — decimal / digit issues")
        st.caption(
            "මේ value scale කරාට පස්සේ decimal විදිහට, නැත්නම් digit 3ට වඩා වැඩි නැහැ. "
            "Multiplier (100/1000) හරිද බලන්න."
        )
        st.dataframe(
            [{
                "Sheet": w["sheet"], "Cell": w["cell"], "Header": w["header"],
                "Original": w["original"], f"Scaled (×{multiplier})": w["scaled"],
                "Issue": w["issue"],
            } for w in all_warnings],
            use_container_width=True, hide_index=True,
        )
    else:
        st.success("Warnings නැහැ — හැම scaled value එකක්ම integer + digit 3ට වඩා වැඩියි. 👍")

    st.subheader("4️⃣ Download")
    base = uploaded.name.rsplit(".", 1)[0]
    st.download_button(
        "⬇️ Clean Excel download කරන්න",
        data=out_bytes,
        file_name=f"{base}_cleaned.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("Output එකේ හැම cell එකක්ම TEXT format ('@'). Marked sheets scale වෙනවා, unmarked clean-only.")
