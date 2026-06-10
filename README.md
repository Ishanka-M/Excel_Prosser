# 📊 Excel Cleaner & Quantity Scaler

Streamlit tool එකක්: Excel upload කරලා → sheets mark කරලා → clean කරලා →
**QUANTITY / Actual Qty** columns `× 100` හෝ `× 1000` scale කරලා → clean Excel download.

## ✨ Features

- `.xlsx` / `.xlsm` upload කරලා **sheet names view + checkbox වලින් mark**
- Mark කරපු sheet වල:
  - **Value paste** — formula cells වල cached value එක ගන්නවා
  - **Clean** — space, non-breaking space (`\xa0`), zero-width සහ control/hidden symbols remove
  - **හැම cell value එකක්ම TEXT format** (`@`) එකට convert
  - **QUANTITY / Actual Qty** column auto-detect (header row scan), `× 100` හෝ `× 1000`, **in-place replace**, scaled value එකත් text
- **Validation** (final scaled value එකට): decimal නම් හෝ digit 3ට වඩා වැඩි නැත්නම් → **notify**
- **Summary** එකක් + clean Excel **download**

## 🚀 Run (local)

```bash
pip install -r requirements.txt
streamlit run app.py
```

> **Windows + Microsoft Store Python note:** `streamlit` command එක හම්බ වෙන්නේ නැත්නම්
> `python -m streamlit run app.py` use කරන්න.

## ☁️ GitHub + Streamlit Community Cloud deploy

1. මේ files (`app.py`, `core.py`, `requirements.txt`) GitHub repo එකකට push කරන්න.
2. https://share.streamlit.io → **New app** → repo + branch + `app.py` තෝරන්න → Deploy.

## 🗂️ Files

| File | Description |
|------|-------------|
| `app.py` | Streamlit UI layer |
| `core.py` | Cleaning + scaling logic (UI-independent, testable) |
| `requirements.txt` | Dependencies |

## ⚙️ Notes

- Header row එක මුල් **rows 30** ඇතුළේ `QUANTITY` හෝ `Actual Qty` තියෙන row එකෙන් auto-detect.
- Output sheet එකේ **mark කරපු sheets පමණයි** (unmarked sheets include වෙන්නේ නැහැ).
- හැම output cell එකක්ම TEXT (`@`) format — Excel එකේ අයෙත් number/date විදිහට auto-convert වෙන්නේ නැහැ.

## 🧾 System / Physical sheet split

Marked sheet එකකට output එකේ **sheet දෙකක්** හැදෙනවා:

| Output sheet | අන්තර්ගතය | තැන |
|--------------|-----------|-----|
| `System <name>` | හැම cell එකක්ම **TEXT** (`@`) — displayed value එකම (dates number_format respect, leading zeros, decimals, spaces එහෙම්ම); QUANTITY/Actual Qty `×100/1000` | original position |
| `Physical <name>` | original sheet එක **verbatim** (value + format + style) | workbook අන්තිමට |

### Value/format faithful (text conversion)

හැම cell එකක්ම text වුණත් **displayed value එක වෙනස් වෙන්නේ නැහැ**:

- **Date:** `4/20/2026` (m/d/yyyy) → text `4/20/2026` (ISO `2026-04-20T...` වෙන්නේ නැහැ). number_format token (m/d/yyyy, dd-mmm-yyyy, h:mm AM/PM ...) respect කරනවා.
- Leading zeros (`0003983473`), `0.000` decimals (`0.092`), internal spaces — එහෙම්ම
- Hidden / zero-width / control chars විතරක් අයින් (visible value නොවෙනස්ව)
- QUANTITY / Actual Qty විතරක් `× multiplier`; අනික් data වලට අලුතෙන් මොකුත් add වෙන්නේ නැහැ


## 🔄 Multi-user / Speed / Online count (update)

- **Multi-user:** Streamlit session එකක් = එක user — state වෙන වෙනම. එකම වෙලාවට කිහිප දෙනෙක්ට පාවිච්චි කරන්න පුළුවන්.
- **Speed:** sheet-name reading + full processing `st.cache_data` වලින් cache — එකම file+settings නැවත දාම instant (users අතරෙත් share වෙනවා).
- **Online users:** sidebar එකේ live count (`st.fragment` heartbeat, ~5s refresh, 20s window). `streamlit>=1.37` ඕනේ.
- **Select all / Clear all** + **Unmarked sheets include (clean-only)** toggle එකත් add කරලා.
