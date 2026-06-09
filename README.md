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

## 🧾 System / Physical sheet split (update)

Marked sheet එකකට output එකේ **sheet දෙකක්** හැදෙනවා:

| Output sheet | අන්තර්ගතය | තැන |
|--------------|-----------|-----|
| `System <name>` | process + scale (×100/1000) + clean + TEXT | original sheet තිබුණ **position** එකේම |
| `Physical <name>` | original sheet එක **එහෙම්ම** (as-is, scale/clean නැතුව) | workbook එකේ **අන්තිමට** |

Unmarked sheets (include කරොත්) clean-only, ඒවාගේ original නම් + position එකේම.

## 🔄 Multi-user / Speed / Online count (update)

- **Multi-user:** Streamlit session එකක් = එක user — state වෙන වෙනම. එකම වෙලාවට කිහිප දෙනෙක්ට පාවිච්චි කරන්න පුළුවන්.
- **Speed:** sheet-name reading + full processing `st.cache_data` වලින් cache — එකම file+settings නැවත දාම instant (users අතරෙත් share වෙනවා).
- **Online users:** sidebar එකේ live count (`st.fragment` heartbeat, ~5s refresh, 20s window). `streamlit>=1.37` ඕනේ.
- **Select all / Clear all** + **Unmarked sheets include (clean-only)** toggle එකත් add කරලා.
