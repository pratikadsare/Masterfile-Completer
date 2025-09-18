# Masterfile Filler (Streamlit)

**One‑click browser tool** to map your **raw CSV/XLSX** into a **marketplace masterfile** using a simple **mapping**.  
Upload **Raw Data → Template → Mapping**, then click **Process & Download** to get the filled XLSX.

---

## 🔧 Features
- 3 tabs: **Upload Raw Data**, **Upload Template (Masterfile)**, **Process & Download**
- Header‑based column mapping **and** single‑cell assignment (e.g., put marketplace name into `B2`)
- Type coercion (`str`, `int`, `float`, `bool`, `date`) and defaults
- Multi‑sheet support (set `template_sheet` per line in mapping)
- Preserves your workbook; writes values in place

---

## 🚀 Deploy from GitHub (Streamlit Community Cloud)
1. Push these files to a new GitHub repo (public or private).
2. Go to **Streamlit Community Cloud** and click **New app**.
3. Pick your repo, branch, and set **Main file path** to `streamlit_app.py`.
4. Click **Deploy** → you get a public URL to share.
5. (Optional) Add `mapping_template.xlsx` to the repo so your team can download a starter mapping.

> Tip: Any time you push to GitHub, Streamlit auto‑rebuilds the app.

---

## 🖱️ How to use
1. **Tab 1:** Upload RAW (CSV/XLSX); if Excel, choose a sheet.
2. **Tab 2:** Upload TEMPLATE (XLSX) and choose the sheet to fill.
3. **Tab 3:** Upload **MAPPING** (XLSX/CSV) and click **⚙️ Process & Download**.

**Mapping columns:**
- `template_sheet` (required): Sheet name inside your template (`Masterfile` recommended)
- `template_header` (row‑wise): Column header text to fill
- `template_cell` (one‑off): Cell address like `B2`
- `raw_column`: Copy from this column in RAW
- `default_value`: Fallback if raw is missing/blank
- `dtype`: `str` | `int` | `float` | `bool` | `date`
- `required`: `true`/`false` → warns if RAW column is missing

A starter mapping is included: **`mapping_template.xlsx`**.

---

## 🧪 Try it locally
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```
Then open the local URL in your browser.

---

## 📁 Sample files
- `sample_data/sample_raw.csv`
- `sample_data/sample_template.xlsx`
- `sample_data/sample_mapping.xlsx`

Use these to verify the flow before using your real files.

---

## 📜 License
MIT License © 2025
