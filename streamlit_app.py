
import io
from io import BytesIO
import pandas as pd
import numpy as np
import streamlit as st
from typing import List, Dict, Tuple, Optional
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Masterfile Filler", layout="wide")

# --- Custom header (as provided) ---
st.markdown("<h1 style='text-align: center;'>🧩 Masterfile Filler</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; font-style: italic;'>Innovating with AI Today ⏐ Leading Automation Tomorrow</h4>", unsafe_allow_html=True)
st.caption("Upload your raw data, the masterfile (template), and a simple 2-column mapping, then click **Process & Download**.")

# ---------- Helpers ----------
def get_excel_sheets(uploaded_file) -> List[str]:
    if uploaded_file is None:
        return []
    try:
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
        return xls.sheet_names
    except Exception:
        try:
            xls = pd.ExcelFile(uploaded_file)  # fallback (xls)
            return xls.sheet_names
        except Exception:
            return []

def read_tabular(uploaded_file, sheet: Optional[str] = None) -> pd.DataFrame:
    """Read CSV/XLSX/XLS into DataFrame. If Excel and sheet is None, use first sheet."""
    if uploaded_file is None:
        return pd.DataFrame()
    name = uploaded_file.name.lower()
    try:
        if name.endswith((".csv", ".txt")):
            return pd.read_csv(uploaded_file)
        # Excel path
        if sheet is not None:
            if name.endswith(".xlsx"):
                return pd.read_excel(uploaded_file, sheet_name=sheet, engine="openpyxl")
            return pd.read_excel(uploaded_file, sheet_name=sheet)
        # default first sheet
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl") if name.endswith(".xlsx") else pd.ExcelFile(uploaded_file)
        first = xls.sheet_names[0]
        return pd.read_excel(xls, sheet_name=first, engine="openpyxl" if name.endswith(".xlsx") else None)
    except Exception as e:
        st.error(f"Failed to read table: {e}")
        return pd.DataFrame()

def find_mapping_columns(df_cols: List[str]) -> Tuple[Optional[str], Optional[str]]:
    """Return (raw_header_col, template_header_col) based on flexible name matches."""
    norm = {c.lower().strip(): c for c in df_cols}
    raw_candidates = [
        "header of row sheet", "header of raw sheet", "raw header", "raw", "source", "source header", "from"
    ]
    tpl_candidates = [
        "header of masterfile template", "template header", "masterfile header", "template", "target", "to"
    ]
    raw_col = None
    tpl_col = None
    for k in raw_candidates:
        if k in norm:
            raw_col = norm[k]; break
    for k in tpl_candidates:
        if k in norm:
            tpl_col = norm[k]; break
    return raw_col, tpl_col

def read_two_col_mapping(uploaded_file) -> pd.DataFrame:
    """Read a 2-column mapping and normalize to columns: raw_header, template_header."""
    if uploaded_file is None:
        return pd.DataFrame()
    name = uploaded_file.name.lower()
    try:
        if name.endswith((".csv", ".txt")):
            m = pd.read_csv(uploaded_file)
        else:
            m = pd.read_excel(uploaded_file, engine="openpyxl" if name.endswith(".xlsx") else None)
    except Exception as e:
        st.error(f"Failed to read mapping: {e}")
        return pd.DataFrame()

    if m.empty:
        return m
    # Identify the two columns by flexible names
    raw_col_name, tpl_col_name = find_mapping_columns(list(m.columns))
    if raw_col_name is None or tpl_col_name is None:
        st.error("Mapping must have two headers: **'header of row sheet'** (or 'header of raw sheet') and **'header of masterfile template'**.")
        return pd.DataFrame()

    m = m[[raw_col_name, tpl_col_name]].copy()
    m.columns = ["raw_header", "template_header"]
    # Clean rows
    m["raw_header"] = m["raw_header"].astype(str).str.strip()
    m["template_header"] = m["template_header"].astype(str).str.strip()
    m = m[(m["raw_header"] != "") & (m["template_header"] != "")]
    # Drop duplicates, keep first
    m = m.drop_duplicates(subset=["template_header"], keep="first").reset_index(drop=True)
    return m

def detect_header_row(ws, needed_headers: List[str], search_rows: int = 20) -> Tuple[int, Dict[str, int]]:
    """Find the row index (1-based) that contains the template headers. Return row and mapping of header->col index."""
    nh = [h.strip().lower() for h in needed_headers if h and str(h).strip()]
    best_row = None
    best_match_count = -1
    header_map = {}
    max_row = min(ws.max_row or 1, search_rows)
    for r in range(1, max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        val_to_idx = {}
        for idx, v in enumerate(row_vals, start=1):
            if v is None:
                continue
            key = str(v).strip().lower()
            if key and key not in val_to_idx:
                val_to_idx[key] = idx
        match_count = sum(1 for h in nh if h in val_to_idx)
        if match_count > best_match_count:
            best_match_count = match_count
            best_row = r
            header_map = {h: val_to_idx[h] for h in nh if h in val_to_idx}
        if match_count == len(nh) and len(nh) > 0:
            break
    return best_row or 1, header_map

def write_table_into_template_2col(
    template_bytes: BytesIO,
    template_sheet: str,
    mapping_df: pd.DataFrame,
    raw_df: pd.DataFrame,
    auto_detect_headers: bool,
    logs: List[str]
) -> BytesIO:
    """Fill the template using a simple 2-column mapping: raw_header -> template_header."""
    wb = load_workbook(filename=template_bytes)
    if template_sheet not in wb.sheetnames:
        ws = wb.create_sheet(template_sheet)
        header_row_idx = 1
        header_map = {}
    else:
        ws = wb[template_sheet]
        needed_headers = mapping_df["template_header"].tolist()
        if auto_detect_headers and len(needed_headers) > 0:
            header_row_idx, header_map = detect_header_row(ws, needed_headers)
        else:
            header_row_idx, header_map = 1, {}

    # Ensure headers exist in the template sheet in the order of mapping
    template_headers = mapping_df["template_header"].tolist()
    current_headers = [ws.cell(row=header_row_idx, column=c).value for c in range(1, ws.max_column + 1)]
    current_norm = [str(x).strip().lower() if x is not None else "" for x in current_headers]
    used_cols = set(header_map.values())
    next_free_col = 1
    while next_free_col in used_cols:
        next_free_col += 1

    for h in template_headers:
        key = h.strip().lower()
        if key in current_norm:
            col_idx = current_norm.index(key) + 1
            header_map[key] = col_idx
        elif key in header_map:
            pass
        else:
            while next_free_col in used_cols:
                next_free_col += 1
            ws.cell(row=header_row_idx, column=next_free_col, value=h)
            header_map[key] = next_free_col
            used_cols.add(next_free_col)
            next_free_col += 1

    start_row = header_row_idx + 1

    # Build list of valid mappings (raw col exists)
    valid_maps = []
    missing_raw = []
    for _, row in mapping_df.iterrows():
        raw_col = row["raw_header"]
        tpl = row["template_header"]
        key = str(tpl).strip().lower()
        col_idx = header_map.get(key)
        if raw_col not in raw_df.columns:
            missing_raw.append(raw_col)
            continue
        if col_idx is None:
            logs.append(f"⚠️ Could not resolve column for template header '{tpl}'.")
            continue
        valid_maps.append((raw_col, col_idx, tpl))

    if missing_raw:
        uniq = sorted(set(missing_raw))
        logs.append(f"⚠️ Missing raw columns (skipped): {', '.join(uniq)}")

    # Write values row-wise
    nrows = len(raw_df)
    for i in range(nrows):
        for raw_col, col_idx, _tpl in valid_maps:
            val = raw_df.iloc[i][raw_col]
            ws.cell(row=start_row + i, column=col_idx, value=val if not pd.isna(val) else "")

    # Best-effort widths
    try:
        for key, col_idx in header_map.items():
            max_width = len(str(ws.cell(row=header_row_idx, column=col_idx).value or ""))
            for rr in range(start_row, start_row + min(nrows, 200)):
                v = ws.cell(row=rr, column=col_idx).value
                max_width = max(max_width, len(str(v)) if v is not None else 0)
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_width + 2, 60)
    except Exception:
        pass

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ---------- SESSION ----------
if "raw_df" not in st.session_state: st.session_state.raw_df = pd.DataFrame()
if "raw_sheet" not in st.session_state: st.session_state.raw_sheet = None
if "template_bytes" not in st.session_state: st.session_state.template_bytes = None
if "template_sheet" not in st.session_state: st.session_state.template_sheet = None
if "mapping_df" not in st.session_state: st.session_state.mapping_df = pd.DataFrame()

tab1, tab2, tab3 = st.tabs(["1) Upload Raw Data", "2) Upload Template (Masterfile)", "3) Process & Download"])

# ---- Tab 1: Raw Data ----
with tab1:
    st.subheader("Upload Raw Data (CSV or Excel)")
    raw_file = st.file_uploader("Choose raw data file", type=["csv","xlsx","xls"], key="raw_uploader")
    raw_sheet = None
    if raw_file is not None and raw_file.name.lower().endswith(("xlsx","xls")):
        try:
            sheets = get_excel_sheets(raw_file)
            if sheets:
                raw_sheet = st.selectbox("Select sheet", sheets, index=0)
                st.session_state.raw_sheet = raw_sheet
                df_prev = read_tabular(raw_file, sheet=raw_sheet)
            else:
                df_prev = read_tabular(raw_file)
        except Exception as e:
            st.error(f"Could not read Excel: {e}")
            df_prev = pd.DataFrame()
    else:
        df_prev = read_tabular(raw_file)

    if raw_file is not None:
        st.session_state.raw_df = df_prev
        st.success(f"Loaded {len(df_prev):,} rows × {len(df_prev.columns):,} columns.")
        st.dataframe(df_prev.head(50), use_container_width=True)

# ---- Tab 2: Template ----
with tab2:
    st.subheader("Upload Marketplace Template (Excel)")
    template_file = st.file_uploader("Choose template (masterfile) Excel", type=["xlsx","xls"], key="template_uploader")
    if template_file is not None:
        template_bytes = BytesIO(template_file.getbuffer())
        st.session_state.template_bytes = template_bytes
        try:
            sheets = get_excel_sheets(template_file)
            sheet_name = st.selectbox("Select template sheet to fill", sheets, index=0) if sheets else None
        except Exception as e:
            st.error(f"Failed to open template: {e}")
            sheet_name = None
        st.session_state.template_sheet = sheet_name
        if sheet_name:
            prev = read_tabular(template_file, sheet=sheet_name)
            st.caption("Preview (first 5 rows):")
            st.dataframe(prev.head(5), use_container_width=True)

# ---- Tab 3: Process & Download ----
with tab3:
    st.subheader("Mapping (2 columns) & Processing")
    st.markdown(
        "**Provide a 2-column mapping (CSV/XLSX)** — exactly these headers:\n"
        "1) **header of row sheet** *(or 'header of raw sheet')* — the column name in your RAW file\n"
        "2) **header of masterfile template** — the target column header in the template sheet"
    )
    mapping_file = st.file_uploader("Upload 2-column mapping (xlsx/csv)", type=["xlsx","xls","csv"], key="mapping_uploader")
    autodetect = st.checkbox("Auto‑detect header row in template", value=True)

    if mapping_file is not None:
        mapping_df = read_two_col_mapping(mapping_file)
        st.session_state.mapping_df = mapping_df
        if not mapping_df.empty:
            st.success(f"Loaded mapping with {len(mapping_df):,} lines.")
            st.dataframe(mapping_df, use_container_width=True)
        else:
            st.warning("Mapping is empty or invalid.")

    def download_two_col_template():
        df = pd.DataFrame({
            "header of row sheet": ["sku","name","description","price","qty","category"],
            "header of masterfile template": ["SKU","Title","Description","Price","Quantity","Category"],
        })
        return df

    c1, c2 = st.columns([1,2])
    with c1:
        if st.button("⬇️ Download mapping template (.xlsx)"):
            df = download_two_col_template()
            from io import BytesIO
            from openpyxl import Workbook
            out = BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as w:
                df.to_excel(w, index=False, sheet_name="mapping")
            out.seek(0)
            st.download_button("Download mapping_template.xlsx", data=out.getvalue(), file_name="mapping_template.xlsx")
    with c2:
        st.info("Tip: Only the two headers above are needed. No defaults, datatypes, or cells required.")

    st.divider()
    can_process = (
        not st.session_state.raw_df.empty
        and st.session_state.template_bytes is not None
        and st.session_state.template_sheet is not None
        and not st.session_state.mapping_df.empty
    )
    if not can_process:
        st.info("Please upload raw data, template (and select sheet), and 2-column mapping to enable processing.")

    if st.button("⚙️ Process & Download", type="primary", disabled=not can_process):
        try:
            out_bytes = write_table_into_template_2col(
                template_bytes = BytesIO(st.session_state.template_bytes.getbuffer()),
                template_sheet = st.session_state.template_sheet,
                mapping_df = st.session_state.mapping_df,
                raw_df = st.session_state.raw_df,
                auto_detect_headers = bool(autodetect),
                logs = []
            )
            st.success("Done! Your filled masterfile is ready to download.")
            st.download_button(
                label="⬇️ Download Filled Masterfile (Excel)",
                data=out_bytes.getvalue(),
                file_name="filled_masterfile.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.exception(e)
