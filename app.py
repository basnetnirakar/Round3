from __future__ import annotations

import zipfile
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st
from lxml import etree
from openpyxl import load_workbook
from io import BytesIO

# Optional: Pillow for better image resampling when resizing chromatograms
try:
    from PIL import Image
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False

# Optional dependency: streamlit-aggrid enables clickable table rows
try:
    from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
    AGGRID_AVAILABLE = True
except Exception:
    AGGRID_AVAILABLE = False

# ============================================================
# Antibody Data Manager
# ============================================================

SHEET_NAME = "Sheet1"
EXCEL_ROW_OFFSET = 2  # pandas row 0 corresponds to Excel row 2 because row 1 is headers

DEFAULT_COLUMNS = [
    "Antibody_name",
    "Human EC50 (nM)",
    "SPR_binding",
    "ka(1/Ms)",
    "kd(1/s)",
    "KD",
    "Remarks",
    "SEC profile PSR",
    "SEC value",
]

st.set_page_config(page_title="Antibody Data Manager", layout="wide")


# -----------------------------
# Utility helpers
# -----------------------------
def format_scalar(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value)


def index_to_excel_col(idx: int) -> str:
    letters = ""
    i = idx + 1
    while i > 0:
        i, rem = divmod(i - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


@st.cache_data(show_spinner=False)
def read_main_table(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=SHEET_NAME)
    df.columns = [str(c).strip() for c in df.columns]
    return df


@st.cache_data(show_spinner=False)
def read_workbook_info(file_bytes: bytes) -> Dict[str, object]:
    wb = load_workbook(BytesIO(file_bytes), data_only=False)
    ws = wb[SHEET_NAME]
    return {
        "sheet_names": wb.sheetnames,
        "max_row": ws.max_row,
        "max_column": ws.max_column,
    }


@st.cache_data(show_spinner=False)
def extract_cell_images_from_xlsx(file_bytes: bytes) -> Dict[str, bytes]:
    image_map: Dict[str, bytes] = {}
    with zipfile.ZipFile(BytesIO(file_bytes), "r") as z:
        names = set(z.namelist())
        rels_name = "xl/richData/_rels/richValueRel.xml.rels"
        sheet_xml_name = "xl/worksheets/sheet1.xml"
        if rels_name not in names or sheet_xml_name not in names:
            return image_map
        rel_root = etree.fromstring(z.read(rels_name))
        rels: Dict[str, str] = {}
        for rel in rel_root:
            rel_id = rel.get("Id")
            target = rel.get("Target")
            if rel_id and target:
                rels[rel_id] = target
        sheet_root = etree.fromstring(z.read(sheet_xml_name))
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        for cell in sheet_root.xpath(".//main:c[@vm]", namespaces=ns):
            cell_ref = cell.get("r")
            vm = cell.get("vm")
            if not cell_ref or not vm:
                continue
            rel_id = f"rId{vm}"
            target = rels.get(rel_id)
            if not target:
                continue
            media_path = "xl/" + target.replace("../", "")
            if media_path not in names:
                continue
            image_map[cell_ref] = z.read(media_path)
    return image_map


@st.cache_data(show_spinner=False)
def build_app_dataframe(file_bytes: bytes) -> pd.DataFrame:
    df = read_main_table(file_bytes).copy()
    image_map = extract_cell_images_from_xlsx(file_bytes)
    excel_rows: List[int] = []
    image_columns: List[str] = []
    for idx, c in enumerate(df.columns):
        col_letter = index_to_excel_col(idx)
        header_hit = "image" in str(c).lower() or "chromatogram" in str(c).lower()
        has_image_in_map = any(k.startswith(col_letter) for k in image_map.keys())
        if header_hit or has_image_in_map:
            image_columns.append(c)
    image_bytes_map: Dict[str, List[Optional[bytes]]] = {c: [] for c in image_columns}
    for i in range(len(df)):
        excel_row = i + EXCEL_ROW_OFFSET
        excel_rows.append(excel_row)
        for col in image_columns:
            col_idx = int(df.columns.get_loc(col))
            col_letter = index_to_excel_col(col_idx)
            cell_ref = f"{col_letter}{excel_row}"
            image_bytes_map[col].append(image_map.get(cell_ref))
    df["excel_row"] = excel_rows
    for col in image_columns:
        df[f"{col}_bytes"] = image_bytes_map[col]
        df[f"has_{col}_image"] = df[f"{col}_bytes"].apply(lambda x: x is not None)
    return df



# -----------------------------
# App UI
# -----------------------------
st.title("Antibody Data Manager")
st.caption("Interactive browser for antibody data, including in-cell EC50 and SPR images.")

uploaded_file = st.file_uploader(
    "Upload your workbook (.xlsx)",
    type=["xlsx"],
    help="Upload the Excel file to load the data.",
)

if uploaded_file is None:
    st.info("Please upload the Excel file to get started.")
    st.stop()

file_bytes = uploaded_file.read()

try:
    app_df = build_app_dataframe(file_bytes)
    workbook_info = read_workbook_info(file_bytes)
except Exception as e:
    st.error(f"Failed to load workbook: {e}")
    st.stop()


# ------------------------------------------------------------------
# Sidebar — Column selector
# ------------------------------------------------------------------
all_data_columns = [
    c for c in app_df.columns
    if not c.endswith("_bytes") and not c.startswith("has_") and c != "excel_row"
]
default_selection = [c for c in DEFAULT_COLUMNS if c in all_data_columns]

st.sidebar.header("Columns to display")
display_present = st.sidebar.multiselect(
    "Select columns",
    options=all_data_columns,
    default=default_selection,
)
if not display_present:
    st.sidebar.warning("Select at least one column.")
    display_present = all_data_columns[:1]


# ------------------------------------------------------------------
# Sidebar — Filters
# ------------------------------------------------------------------
st.sidebar.header("Filters")
search_text = st.sidebar.text_input("Search antibody or remarks")
only_with_ec50_image = st.sidebar.checkbox("Only rows with EC50 image")
only_with_spr_image = st.sidebar.checkbox("Only rows with SPR image")
only_spr_binding_positive = st.sidebar.checkbox("Only SPR_binding = 1")

filtered_df = app_df.copy()

if search_text:
    q = search_text.strip()
    mask = (
        filtered_df["Antibody_name"].astype(str).str.contains(q, case=False, na=False)
        | filtered_df["Remarks"].astype(str).str.contains(q, case=False, na=False)
    )
    filtered_df = filtered_df[mask]

if only_with_ec50_image:
    ec50_flag = next((c for c in filtered_df.columns if c.startswith("has_") and "ec50" in c.lower()), None)
    if ec50_flag:
        filtered_df = filtered_df[filtered_df[ec50_flag]]

if only_with_spr_image:
    spr_flag = next((c for c in filtered_df.columns if c.startswith("has_") and "spr" in c.lower()), None)
    if spr_flag:
        filtered_df = filtered_df[filtered_df[spr_flag]]

if only_spr_binding_positive:
    filtered_df = filtered_df[filtered_df["SPR_binding"].astype(str) == "1"]

# ------------------------------------------------------------------
# Main table
# ------------------------------------------------------------------
st.subheader("Antibody table")

has_flags = [c for c in filtered_df.columns if c.startswith("has_") and c.endswith("_image")]
rename_map = {}
for flag in has_flags:
    base = flag[len("has_") : -len("_image")] if flag.endswith("_image") else flag
    rename_map[flag] = f"{base.replace('_', ' ').strip().title()} image"

ag_columns = display_present.copy()


for flag in has_flags:
    if flag not in ag_columns:
        ag_columns.append(flag)
if "excel_row" in filtered_df.columns and "excel_row" not in ag_columns:
    ag_columns.append("excel_row")

ag_table = filtered_df[ag_columns].copy()
ag_table = ag_table.rename(columns=rename_map)

if "selected_excel_row" not in st.session_state and not filtered_df.empty:
    st.session_state.selected_excel_row = int(filtered_df.iloc[0]["excel_row"])

if AGGRID_AVAILABLE:
    gb = GridOptionsBuilder.from_dataframe(ag_table)
    gb.configure_selection(selection_mode="single", use_checkbox=False)
    grid_options = gb.build()
    grid_response = AgGrid(
        ag_table,
        gridOptions=grid_options,
        height=300,
        fit_columns_on_grid_load=True,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
    )

    def _normalize_selected_rows(obj):
        if obj is None:
            return []
        try:
            if isinstance(obj, pd.DataFrame):
                return obj.to_dict("records")
        except Exception:
            pass
        if isinstance(obj, dict):
            return [obj]
        if isinstance(obj, (list, tuple)):
            return list(obj)
        try:
            return list(obj)
        except Exception:
            return []

    selected_rows_raw = grid_response.get("selected_rows", [])
    selected_rows = _normalize_selected_rows(selected_rows_raw)
    if len(selected_rows) > 0:
        try:
            sel_row = selected_rows[0]
            sel_excel = None
            if isinstance(sel_row, dict):
                sel_excel = sel_row.get("excel_row")
                if sel_excel is None:
                    for k, v in sel_row.items():
                        if str(k).lower().replace(" ", "_") == "excel_row":
                            sel_excel = v
                            break
            if sel_excel is not None:
                st.session_state.selected_excel_row = int(sel_excel)
        except Exception as e:
            st.error(f"Selection handling error: {e}")
else:
    st.info("Install `streamlit-aggrid` for clickable row selection.")

if filtered_df.empty:
    st.info("No rows match the selected filters.")
    st.stop()

if "selected_excel_row" not in st.session_state:
    st.session_state.selected_excel_row = int(filtered_df.iloc[0]["excel_row"])

if st.session_state.selected_excel_row not in filtered_df["excel_row"].values:
    st.session_state.selected_excel_row = int(filtered_df.iloc[0]["excel_row"])

selected_row = filtered_df[filtered_df["excel_row"] == st.session_state.selected_excel_row].iloc[0]

st.markdown(f"## Details: {selected_row['Antibody_name']}")

st.download_button(
    label=f"Download {uploaded_file.name}",
    data=file_bytes,
    file_name=uploaded_file.name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ------------------------------------------------------------------
# Tabs
# ------------------------------------------------------------------
image_tab, raw_tab, gallery_tab = st.tabs(["Images", "Raw row", "Gallery"])

with image_tab:
    has_flags_local = [c for c in app_df.columns if c.startswith("has_") and c.endswith("_image")]
    image_columns = [f[len("has_") : -len("_image")] for f in has_flags_local]

    nav_col1, nav_col2 = st.columns(2)
    with nav_col1:
        if st.button("← Prev"):
            try:
                excel_rows_list = list(filtered_df["excel_row"].astype(int))
                cur = int(st.session_state.selected_excel_row)
                pos = excel_rows_list.index(cur) if cur in excel_rows_list else 0
                st.session_state.selected_excel_row = int(excel_rows_list[max(pos - 1, 0)])
                st.rerun()
            except Exception:
                pass
    with nav_col2:
        if st.button("Next →"):
            try:
                excel_rows_list = list(filtered_df["excel_row"].astype(int))
                cur = int(st.session_state.selected_excel_row)
                pos = excel_rows_list.index(cur) if cur in excel_rows_list else 0
                st.session_state.selected_excel_row = int(excel_rows_list[min(pos + 1, len(excel_rows_list) - 1)])
                st.rerun()
            except Exception:
                pass

    st.markdown("### Raw row (quick view)")
    raw_columns = [
        c for c in app_df.columns
        if not c.endswith("_bytes") and not c.startswith("has_")
    ]
    st.dataframe(pd.DataFrame([selected_row[raw_columns].to_dict()]), use_container_width=True, hide_index=True)
    st.caption("Note: EC50_Image and SPR-Mu often show '#VALUE!' because images are stored as rich-value cell images.")

    if not image_columns:
        st.info("No image-like columns were detected in the workbook.")
    else:
        images_to_show = []
        for col in image_columns:
            img_bytes = selected_row.get(f"{col}_bytes")
            if img_bytes is not None:
                images_to_show.append((col, img_bytes))

        for row_imgs in [images_to_show[i : i + 2] for i in range(0, min(len(images_to_show), 4), 2)]:
            cols = st.columns(2)
            for idx, (col_name, img_bytes) in enumerate(row_imgs):
                with cols[idx]:
                    st.markdown(f"**{col_name}**")
                    target_width = 200 if "chromatogram" in col_name.lower() else 600
                    try:
                        if PIL_AVAILABLE:
                            img = Image.open(BytesIO(img_bytes))
                            if img.width > target_width:
                                ratio = target_width / float(img.width)
                                img = img.resize((target_width, int(img.height * ratio)), Image.LANCZOS)
                            st.image(img, width=target_width)
                        else:
                            st.image(img_bytes, width=target_width)
                    except Exception:
                        st.image(img_bytes, width=target_width)

    st.markdown("### Summary")
    summary_fields = [f for f in display_present if f in selected_row.index]
    st.table(pd.DataFrame([{"Field": f, "Value": format_scalar(selected_row.get(f))} for f in summary_fields]))

    st.markdown("### Workbook info")
    st.write(f"**File:** {uploaded_file.name}")
    st.write(f"**Sheet names:** {', '.join(workbook_info['sheet_names'])}")
    st.write(f"**Rows:** {workbook_info['max_row']}  |  **Columns:** {workbook_info['max_column']}")


with raw_tab:
    st.markdown("### Raw row values")
    raw_columns = [
        c for c in app_df.columns
        if not c.endswith("_bytes") and not c.startswith("has_")
    ]
    st.dataframe(pd.DataFrame([selected_row[raw_columns].to_dict()]), use_container_width=True, hide_index=True)
    st.caption("Note: EC50_Image and SPR-Mu often show '#VALUE!' because images are stored as rich-value cell images.")

with gallery_tab:
    st.markdown("### All extracted images for filtered rows")
    image_columns_gallery = [f[len("has_") : -len("_image")] for f in app_df.columns if f.startswith("has_") and f.endswith("_image")]

    if not image_columns_gallery:
        st.info("No extracted images available for gallery.")
    else:
        gallery_choice = st.radio("Gallery type", image_columns_gallery, horizontal=True)
        for _, row in filtered_df.iterrows():
            image_bytes = row.get(f"{gallery_choice}_bytes")
            if image_bytes is None:
                continue
            st.markdown(f"**{row['Antibody_name']}**")
            if PIL_AVAILABLE:
                try:
                    img = Image.open(BytesIO(image_bytes))
                    max_w = 600
                    if img.width > max_w:
                        ratio = max_w / float(img.width)
                        img = img.resize((max_w, int(img.height * ratio)), Image.LANCZOS)
                    st.image(img, width=max_w)
                except Exception:
                    st.image(image_bytes, width=600)
            else:
                st.image(image_bytes, width=600)

st.markdown("---")
st.write(
    "This app supports any Excel workbook with Sheet1 containing one antibody per row. "
    "Use the Annotations tab to add notes per row, then download the file with your annotations included."
)
