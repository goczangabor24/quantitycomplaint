import re
from typing import List

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Issue Table Creator",
    page_icon="📋",
    layout="wide",
)


# =========================
# Helpers
# =========================
def read_input_file(uploaded_file) -> pd.DataFrame:
    file_name = uploaded_file.name.lower()

    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        return pd.read_excel(uploaded_file)

    if file_name.endswith(".csv"):
        return pd.read_csv(uploaded_file, sep="\t", encoding="utf-16")

    raise ValueError("Unsupported file format. Please upload an Excel or CSV file.")

def parse_multi_input(raw_text: str) -> List[str]:
    """
    Accept values like:
    12345,12346
    12345, 12346
    """
    if not raw_text:
        return []

    parts = [part.strip() for part in re.split(r",\s*", raw_text.strip())]
    return [part for part in parts if part]


def normalize_number(value):
    """
    Convert values to numeric safely.
    Works with:
    - normal numeric cells
    - strings like 1.234,56
    - strings like -12
    """
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    # keep only digits, minus, comma, dot
    text = re.sub(r"[^0-9,.\-]", "", text)
    if not text:
        return None

    # European number support:
    # 1.234,56 -> 1234.56
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        # If no comma, keep decimal dot if present
        # but remove thousands dots only if obviously European grouping
        # For simplicity, leave as-is.
        pass

    try:
        return float(text)
    except ValueError:
        return None


def to_int_series(series: pd.Series) -> pd.Series:
    """Convert a series to integer-like values safely."""
    numeric = series.apply(normalize_number)
    return numeric.fillna(0).round().astype(int)


def to_float_series(series: pd.Series) -> pd.Series:
    """Convert a series to float safely."""
    numeric = series.apply(normalize_number)
    return numeric.astype(float)


def format_price_series(series: pd.Series) -> pd.Series:
    """
    Keep price as numeric float for editing/export.
    """
    return to_float_series(series)


def build_output_table(source_rows: pd.DataFrame, issue_type: str) -> pd.DataFrame:
    """
    Column mapping (0-indexed):
    A = 0
    B = 1
    C = 2
    H = 7
    M = 12
    N = 13
    O = 14
    """
    col_a = source_rows.iloc[:, 0].astype(str)
    col_b = source_rows.iloc[:, 1].astype(str)
    col_c = source_rows.iloc[:, 2].astype(str)
    col_h = to_float_series(source_rows.iloc[:, 7]).round(2)
    col_m = to_int_series(source_rows.iloc[:, 12])
    col_n = to_int_series(source_rows.iloc[:, 13])
    col_o = to_int_series(source_rows.iloc[:, 14])

    if issue_type == "Damage":
        new_col_6 = col_m + col_o  # O is negative -> effectively subtraction
    else:  # Shortage
        new_col_6 = col_n

    output = pd.DataFrame(
        {
            "": [""] * len(source_rows),
            " ": [""] * len(source_rows),
            "LA-ID": col_a,
            "Shop Article": col_c,
            "Supplier Art-ID": col_b,
            "Quantity Delivered": new_col_6.astype(int),
            "Quantity Ordered": col_m.astype(int),
            "  ": [""] * len(source_rows),
            "Purchase Price 1": col_h,
        }
    )

    output["Purchase Price 1"] = output["Purchase Price 1"].map(
        lambda x: "" if pd.isna(x) else f"{x:.2f}".replace(".", ",")
    )

    return output.reset_index(drop=True)

    return output.reset_index(drop=True)


def deduplicate_rows_by_la_id(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove duplicate LA-IDs while keeping the first occurrence.
    LA-ID = column A = index 0 in source df before mapping,
    but here after mapping it's the 'LA-ID' column.
    """
    if "LA-ID" not in df.columns:
        return df
    return df.drop_duplicates(subset=["LA-ID"], keep="first").reset_index(drop=True)


# =========================
# UI
# =========================
st.title("📋 Issue Table Creator")

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xls", "csv"],
)

issue_type = st.selectbox(
    'Issue Type:',
    options=["Damage", "Shortage"],
)

if uploaded_file is not None:
    try:
        source_df = read_input_file(uploaded_file)
    except Exception as exc:
        st.error(f"Could not read the uploaded file: {exc}")
        st.stop()

    if source_df.empty:
        st.warning("The uploaded file is empty.")
        st.stop()

    if source_df.shape[1] < 15:
        st.error("The uploaded file must contain at least columns A to O.")
        st.stop()

    # Drop first row
    working_df = source_df.iloc[1:].copy().reset_index(drop=True)

    if working_df.empty:
        st.warning("There are no rows left after dropping the first row.")
        st.stop()

    # Normalize LA-ID column (A)
    la_id_col = working_df.columns[0]
    working_df[la_id_col] = working_df[la_id_col].astype(str).str.strip()

    # Precompute numeric columns
    col_m_num = to_int_series(working_df.iloc[:, 12])  # M
    col_n_num = to_int_series(working_df.iloc[:, 13])  # N
    col_o_num = working_df.iloc[:, 14].apply(normalize_number).fillna(0)  # O as float for negativity

    if issue_type == "Damage":
        # Base rows: only rows where O is negative
        damage_mask = col_o_num < 0
        base_rows = working_df[damage_mask].copy()

        st.subheader("Editable output")

        additional_input = st.text_input(
            "Further Items Affected",
            placeholder="Example: 12345, 12346, 12347",
        )

        additional_ids = parse_multi_input(additional_input)

        if additional_ids:
            additional_rows = working_df[working_df[la_id_col].isin(additional_ids)].copy()

            missing_ids = [
                item for item in additional_ids
                if item not in set(additional_rows[la_id_col].astype(str).tolist())
            ]
            if missing_ids:
                st.warning("These LA-ID values were not found: " + ", ".join(missing_ids))

            selected_rows = pd.concat([base_rows, additional_rows], ignore_index=True)
        else:
            selected_rows = base_rows.copy()

        if selected_rows.empty:
            st.info("No rows found with negative values in column O, and no further items were added.")
            st.stop()

        output_df = build_output_table(selected_rows, issue_type="Damage")
        output_df = deduplicate_rows_by_la_id(output_df)

    else:  # Shortage
        # Base rows: where N != M
        shortage_mask = col_n_num != col_m_num
        selected_rows = working_df[shortage_mask].copy()

        if selected_rows.empty:
            st.info("No rows found where column N and column M differ.")
            st.stop()

        st.subheader("Editable output")
        output_df = build_output_table(selected_rows, issue_type="Shortage")
        output_df = deduplicate_rows_by_la_id(output_df)

    edited_df = st.data_editor(
        output_df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="editable_output_table",
    )

    csv_data = edited_df.to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        "Download output as CSV",
        data=csv_data,
        file_name=f"{issue_type.lower()}_output.csv",
        mime="text/csv",
    )
