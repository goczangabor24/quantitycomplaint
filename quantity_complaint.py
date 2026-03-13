import io
import re
from typing import List

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Affected Items Mapper", page_icon="📋", layout="wide")


# ---------- Helpers ----------
def normalize_number(value) -> float | None:
    """Convert European-style numbers like 1.234,56 to float."""
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None

    # Keep digits, minus sign, comma and dot
    text = re.sub(r"[^0-9,.-]", "", text)
    if not text:
        return None

    # European format: thousands='.', decimal=','
    text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def parse_affected_ids(raw_text: str) -> List[str]:
    if not raw_text:
        return []
    parts = [part.strip() for part in re.split(r",\s*", raw_text.strip())]
    return [part for part in parts if part]


def read_input_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file, sep="\t", encoding="utf-16")
    if name.endswith(".xlsx"):
        return pd.read_excel(uploaded_file)
    if name.endswith(".xls"):
        return pd.read_excel(uploaded_file)
    raise ValueError("Unsupported file format. Please upload a CSV or Excel file.")


def build_output(filtered_df: pd.DataFrame) -> pd.DataFrame:
    """
    Output layout:
    1 empty
    2 empty
    3 A
    4 C
    5 B
    6 L
    7 K
    8 empty
    9 H
    """
    output = pd.DataFrame(
        {
            " ": [""] * len(filtered_df),
            "  ": [""] * len(filtered_df),
            "LA-ID": filtered_df.iloc[:, 0].astype(str),           # A
            "Shop Article": filtered_df.iloc[:, 2],                # C
            "Supplier Art-ID": filtered_df.iloc[:, 1],             # B
            "Packages Delivered": filtered_df.iloc[:, 11],         # L
            "Packages Ordered": filtered_df.iloc[:, 10],           # K
            "   ": [""] * len(filtered_df),
            "Purchase Price 1": filtered_df.iloc[:, 7],            # H
        }
    )
    return output.reset_index(drop=True)


def to_clipboard_tsv(df: pd.DataFrame) -> str:
    buffer = io.StringIO()
    df.to_csv(buffer, sep="\t", index=False)
    return buffer.getvalue()


# ---------- UI ----------
st.title("📋 Affected Items Mapper")
st.caption("Upload the order export, automatically detect negative O-column rows, or manually filter by LA-ID.")

uploaded_file = st.file_uploader(
    "Upload file",
    type=["csv", "xlsx", "xls"],
    help="Supports the Zooplus-style CSV export and Excel files.",
)

if uploaded_file is not None:
    try:
        source_df = read_input_file(uploaded_file)
    except Exception as exc:
        st.error(f"Could not read the file: {exc}")
        st.stop()

    if source_df.empty:
        st.warning("The uploaded file is empty.")
        st.stop()

    if source_df.shape[1] < 15:
        st.error("The uploaded file does not have enough columns. I need at least columns A to O.")
        st.stop()

    # Drop the first row (Total row)
    working_df = source_df.iloc[1:].copy().reset_index(drop=True)

    if working_df.empty:
        st.warning("There are no data rows after removing the first row.")
        st.stop()

    # Normalize the LA-ID column for matching
    la_id_col = working_df.columns[0]
    working_df[la_id_col] = working_df[la_id_col].astype(str).str.strip()

    # O column = 15th column -> index 14
    o_col = working_df.columns[14]
    o_numeric = working_df.iloc[:, 14].apply(normalize_number)
    negative_mask = o_numeric.fillna(0) < 0
    negative_rows = working_df[negative_mask].copy()

    st.subheader("Source preview")
    st.dataframe(working_df, use_container_width=True, hide_index=True)

    if not negative_rows.empty:
        st.success(f"Found {len(negative_rows)} row(s) with a negative value in column O ({o_col}). Auto-filter applied.")
        filtered_df = negative_rows
    else:
        st.info("No negative values found in column O. Please provide affected LA-ID values.")
        affected_input = st.text_input(
            "Which items are affected?",
            placeholder="Example: 67354, 69579, 96105",
        )

        affected_ids = parse_affected_ids(affected_input)

        if not affected_input:
            st.stop()

        if not affected_ids:
            st.warning("Please enter at least one valid LA-ID.")
            st.stop()

        filtered_df = working_df[working_df[la_id_col].isin(affected_ids)].copy()

        missing_ids = [item for item in affected_ids if item not in set(filtered_df[la_id_col].tolist())]
        if missing_ids:
            st.warning("These LA-ID values were not found in the file: " + ", ".join(missing_ids))

    if filtered_df.empty:
        st.warning("No matching rows found for the current filter.")
        st.stop()

    output_df = build_output(filtered_df)

    st.subheader("Editable output")
    st.caption("You can edit the table below. The copyable TSV box updates from the edited version.")

    edited_df = st.data_editor(
        output_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key="editable_output_table",
    )

    st.subheader("Copyable output")
    st.caption("Copy everything from the box below and paste it directly into Excel.")
    tsv_text = to_clipboard_tsv(edited_df)
    st.text_area(
        "Excel-ready TSV",
        value=tsv_text,
        height=240,
    )

    csv_bytes = edited_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Download edited output as CSV",
        data=csv_bytes,
        file_name="affected_items_output.csv",
        mime="text/csv",
    )
