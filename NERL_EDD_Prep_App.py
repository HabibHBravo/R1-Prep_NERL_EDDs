"""Streamlit interface for preparing NERL EDD workbooks."""

#V1.0

from pathlib import Path
import zipfile

import streamlit as st

from NERL_EDD_Prep import process_edd_bytes


st.set_page_config(
    page_title="NERL EDD Prep",
    page_icon=":bar_chart:",
    layout="centered",
)

st.title("NERL EDD Prep")
st.write(
    "Upload an Excel EDD workbook to apply the NERL field, date, result, "
    "and column-ordering rules."
)

uploaded_file = st.file_uploader(
    "Select an EDD workbook",
    type=["xlsx"],
    help="Only .xlsx workbooks are supported.",
)

if uploaded_file is not None:
    st.caption(f"Selected file: {uploaded_file.name}")

    if st.button(
        "Prepare and export EDD",
        type="primary",
        use_container_width=True,
    ):
        source_name = Path(uploaded_file.name).stem
        output_name = f"{source_name}_prepared.xlsx"

        try:
            with st.spinner("Preparing workbook..."):
                prepared_workbook = process_edd_bytes(uploaded_file.getvalue())

            st.success("EDD prepared successfully.")
            st.download_button(
                "Download prepared EDD",
                data=prepared_workbook,
                file_name=output_name,
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )
        except (OSError, ValueError, zipfile.BadZipFile) as error:
            st.error(f"Unable to prepare this workbook: {error}")
