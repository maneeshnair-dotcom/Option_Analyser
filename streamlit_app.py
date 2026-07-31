import io
import tempfile
from pathlib import Path

import streamlit as st

from Nifty50Analyser_csvload import analyze_csv, save_excel_with_colors

st.set_page_config(page_title="NIFTY Options Analyser", layout="wide")
st.title("NIFTY Options Analyser")


def process_tab(label: str):
    uploaded = st.file_uploader(f"Upload {label} CSV", type="csv", key=f"{label}_upload")

    if uploaded is None:
        st.info(f"Upload the {label} CSV file to see the analysis.")
        return

    # Write to a temp file, preserving the original filename so the
    # analyser's auto-detect (stock name / interval) works the same way
    # it does when run from the command line.
    tmp_dir = Path(tempfile.mkdtemp())
    tmp_path = tmp_dir / uploaded.name
    tmp_path.write_bytes(uploaded.getvalue())

    try:
        result = analyze_csv(str(tmp_path))
    except Exception as e:
        st.error(f"Could not process {label} file: {e}")
        return

    st.dataframe(result, use_container_width=True)

    excel_buffer = io.BytesIO()
    save_excel_with_colors(result, excel_buffer)
    excel_buffer.seek(0)

    st.download_button(
        label=f"Download {label} Excel",
        data=excel_buffer,
        file_name=f"{label}_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"{label}_download",
    )


tab_ce, tab_pe = st.tabs(["CE", "PE"])

with tab_ce:
    process_tab("CE")

with tab_pe:
    process_tab("PE")
