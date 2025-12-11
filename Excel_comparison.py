import streamlit as st
import pandas as pd

st.set_page_config(layout="wide", page_title="Excel Comparator")

st.markdown("<h2 style='text-align:center;'>🔍 Excel File Comparator</h2>", unsafe_allow_html=True)
st.write("Upload two Excel files to compare them sheet-wise and cell-wise.")

# ----------------------------------------
# FILE UPLOAD
# ----------------------------------------
file1 = st.file_uploader("Upload First Excel File", type=["xlsx"])
file2 = st.file_uploader("Upload Second Excel File", type=["xlsx"])

if file1 and file2:

    # Load Excel files
    try:
        xls1 = pd.ExcelFile(file1)
        xls2 = pd.ExcelFile(file2)
    except:
        st.error("Unable to read one of the files. Please upload valid Excel (.xlsx) files.")
        st.stop()

    # List common sheets
    common_sheets = set(xls1.sheet_names).intersection(xls2.sheet_names)

    if not common_sheets:
        st.warning("No common sheets found between the two files.")
    else:
        sheet = st.selectbox("Select Sheet to Compare", list(common_sheets))

        df1 = pd.read_excel(file1, sheet_name=sheet)
        df2 = pd.read_excel(file2, sheet_name=sheet)

        st.subheader(f"📄 Comparing Sheet: {sheet}")

        # ----------------------------------------
        # ALIGN DATAFRAMES BY COLUMNS
        # ----------------------------------------
        df1, df2 = df1.align(df2, join="outer", axis=1)

        # ----------------------------------------
        # FIND DIFFERENCES
        # ----------------------------------------
        diff_mask = df1.ne(df2)

        # Highlight function
        def highlight_changes(val):
            return "background-color: #ffdddd" if val else ""

        st.subheader("🟥 Cells with Differences Highlighted (True = Different)")
        st.dataframe(diff_mask.style.applymap(highlight_changes))

        # ----------------------------------------
        # SHOW ORIGINAL VALUES SIDE BY SIDE
        # ----------------------------------------
        st.subheader("📌 Data From File 1")
        st.dataframe(df1)

        st.subheader("📌 Data From File 2")
        st.dataframe(df2)

        # ----------------------------------------
        # SHOW ONLY ROWS WITH AT LEAST ONE DIFFERENCE
        # ----------------------------------------
        changed_rows = df1[df1.ne(df2).any(axis=1)]

        st.subheader("🔧 Rows That Differ")
        st.dataframe(changed_rows)
