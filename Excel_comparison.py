import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide", page_title="Excel Comparator")

st.markdown("<h2 style='text-align:center;'>🔍 Excel File Comparator with Export</h2>", unsafe_allow_html=True)
st.write("Upload two Excel files to compare and download the difference report.")

# ----------------------------------------
# FILE UPLOAD
# ----------------------------------------
file1 = st.file_uploader("Upload First Excel File", type=["xlsx"])
file2 = st.file_uploader("Upload Second Excel File", type=["xlsx"])

if file1 and file2:

    xls1 = pd.ExcelFile(file1)
    xls2 = pd.ExcelFile(file2)

    common_sheets = set(xls1.sheet_names).intersection(xls2.sheet_names)

    if not common_sheets:
        st.warning("No common sheets found between the two files.")
        st.stop()

    sheet = st.selectbox("Select Sheet to Compare", list(common_sheets))

    df1 = pd.read_excel(file1, sheet_name=sheet)
    df2 = pd.read_excel(file2, sheet_name=sheet)

    st.subheader(f"📄 Comparing Sheet: {sheet}")

    # ----------------------------------------
    # ALIGN
    # ----------------------------------------
    df1, df2 = df1.align(df2, join="outer", axis=1)

    # ----------------------------------------
    # DIFFERENCE REPORT
    # ----------------------------------------
    diff = pd.DataFrame(index=df1.index, columns=df1.columns)

    for col in df1.columns:
        diff[col] = df1[col].astype(str) + " → " + df2[col].astype(str)
        diff[col] = diff[col].where(df1[col] != df2[col], "")  # show only changed cells

    st.subheader("🟥 Differences (Old → New)")
    st.dataframe(diff)

    # Rows with differences
    changed_rows = df1[df1.ne(df2).any(axis=1)]

    st.subheader("🔧 Rows That Differ")
    st.dataframe(changed_rows)

    # ----------------------------------------
    # EXPORT TO CSV
    # ----------------------------------------
    csv_buffer = io.StringIO()
    diff.to_csv(csv_buffer, index=False)
    st.download_button(
        "⬇ Download CSV Difference Report",
        data=csv_buffer.getvalue(),
        file_name="difference_report.csv",
        mime="text/csv"
    )

    # ----------------------------------------
    # EXPORT TO EXCEL
    # ----------------------------------------
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        diff.to_excel(writer, sheet_name="Differences", index=False)
        changed_rows.to_excel(writer, sheet_name="ChangedRowsOnly", index=False)

    st.download_button(
        "⬇ Download Excel Difference Report",
        data=excel_buffer.getvalue(),
        file_name="difference_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
