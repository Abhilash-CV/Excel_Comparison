import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide", page_title="Excel/CSV Comparator")

st.markdown("<h2 style='text-align:center;'>🔍 Excel / CSV Comparator with Export</h2>", unsafe_allow_html=True)
st.write("Upload two files (Excel or CSV) to compare and download the difference report.")

# ------------------------------------------------------
# Function to read ANY uploaded file (CSV or Excel)
# ------------------------------------------------------
def load_file(uploaded_file):
    if uploaded_file is None:
        return None, None  # df, sheetnames

    filename = uploaded_file.name.lower()

    if filename.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
        return df, None  # CSV has no sheets

    elif filename.endswith(".xlsx"):
        xls = pd.ExcelFile(uploaded_file)
        return xls, xls.sheet_names

    else:
        st.error("Unsupported file type. Upload only CSV or XLSX.")
        return None, None


# ------------------------------------------------------
# FILE UPLOADS
# ------------------------------------------------------
file1 = st.file_uploader("Upload First File (CSV or Excel)", type=["csv", "xlsx"])
file2 = st.file_uploader("Upload Second File (CSV or Excel)", type=["csv", "xlsx"])

if file1 and file2:

    obj1, sheets1 = load_file(file1)
    obj2, sheets2 = load_file(file2)

    # ------------------------------------------------------
    # CASE 1: BOTH CSV files
    # ------------------------------------------------------
    if isinstance(obj1, pd.DataFrame) and isinstance(obj2, pd.DataFrame):

        st.subheader("📄 Comparing CSV Files (Single Sheet Mode)")
        df1, df2 = obj1.align(obj2, join="outer", axis=1)

        # Create diff
        diff = df1.astype(str) + " → " + df2.astype(str)
        diff = diff.where(df1.ne(df2), "")

        st.dataframe(diff)

        # Export section
        csv_buffer = io.StringIO()
        diff.to_csv(csv_buffer, index=False)
        st.download_button("⬇ Download CSV Difference Report", csv_buffer.getvalue(),
                           "difference_report.csv", "text/csv")

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
            diff.to_excel(writer, sheet_name="Differences", index=False)

        st.download_button("⬇ Download Excel Difference Report", excel_buffer.getvalue(),
                           "difference_report.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ------------------------------------------------------
    # CASE 2: BOTH Excel files
    # ------------------------------------------------------
    elif sheets1 and sheets2:

        common_sheets = set(sheets1).intersection(sheets2)

        if not common_sheets:
            st.error("No common sheets found between the two Excel files.")
            st.stop()

        sheet = st.selectbox("Select Sheet to Compare", sorted(common_sheets))

        df1 = pd.read_excel(file1, sheet_name=sheet)
        df2 = pd.read_excel(file2, sheet_name=sheet)

        st.subheader(f"📄 Comparing Excel Sheet: {sheet}")

        df1, df2 = df1.align(df2, join="outer", axis=1)

        diff = df1.astype(str) + " → " + df2.astype(str)
        diff = diff.where(df1.ne(df2), "")

        st.dataframe(diff)

        # Export
        csv_buffer = io.StringIO()
        diff.to_csv(csv_buffer, index=False)

        st.download_button("⬇ Download CSV Difference Report", csv_buffer.getvalue(),
                           "difference_report.csv", "text/csv")

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
            diff.to_excel(writer, sheet_name="Differences", index=False)

        st.download_button("⬇ Download Excel Difference Report", excel_buffer.getvalue(),
                           "difference_report.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    else:
        st.error("You cannot compare CSV with Excel. Upload both files in same format.")
