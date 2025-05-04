import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="VLOOKUP Tool", layout="wide")
st.title("🔍 Easy VLOOKUP Tool (Excel)")

mode = st.radio("Choose Comparison Mode:", ["Compare two Excel files", "Compare two sheets in one file"])

file_a = None
file_b = None

if mode == "Compare two Excel files":
    file_a = st.file_uploader("Upload File A (Main File)", type=["xlsx"], key="file_a")
    file_b = st.file_uploader("Upload File B (Reference File)", type=["xlsx"], key="file_b")
elif mode == "Compare two sheets in one file":
    file_a = st.file_uploader("Upload Single Excel File", type=["xlsx"], key="file_single")
    file_b = file_a  # Same file used for both sheets

if file_a and file_b:
    sheets_a = pd.ExcelFile(file_a).sheet_names
    sheets_b = pd.ExcelFile(file_b).sheet_names

    sheet_a = st.selectbox("Select Sheet from File A", sheets_a, key="sheet_a")
    sheet_b = st.selectbox("Select Sheet from File B", sheets_b, key="sheet_b")

    df_a = pd.read_excel(file_a, sheet_name=sheet_a)
    df_b = pd.read_excel(file_b, sheet_name=sheet_b)

    col_a = st.selectbox("Select Lookup Column from File A", df_a.columns)
    col_b = st.selectbox("Select Match Column from File B", df_b.columns)

    bring_cols = st.multiselect("Select Columns to Bring from File B", df_b.columns)

    if st.button("🔍 Run VLOOKUP"):
        df_merged = pd.merge(df_a, df_b[[col_b] + bring_cols], left_on=col_a, right_on=col_b, how="left")
        st.success("VLOOKUP Completed!")

        st.dataframe(df_merged.head(100))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_merged.to_excel(writer, index=False, sheet_name='VLOOKUP_Result')
            writer.save()
        st.download_button(
            label="📥 Download Result Excel",
            data=output.getvalue(),
            file_name="vlookup_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )