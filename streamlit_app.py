import streamlit as st
import pandas as pd
import io
from rapidfuzz import process, fuzz

st.set_page_config(page_title="VLOOKUP Tool", layout="wide")
st.title("🔍 JC VLOOKUP Tool ")

# Select mode: two files or one file with two sheets
mode = st.radio("Choose Comparison Mode:", ["Compare two Excel files", "Compare two sheets in one file"])

file_a = None
file_b = None

# Upload Excel files
if mode == "Compare two Excel files":
    file_a = st.file_uploader("Upload File A (Main File)", type=["xlsx"], key="file_a")
    file_b = st.file_uploader("Upload File B (Reference File)", type=["xlsx"], key="file_b")
elif mode == "Compare two sheets in one file":
    file_a = st.file_uploader("Upload Single Excel File", type=["xlsx"], key="file_single")
    file_b = file_a

# Proceed if both files are uploaded
if file_a and file_b:
    try:
        sheets_a = pd.ExcelFile(file_a).sheet_names
        sheets_b = pd.ExcelFile(file_b).sheet_names

        sheet_a = st.selectbox("Select Sheet from File A", sheets_a, key="sheet_a")
        sheet_b = st.selectbox("Select Sheet from File B", sheets_b, key="sheet_b")

        df_a = pd.read_excel(file_a, sheet_name=sheet_a)
        df_b = pd.read_excel(file_b, sheet_name=sheet_b)

        col_a = st.selectbox("Select Lookup Column from File A", df_a.columns)
        col_b = st.selectbox("Select Match Column from File B", df_b.columns)
        bring_cols = st.multiselect("Select Columns to Bring from File B", df_b.columns)

        use_fuzzy = st.checkbox("Enable Fuzzy Matching")
        threshold = st.slider("Fuzzy Match Threshold (only if enabled)", 0, 100, 80)

        if st.button("🔍 Run VLOOKUP"):
            if use_fuzzy:
                # Fuzzy Matching
                matches = []
                match_scores = []
                choices = df_b[col_b].astype(str).tolist()

                for value in df_a[col_a].astype(str):
                    best_match = process.extractOne(value, choices, scorer=fuzz.ratio)
                    if best_match and best_match[1] >= threshold:
                        matches.append(best_match[0])
                        match_scores.append(best_match[1])
                    else:
                        matches.append(None)
                        match_scores.append(None)

                df_a["_fuzzy_match_key"] = matches
                df_a["_match_score"] = match_scores

                df_b_renamed = df_b.rename(columns={col_b: "_fuzzy_match_key"})
                safe_bring_cols = [col for col in bring_cols if col != col_b]

                df_merged = pd.merge(df_a, df_b_renamed[["_fuzzy_match_key"] + safe_bring_cols],
                                     on="_fuzzy_match_key", how="left")
                st.success("✅ Fuzzy VLOOKUP Completed!")
            else:
                df_merged = pd.merge(df_a, df_b[[col_b] + bring_cols],
                                     left_on=col_a, right_on=col_b, how="left")
                st.success("✅ Exact VLOOKUP Completed!")

            st.dataframe(df_merged.head(100))

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_merged.to_excel(writer, index=False, sheet_name='VLOOKUP_Result')

            st.download_button(
                label="📥 Download Result Excel",
                data=output.getvalue(),
                file_name="vlookup_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ An error occurred: {e}")
