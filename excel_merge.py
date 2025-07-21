


import streamlit as st
import pandas as pd
from io import BytesIO
import tempfile
import os

st.set_page_config(page_title="Excel Merger App", layout="centered")

st.title("📊 Excel Sheet Merger")
st.subheader("Upload multiple Excel files (.xls or .xlsx), and merge them into one combined file.")

uploaded_files = st.file_uploader("Upload Excel Files", type=["xls", "xlsx"], accept_multiple_files=True)
merge_clicked = st.button("🔁 Merge Files", key="merge_button")

if merge_clicked:
    if uploaded_files:
        all_dfs = []

        for uploaded_file in uploaded_files:
            # Save uploaded file to a temporary path
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp:
                tmp.write(uploaded_file.read())
                temp_path = tmp.name

            try:
                if uploaded_file.name.endswith('.xls'):
                    df = pd.read_excel(temp_path, engine='xlrd')
                else:
                    df = pd.read_excel(temp_path)
            except Exception as e:
                #st.warning(f"⚠️ {uploaded_file.name} is not a real Excel file. Trying as HTML table...")
                try:
                    df = pd.read_html(temp_path)[0]
                except Exception as e2:
                    st.error(f"❌ Failed to read `{uploaded_file.name}` even as HTML.")
                    os.unlink(temp_path)  # cleanup
                    continue

            all_dfs.append(df)
            os.unlink(temp_path)  # cleanup

        if all_dfs:
            merged_df = pd.concat(all_dfs, ignore_index=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                merged_df.to_excel(writer, index=False)
            output.seek(0)

            st.success("✅ Merged successfully!")
            st.download_button(
                label="📥 Download Merged Excel",
                data=output,
                file_name="Merged_School_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ No valid data found in the uploaded files.")
    else:
        st.warning("⚠️ Please upload at least one Excel file.")
