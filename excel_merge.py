import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
from io import BytesIO, StringIO
import tempfile
import os

st.set_page_config(page_title="Excel Merger App", layout="centered")
st.title("📊 Excel Sheet Merger")
st.subheader("Upload multiple Excel files (.xls or .xlsx), and merge them into one combined file.")

uploaded_files = st.file_uploader("Upload Excel Files", type=["xls", "xlsx"], accept_multiple_files=True)
merge_clicked = st.button("🔁 Merge Files", key="merge_button")

def extract_table_from_html(file_path):
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        soup = BeautifulSoup(f, 'lxml')
        table = soup.find("table")
        if table is None:
            return None
        rows = table.find_all("tr")
        data = []
        for row in rows:
            cols = [col.get_text(strip=True) for col in row.find_all(["td", "th"])]
            data.append(cols)
        df = pd.DataFrame(data)
        df.columns = df.iloc[0]  # First row as header
        df = df.drop(index=0).reset_index(drop=True)
        return df

if merge_clicked:
    if uploaded_files:
        all_dfs = []

        for uploaded_file in uploaded_files:
            suffix = os.path.splitext(uploaded_file.name)[1]
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(uploaded_file.read())
                temp_path = tmp.name

            try:
                if uploaded_file.name.endswith('.xls'):
                    df = pd.read_excel(temp_path, engine='xlrd')
                else:
                    df = pd.read_excel(temp_path)
            except Exception:
                st.warning(f"⚠️ {uploaded_file.name} is not a real Excel file. Trying custom HTML parsing...")
                try:
                    df = extract_table_from_html(temp_path)
                    if df is None:
                        raise Exception("No table found")
                except Exception:
                    st.error(f"❌ Failed to read `{uploaded_file.name}` even as custom HTML.")
                    os.unlink(temp_path)
                    continue

            df.columns = pd.io.parsers.ParserBase({'names': df.columns})._maybe_dedup_names(df.columns)
            df.columns = df.columns.str.strip()
            all_dfs.append(df)
            os.unlink(temp_path)

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
