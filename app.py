import streamlit as st
import pandas as pd
import requests
import io
import re

st.set_page_config(
    page_title="Multi-Sheet Excel Comparison Tool",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Multi-Sheet Excel Comparison Tool")
st.write("Compare two Excel workbooks tab-by-tab. Choose whether to upload local files or paste public URLs (like Google Sheets).")

# --- Helper Functions ---
def get_direct_download_url(url: str) -> str:
    """Converts web preview links into direct download links."""
    url = url.strip()
    if "docs.google.com/spreadsheets" in url:
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
        if match:
            doc_id = match.group(1)
            return f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx"
    if "dropbox.com" in url:
        return url.replace("dl=0", "dl=1") if "dl=0" in url else url
    return url

def load_excel_from_url(url: str) -> pd.ExcelFile:
    download_url = get_direct_download_url(url)
    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(download_url, headers=headers, timeout=20)
    response.raise_for_status()
    return pd.ExcelFile(io.BytesIO(response.content))

def col_to_letter(col_idx: int) -> str:
    letter = ""
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter

# --- UI Setup: Tabs for Input Methods ---
tab1, tab2 = st.tabs(["📂 Upload Local Files", "🔗 Paste Web URLs"])

# Variables to hold the parsed Excel files once loaded
wb1, wb2 = None, None

# --- Tab 1: File Upload ---
with tab1:
    st.subheader("Drag and drop Excel files")
    col1_up, col2_up = st.columns(2)
    with col1_up:
        file1 = st.file_uploader("Upload Spreadsheet 1", type=["xlsx", "xls"], key="up1")
    with col2_up:
        file2 = st.file_uploader("Upload Spreadsheet 2", type=["xlsx", "xls"], key="up2")
        
    if st.button("🚀 Compare Uploaded Files", type="primary"):
        if not file1 or not file2:
            st.warning("Please upload both files to compare.")
        else:
            wb1 = pd.ExcelFile(file1)
            wb2 = pd.ExcelFile(file2)

# --- Tab 2: URL Input ---
with tab2:
    st.subheader("Paste Google Sheets or Direct Excel URLs")
    col1_url, col2_url = st.columns(2)
    with col1_url:
        url1 = st.text_input("Spreadsheet 1 URL", placeholder="https://docs.google.com/spreadsheets/d/...")
    with col2_url:
        url2 = st.text_input("Spreadsheet 2 URL", placeholder="https://docs.google.com/spreadsheets/d/...")

    if st.button("🚀 Compare URLs", type="primary"):
        if not url1 or not url2:
            st.warning("Please enter URLs for both spreadsheets.")
        else:
            with st.spinner("Downloading workbooks..."):
                try:
                    wb1 = load_excel_from_url(url1)
                    wb2 = load_excel_from_url(url2)
                except Exception as e:
                    st.error(f"Failed to download workbooks: {e}")

# --- Unified Processing & Comparison ---
# This block runs if workbooks were successfully loaded from either tab
if wb1 and wb2:
    with st.spinner("Comparing corresponding sheets..."):
        try:
            sheets1 = wb1.sheet_names
            sheets2 = wb2.sheet_names

            common_sheets = [s for s in sheets1 if s in sheets2]
            unmatched_sheets1 = [s for s in sheets1 if s not in sheets2]
            unmatched_sheets2 = [s for s in sheets2 if s not in sheets1]

            if unmatched_sheets1 or unmatched_sheets2:
                with st.expander("⚠️ Sheet Name Mismatches Detected"):
                    if unmatched_sheets1:
                        st.write(f"Sheets only in File 1: `{', '.join(unmatched_sheets1)}`")
                    if unmatched_sheets2:
                        st.write(f"Sheets only in File 2: `{', '.join(unmatched_sheets2)}`")

            if not common_sheets:
                st.error("No matching sheet names were found between the two workbooks. Make sure tab names match exactly.")
            else:
                all_diffs = []

                for sheet_name in common_sheets:
                    df1 = wb1.parse(sheet_name)
                    df2 = wb2.parse(sheet_name)

                    df1_filled = df1.fillna("[Empty]")
                    df2_filled = df2.fillna("[Empty]")

                    max_rows = max(len(df1), len(df2))
                    max_cols = max(len(df1.columns), len(df2.columns))

                    for row in range(max_rows):
                        for col in range(max_cols):
                            val1 = df1_filled.iat[row, col] if row < len(df1) and col < len(df1.columns) else "[Empty]"
                            val2 = df2_filled.iat[row, col] if row < len(df2) and col < len(df2.columns) else "[Empty]"

                            if str(val1) != str(val2):
                                cell_a1 = f"{col_to_letter(col)}{row + 2}" # +2 to account for 0-index & pandas header row
                                all_diffs.append({
                                    "Sheet Name": sheet_name,
                                    "Cell": cell_a1,
                                    "Row #": row + 2, # Account for header
                                    "Column #": col + 1,
                                    "File 1 Value": val1,
                                    "File 2 Value": val2
                                })

                st.divider()

                if all_diffs:
                    diff_df = pd.DataFrame(all_diffs)
                    
                    st.metric(label="Total Differences Across All Sheets", value=len(diff_df))
                    
                    # Tab-based filtering view
                    st.subheader("Differences Breakdown")
                    selected_sheet = st.selectbox("Filter report by sheet:", ["All Sheets"] + common_sheets)
                    
                    display_df = diff_df if selected_sheet == "All Sheets" else diff_df[diff_df["Sheet Name"] == selected_sheet]
                    st.dataframe(display_df, use_container_width=True)

                    # CSV Export
                    csv_data = diff_df.to_csv(index=False).encode("utf-8")
                    st.download_button(
                        label="📥 Download Full Comparison Report (CSV)",
                        data=csv_data,
                        file_name="multi_sheet_comparison_report.csv",
                        mime="text/csv"
                    )
                else:
                    st.success("🎉 All matching sheets across both workbooks are completely identical!")

        except Exception as e:
            st.error(f"An error occurred while comparing the workbooks: {e}")
