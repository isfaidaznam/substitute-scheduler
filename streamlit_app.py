import streamlit as st
import pandas as pd
import numpy as np
import openpyxl

if 'sheet_to_load' not in st.session_state:
    st.session_state.sheet_to_load = None

st.title("Subtitute Teacher Timetable Generator")

# Utility functions
def transform_gsheet_url(url: str) -> str:
    if "docs.google.com/spreadsheets" in url and "/edit" in url:
        url = url.split("/edit")[0] + "/export?format=xlsx"
    return url

def remove_first_row_if_none(df: pd.DataFrame) -> pd.DataFrame:
    if df.shape[0] == 0:
        return df
    first_row = df.iloc[0]
    if first_row.isnull().all():
        return df.iloc[1:].reset_index(drop=True)
    else:
        return df

def remove_first_column_if_none(df: pd.DataFrame) -> pd.DataFrame:
    if df.shape[1] == 0:
        return df
    first_col = df.iloc[:, 0]
    if first_col.isnull().all():
        return df.iloc[:, 1:].reset_index(drop=True)
    else:
        return df

def rename_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    counts = {}
    new_cols = []
    for col in df.columns:
        if col not in counts:
            counts[col] = 0
            new_cols.append(col)
        else:
            counts[col] += 1
            new_cols.append(f"{col}_({counts[col]})")
    df.columns = new_cols
    return df

# Main form
with st.form("gsheet_form"):
    gsheet_url = st.text_input("Paste your Google Sheet URL:")
    submitted = st.form_submit_button("Submit Google Sheet URL")

if submitted and gsheet_url:
    try:
        export_url = transform_gsheet_url(gsheet_url)
        xls = pd.ExcelFile(export_url, engine='openpyxl')
        sheets = xls.sheet_names

        submitted_sheet = None
        sheet_to_load = sheets[0]
        if len(sheets) > 1:
            with st.form("sheet_form"):
                sheet_to_load = st.selectbox("Multiple sheets found. Select one to load:", sheets)
                submitted_sheet = st.form_submit_button("Submit Sheet")

        if sheet_to_load and submitted_sheet:
            time_table_data = pd.read_excel(export_url, sheet_name=sheet_to_load, engine='openpyxl')

            # Clean data
            while True:
                prev_len = len(time_table_data)
                time_table_data = remove_first_row_if_none(time_table_data)
                if len(time_table_data) == prev_len:
                    break

            time_table_data = time_table_data.replace(['', ' ', 'NA', 'null', None], np.nan)
            time_table_data = time_table_data.dropna(axis=1, how='all')

            if len([key for key in time_table_data.keys().to_list() if "unnamed" in key.lower()]) >= len(time_table_data.columns) / 2:
                time_table_data.columns = time_table_data.iloc[0]
                time_table_data = time_table_data[1:].reset_index(drop=True)

            time_table_data = rename_duplicate_columns(time_table_data)
            time_table_data = time_table_data.dropna(axis=0, how='all')

            st.success(f"Data loaded successfully from sheet: {sheet_to_load}")
            st.dataframe(time_table_data.head())

            # Column Extraction
            extract_column = {
                "teacher's name": "teacher",
                "day of the week": "day",
                "class's time": "time",
                "class's subject": "subject",
                "class's name": "class"
            }

            selected_columns = {}
            for ec_key in extract_column.keys():
                st.subheader(f"Select column for: {ec_key}")
                column = st.selectbox(f"Which column represents {ec_key}?", time_table_data.columns, key=ec_key)
                selected_columns[extract_column[ec_key]] = column

            for key, col_name in selected_columns.items():
                time_table_data[key] = time_table_data[col_name]

            final_columns = list(selected_columns.keys())
            time_table_data = time_table_data[final_columns]
            time_table_data = time_table_data.dropna(axis=0, how='all')

            st.subheader("Final Extracted Data")
            st.dataframe(time_table_data)

            # Download
            def to_excel(df):
                from io import BytesIO
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                return output.getvalue()

            st.download_button(
                "Download as Excel",
                data=to_excel(time_table_data),
                file_name="cleaned_timetable.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"An error occurred: {e}")
