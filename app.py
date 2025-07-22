import os
import pandas as pd
import numpy as np
import streamlit as st
import zipfile
import io

st.set_page_config(page_title="Multi-Header Excel Report Generator", layout="wide")
st.title("📦 Upload ZIP & Generate Multi-Header Excel")

uploaded_file = st.file_uploader("Upload a ZIP file containing CSV/XLS/XLSX files", type=["zip"])

if uploaded_file:
    with zipfile.ZipFile(uploaded_file) as z:
        all_files = z.namelist()
        st.success(f"{len(all_files)} files found in the archive")

        file_instructions = [
            {
                "pattern": "Daily-AdX",
                "sheet_name": "Report data",
                "details": ["Ad Exchange impressions", "Ad Exchange revenue ($)"]
            },
            {
                "pattern": "Daily-Preferred Deals",
                "sheet_name": "Report data",
                "details": ["Total impressions", "Total CPM and CPC revenue ($)"]
            },
            {
                "pattern": "Magnite",
                "sheet_name": "Report",
                "details": ["Paid Impressions", "Publisher Net Revenue"]
            },
            {
                "pattern": "Net_Revenue_Report_for",
                "sheet_name": "Ad Requests, Ads Sent by Date",
                "details": ["Impressions", "Net Revenue"]
            },
            {
                "pattern": "Xandr_Daily_updated",
                "sheet_name": "Report Data",
                "details": ["imps", "revenue"]
            },
            {
                "pattern": "Citrus Ads Daily Report",
                "sheet_name": None,
                "details": ["Ad Renders", "Publisher Revenue"]
            },
            {
                "pattern": "Sharethrough",
                "sheet_name": None,
                "details": ["Rendered Impressions", "Earnings"]
            },
            {
                "pattern": "Daily-Open Bidding",
                "sheet_name": "Report data",
                "details": ["Yield group impressions", "Yield group estimated revenue ($)"]
            },
        ]

        file_dataframes = []

        for instruction in file_instructions:
            matched_file = next((f for f in all_files if instruction["pattern"] in f), None)
            if not matched_file:
                st.warning(f"❌ No file found for pattern: {instruction['pattern']}")
                continue

            try:
                if matched_file.endswith(".csv"):
                    df = pd.read_csv(z.open(matched_file))
                else:
                    df = pd.read_excel(z.open(matched_file), sheet_name=instruction["sheet_name"])
            except Exception as e:
                st.error(f"❌ Error reading {matched_file}: {e}")
                continue

            columns_data = {}

            if instruction["pattern"] == "Daily-AdX" and "Programmatic channel" in df.columns:
                grouped = df.groupby("Programmatic channel")
                for metric in instruction["details"]:
                    for breakdown in ["Open Auction", "Private Auction"]:
                        if breakdown in grouped.groups:
                            sub_df = grouped.get_group(breakdown)
                            match_cols = [col for col in sub_df.columns if col.strip().lower() == metric.strip().lower()]
                            if match_cols:
                                col_data = sub_df[match_cols[0]].reset_index(drop=True)
                            else:
                                col_data = pd.Series([np.nan] * len(sub_df))
                        else:
                            col_data = pd.Series([np.nan] * len(df))
                        columns_data[(matched_file, metric, breakdown)] = col_data
            else:
                for metric in instruction["details"]:
                    match_cols = [col for col in df.columns if col.strip().lower() == metric.strip().lower()]
                    if match_cols:
                        col_data = df[match_cols[0]].reset_index(drop=True)
                    else:
                        col_data = pd.Series([np.nan] * len(df))
                    columns_data[(matched_file, metric, "")] = col_data

            file_df = pd.DataFrame(columns_data)
            file_df.columns = pd.MultiIndex.from_tuples(file_df.columns)
            file_dataframes.append(file_df)

        if file_dataframes:
            final_df = pd.concat(file_dataframes, axis=1).fillna("")

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                sheetname = "Summary"
                workbook = writer.book
                worksheet = workbook.add_worksheet(sheetname)
                writer.sheets[sheetname] = worksheet

                top_format = workbook.add_format({
                    'bold': True, 'align': 'center', 'valign': 'vcenter',
                    'text_wrap': True, 'border': 1, 'bg_color': '#DCE6F1'
                })
                mid_format = workbook.add_format({
                    'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1
                })
                sub_format = workbook.add_format({
                    'align': 'center', 'valign': 'vcenter', 'border': 1
                })
                cell_format = workbook.add_format({'align': 'center'})

                level0 = final_df.columns.get_level_values(0)
                level1 = final_df.columns.get_level_values(1)
                level2 = final_df.columns.get_level_values(2)

                col_idx = 0
                start_idx = 0
                while start_idx < len(final_df.columns):
                    current_label = level0[start_idx]
                    span = 1
                    while start_idx + span < len(final_df.columns) and level0[start_idx + span] == current_label:
                        span += 1
                    worksheet.merge_range(0, col_idx, 0, col_idx + span - 1, current_label, top_format)
                    col_idx += span
                    start_idx += span

                for i in range(len(final_df.columns)):
                    worksheet.write(1, i, level1[i], mid_format)
                    worksheet.write(2, i, level2[i] if level2[i] else '', sub_format)

                for row_idx in range(len(final_df)):
                    for col_idx, val in enumerate(final_df.iloc[row_idx]):
                        worksheet.write(row_idx + 3, col_idx, val, cell_format)

                worksheet.freeze_panes(3, 0)
                worksheet.autofilter(2, 0, 2 + len(final_df), len(final_df.columns) - 1)
                worksheet.set_column(0, len(final_df.columns) - 1, 30)

            st.download_button(
                label="📥 Download Final Excel Report",
                data=buffer.getvalue(),
                file_name="Full_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No matching files found in ZIP.")
