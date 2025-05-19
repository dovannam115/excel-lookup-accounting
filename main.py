import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.title("🔁 Excel Lookup Tool")

ban_ra_file = st.file_uploader("📤 Upload file Bán ra", type=["xlsx"])
nxt_t4_file = st.file_uploader("📤 Upload file NXT", type=["xlsx"])

if ban_ra_file and nxt_t4_file:
    if st.button("🚀 Chạy Lookup"):
        # Đọc sheet cần tra cứu từ BAN_RA
        ban_ra_df = pd.read_excel(ban_ra_file, sheet_name="Smart_KTSC_OK")

        # Đọc dữ liệu từ NXT T4
        nxt_t4_df = pd.read_excel(nxt_t4_file, sheet_name="F8_D", skiprows=22)
        nxt_t4_df.columns.values[[2, 4, 14]] = ['target_col', 'match_col', 'compare_col']

        q_col = ban_ra_df.columns[16]
        z_col = ban_ra_df.columns[25]

        results = []
        for _, row in ban_ra_df.iterrows():
            q_value = row[q_col]
            z_value = row[z_col]
            mask = (nxt_t4_df['match_col'] == q_value) & (nxt_t4_df['compare_col'] <= z_value)
            filtered = nxt_t4_df[mask].copy()
            if not filtered.empty:
                filtered['diff'] = z_value - filtered['compare_col']
                matched_row = filtered.loc[filtered['diff'].idxmin()]
                results.append(matched_row['target_col'])
            else:
                results.append("Không tìm thấy")
        
        ban_ra_df['lookup_result'] = results

        # Load lại workbook gốc từ BAN_RA
        ban_ra_file.seek(0)
        wb = load_workbook(filename=ban_ra_file)

        # Ghi đè sheet Smart_KTSC_OK
        if "Smart_KTSC_OK" in wb.sheetnames:
            ws = wb["Smart_KTSC_OK"]
            wb.remove(ws)
        ws_new = wb.create_sheet("Smart_KTSC_OK")

        for r in dataframe_to_rows(ban_ra_df, index=False, header=True):
            ws_new.append(r)

        # Ghi lại workbook vào output
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ DONE")
        st.download_button(
            label="📥 Tải file kết quả",
            data=output,
            file_name="BAN_RA_lookup_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
