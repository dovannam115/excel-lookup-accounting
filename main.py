import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

st.title("🔁 Excel Lookup Tool")
st.markdown("Tải lên 2 file Excel: `BAN_RA.xlsx` và `NXT T4.xlsx`")

ban_ra_file = st.file_uploader("📤 Upload BAN_RA.xlsx", type=["xlsx"])
nxt_t4_file = st.file_uploader("📤 Upload NXT T4.xlsx", type=["xlsx"])

if ban_ra_file and nxt_t4_file:
    if st.button("🚀 Chạy tra cứu"):
        # Đọc sheet cần tra cứu từ file BAN_RA
        ban_ra_df = pd.read_excel(ban_ra_file, sheet_name="Smart_KTSC_OK")

        # Đọc dữ liệu từ file NXT T4 (sheet F8_D, bỏ qua 22 dòng đầu)
        nxt_t4_df = pd.read_excel(nxt_t4_file, sheet_name="F8_D", skiprows=22)
        nxt_t4_df.columns.values[[2, 4, 14]] = ['target_col', 'match_col', 'compare_col']

        q_col = ban_ra_df.columns[16]  # tương đương Q2
        z_col = ban_ra_df.columns[25]  # tương đương Z2

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

        # Thêm kết quả vào sheet
        ban_ra_df['lookup_result'] = results

        # Load tất cả sheet từ file gốc BAN_RA
        ban_ra_file.seek(0)  # reset stream
        with pd.ExcelFile(ban_ra_file) as xls:
            all_sheets = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}

        # Ghi đè sheet Smart_KTSC_OK
        all_sheets["Smart_KTSC_OK"] = ban_ra_df

        # Ghi file Excel kết quả
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)

        st.success("✅ Xử lý xong! Sheet `Smart_KTSC_OK` đã được cập nhật.")
        st.download_button(
            label="📥 Tải file kết quả",
            data=output,
            file_name="BAN_RA_lookup_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
