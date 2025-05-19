import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Excel Lookup Tool", layout="centered")
st.title("🔍 Excel Lookup Tool")

option = st.radio("📌 Chọn chức năng", ["🔁 Lookup Bán ra & NXT", "📄 Lookup theo mapping"])

# --- Chức năng 1 ---
if option == "🔁 Lookup Bán ra & NXT":
    ban_ra_file = st.file_uploader("📤 Upload file Bán ra", type=["xlsx"], key="ban_ra")
    nxt_t4_file = st.file_uploader("📤 Upload file NXT", type=["xlsx"], key="nxt_t4")

    if ban_ra_file and nxt_t4_file:
        if st.button("🚀 Chạy Lookup"):
            ban_ra_df = pd.read_excel(ban_ra_file, sheet_name="Smart_KTSC_OK")
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

            ban_ra_file.seek(0)
            wb = load_workbook(filename=ban_ra_file)
            if "Smart_KTSC_OK" in wb.sheetnames:
                ws = wb["Smart_KTSC_OK"]
                wb.remove(ws)
            ws_new = wb.create_sheet("Smart_KTSC_OK")
            for r in dataframe_to_rows(ban_ra_df, index=False, header=True):
                ws_new.append(r)
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

# --- Chức năng 2 ---
elif option == "📄 Lookup theo mapping":
    data_file = st.file_uploader("📤 Upload file Data", type=["xlsx"], key="data")
    mapping_file = st.file_uploader("📤 Upload file Mapping", type=["xlsx"], key="mapping")

    error_threshold = st.number_input("🔧 Nhập phần trăm sai số cho phép (vd: 3% là 0.03)", min_value=0.0, max_value=1.0, value=0.03, step=0.01)

    if data_file and mapping_file:
        if st.button("🚀 Chạy Lookup Mapping"):
            data_df = pd.read_excel(data_file)
            mapping_df = pd.read_excel(mapping_file)

            # Đặt tên cột đúng theo yêu cầu
            data_df.columns.values[[0, 4]] = ['TENDM', 'DGVND']
            mapping_df.columns.values[[2, 4, 6]] = ['target_col', 'match_col', 'compare_col']

            # Hàm mô phỏng chính xác công thức Excel MATCH(1,...)
            def lookup(row):
                try:
                    matches = mapping_df[
                        (mapping_df['match_col'] == row['TENDM']) &
                        (mapping_df['compare_col'].notnull()) &
                        (row['DGVND'] != 0) &
                        ((mapping_df['compare_col'] - row['DGVND']).abs() / row['DGVND'] <= error_threshold)
                    ]
                    if not matches.empty:
                        return matches.iloc[0]['target_col']
                    else:
                        return "Không tìm thấy"
                except:
                    return "Không tìm thấy"

            data_df['lookup_result'] = data_df.apply(lookup, axis=1)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                data_df.to_excel(writer, index=False, sheet_name="Data_Result")
            output.seek(0)

            st.success("✅ Lookup thành công!")
            st.download_button(
                label="📥 Tải file kết quả",
                data=output,
                file_name="data_lookup_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
