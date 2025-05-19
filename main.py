
import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🔁 Excel Lookup Tool")
st.markdown("Tải lên 2 file Excel: `BAN_RA.xlsx` và `NXT T4.xlsx`")

ban_ra_file = st.file_uploader("📤 Upload BAN_RA.xlsx", type=["xlsx"])
nxt_t4_file = st.file_uploader("📤 Upload NXT T4.xlsx", type=["xlsx"])

if ban_ra_file and nxt_t4_file:
    if st.button("🚀 Chạy tra cứu"):
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

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            ban_ra_df.to_excel(writer, sheet_name="Smart_KTSC_OK", index=False)
        output.seek(0)

        st.success("✅ Xử lý xong!")
        st.download_button(
            label="📥 Tải file kết quả",
            data=output,
            file_name="BAN_RA_lookup_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
