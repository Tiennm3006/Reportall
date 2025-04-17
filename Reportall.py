import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO
import os

st.set_page_config(page_title="Báo cáo công tác kinh doanh", layout="wide")
st.title("📊 Báo cáo công tác kinh doanh")

# Tabs chia hai phần
tab1, tab2 = st.tabs(["📋 Kiểm tra hệ thống đo đếm", "🔌 Cắt điện do chưa trả tiền"])

# ---------- HÀM XỬ LÝ ---------- #
...

# [TAB 1 giữ nguyên như trước]

# ---------- TAB 2: PHÂN TÍCH CẮT ĐIỆN ---------- #
with tab2:
    uploaded_cut = st.file_uploader("📤 Tải lên file Excel công tác cắt điện", type=["xlsx"], key="catdien")

    if uploaded_cut:
        df_cut = pd.read_excel(uploaded_cut)
        df_cut.columns = df_cut.columns.str.strip()

        if "Điện lực" not in df_cut.columns:
            st.error("❌ Không tìm thấy cột 'Điện lực' trong file. Hãy kiểm tra lại.")
            st.stop()

        df_cut = df_cut.dropna(subset=["Điện lực"])
        df_cut = df_cut[df_cut["Điện lực"].str.upper() != "TỔNG"]
        df_cut["Khách hàng nợ quá hạn chưa cắt điện"] = df_cut["Khách hàng nợ quá hạn chưa cắt điện"].astype(int)
        df_cut["Số tiền"] = df_cut["Số tiền"].astype(float)

        tong_kh = df_cut["Khách hàng nợ quá hạn chưa cắt điện"].sum()
        tong_tien = df_cut["Số tiền"].sum()
        nhan_xet = f"Hiện tại còn {tong_kh:,} khách hàng chưa bị cắt điện với tổng số tiền nợ {tong_tien:,.0f} đ. Cần rà soát các đơn vị có số lượng lớn và số tiền cao để ưu tiên xử lý."

        top_kh = df_cut.sort_values(by="Khách hàng nợ quá hạn chưa cắt điện", ascending=False).head(3)
        bot_kh = df_cut.sort_values(by="Khách hàng nợ quá hạn chưa cắt điện", ascending=True).head(3)

        chart_top_kh = save_bar_chart(top_kh, "Điện lực", "Khách hàng nợ quá hạn chưa cắt điện", "Top 3 KH chưa cắt")
        chart_bot_kh = save_bar_chart(bot_kh, "Điện lực", "Khách hàng nợ quá hạn chưa cắt điện", "Bottom 3 KH chưa cắt")
        chart_all_kh = save_overall_chart(df_cut, "Điện lực", "Khách hàng nợ quá hạn chưa cắt điện", "Tổng hợp KH chưa cắt")

        st.metric("Tổng KH chưa cắt", f"{tong_kh:,}")
        st.metric("Tổng số tiền", f"{tong_tien:,.0f} đ")
        st.info(nhan_xet)

        st.subheader("📈 Biểu đồ")
        st.image(chart_top_kh)
        st.image(chart_bot_kh)
        st.image(chart_all_kh)

        if st.button("📄 Xuất báo cáo Word", key="tab2_word"):
            tong_quan = {
                "Tổng KH chưa cắt": f"{tong_kh:,}",
                "Tổng số tiền": f"{tong_tien:,.0f} đ"
            }
            word_file = export_word_report(tong_quan, df_cut[["Điện lực", "Khách hàng nợ quá hạn chưa cắt điện", "Số tiền"]], top_kh, bot_kh, [chart_top_kh, chart_bot_kh, chart_all_kh], nhan_xet, filename="Bao_cao_CatDien.docx")
            st.download_button("📥 Tải báo cáo Word", data=word_file, file_name="Bao_cao_CatDien.docx")
