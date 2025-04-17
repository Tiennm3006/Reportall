import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from chromadb import PersistentClient
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

# ---------- KHỞI TẠO CHROMADB ---------- #
client = PersistentClient(path="./chroma_storage")
collection = client.get_or_create_collection(name="baocao_files")

# ---------- HÀM XỬ LÝ ---------- #
def save_bar_chart(data, x_col, y_col, title):
    fig, ax = plt.subplots()
    ax.bar(data[x_col], data[y_col])
    ax.set_title(title)
    ax.set_ylabel(y_col)
    ax.tick_params(axis='x', rotation=45)
    fig.tight_layout()
    buffer = BytesIO()
    fig.savefig(buffer, format='png')
    plt.close(fig)
    buffer.seek(0)
    return buffer

def save_overall_chart(data, x_col, y_col, title):
    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(data[x_col], data[y_col])
    ax.set_title(title)
    ax.set_ylabel(y_col)
    ax.tick_params(axis='x', rotation=45)
    for bar in bars:
        height = bar.get_height()
        ax.annotate(f"{height:,.0f}", xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')
    fig.tight_layout()
    buffer = BytesIO()
    fig.savefig(buffer, format='png')
    plt.close(fig)
    buffer.seek(0)
    return buffer

def set_table_border(table):
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else tbl._element.get_or_add_tblPr()
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'auto')
        tblBorders.append(border)
    tblPr.append(tblBorders)

def export_word_report(tong_quan, full_df, top3, bot3, charts, nhan_xet, filename):
    doc = Document()
    doc.add_heading("BÁO CÁO ĐÁNH GIÁ", 0)

    doc.add_heading("Tổng quan", level=1)
    for k, v in tong_quan.items():
        doc.add_paragraph(f"{k}: {v}")

    doc.add_heading("Nhận xét", level=1)
    doc.add_paragraph(nhan_xet)

    def add_table(title, df):
        doc.add_heading(title, level=2)
        table = doc.add_table(rows=1, cols=len(df.columns))
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        for i, col in enumerate(df.columns):
            hdr_cells[i].text = str(col)
        for _, row in df.iterrows():
            row_cells = table.add_row().cells
            for i, val in enumerate(row):
                row_cells[i].text = str(val)
        set_table_border(table)

    add_table("Toàn bộ dữ liệu", full_df)
    add_table("Top 3 cao nhất", top3)
    add_table("Top 3 thấp nhất", bot3)

    doc.add_heading("Biểu đồ minh họa", level=1)
    for chart in charts:
        doc.add_picture(chart, width=Inches(5.5))

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------- TAB 1 ---------- #
with tab1:
    uploaded_file = st.file_uploader("Tải lên file Excel chứa sheet 'Tong hop luy ke'", type=["xlsx"], key="kiemtra")

    if uploaded_file:
        df = pd.read_excel(uploaded_file, sheet_name="Tong hop luy ke", header=None)
        df_cleaned = df.iloc[4:].copy()
        df_cleaned.columns = ["STT", "Điện lực", "1P_GT", "1P_TT", "3P_GT", "3P_TT",
                              "TU", "TI", "Tổng công tơ", "Kế hoạch", "Tỷ lệ"]
        df_cleaned = df_cleaned[df_cleaned["Điện lực"].notna()]
        cols_to_numeric = ["1P_GT", "1P_TT", "3P_GT", "3P_TT", "TU", "TI", "Tổng công tơ", "Kế hoạch", "Tỷ lệ"]
        df_cleaned[cols_to_numeric] = df_cleaned[cols_to_numeric].apply(pd.to_numeric, errors='coerce')
        df_cleaned["Tỷ lệ"] = df_cleaned["Tỷ lệ"] * 100

        total_current = df_cleaned["Tổng công tơ"].sum()
        total_plan = df_cleaned["Kế hoạch"].sum()
        current_date = datetime.now()
        days_passed = (current_date - datetime(2025, 1, 1)).days
        days_total = (datetime(2025, 9, 30) - datetime(2025, 1, 1)).days
        avg_per_day = total_current / days_passed
        forecast_total = avg_per_day * days_total
        forecast_ratio = forecast_total / total_plan
        danh_gia = "✅ ĐẠT kế hoạch" if forecast_ratio >= 1 else "❌ CHƯA ĐẠT kế hoạch"

        nhan_xet = f"Tính đến hiện tại, khối lượng công tác đạt khoảng {forecast_ratio*100:.1f}%. Với tốc độ trung bình hiện nay ({avg_per_day:.2f} thiết bị/ngày), dự báo tổng thực hiện đến 30/09/2025 là {int(forecast_total):,} thiết bị. Do đó, {danh_gia.lower()}."

        df_sorted = df_cleaned.sort_values(by="Tỷ lệ", ascending=False)
        top_3 = df_sorted.head(3)
        bottom_3 = df_sorted.tail(3)

        st.subheader("Tổng quan và dự báo")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Tổng đã thực hiện", f"{total_current:,}")
            st.metric("Kế hoạch", f"{total_plan:,}")
            st.metric("Tốc độ TB/ngày", f"{avg_per_day:.2f}")
        with col2:
            st.metric("Dự báo đến 30/09/2025", f"{int(forecast_total):,}")
            st.metric("Tỷ lệ dự báo", f"{forecast_ratio*100:.2f}%")
            st.info(danh_gia)

        st.subheader("Top 3 tỷ lệ cao nhất")
        st.dataframe(top_3[["Điện lực", "Tổng công tơ", "Kế hoạch", "Tỷ lệ"]])

        st.subheader("Bottom 3 tỷ lệ thấp nhất")
        st.dataframe(bottom_3[["Điện lực", "Tổng công tơ", "Kế hoạch", "Tỷ lệ"]])

        chart_top = save_bar_chart(top_3, "Điện lực", "Tỷ lệ", "Top 3 Tỷ lệ cao")
        chart_bot = save_bar_chart(bottom_3, "Điện lực", "Tỷ lệ", "Bottom 3 Tỷ lệ thấp")
        chart_all = save_overall_chart(df_sorted, "Điện lực", "Tỷ lệ", "Tổng hợp Tỷ lệ hoàn thành")

        st.image(chart_top)
        st.image(chart_bot)
        st.image(chart_all)

        if st.button("📄 Xuất báo cáo Word", key="xuat_tab1"):
            tong_quan = {
                "Tổng đã thực hiện": f"{total_current:,}",
                "Kế hoạch": f"{total_plan:,}",
                "Tốc độ TB/ngày": f"{avg_per_day:.2f}",
                "Dự báo đến 30/09/2025": f"{int(forecast_total):,}",
                "Tỷ lệ dự báo": f"{forecast_ratio*100:.2f}%",
                "Đánh giá": danh_gia
            }
            report = export_word_report(tong_quan, df_cleaned[["Điện lực", "Tổng công tơ", "Kế hoạch", "Tỷ lệ"]], top_3, bottom_3, [chart_top, chart_bot, chart_all], nhan_xet, filename="Bao_cao_tab1.docx")
            st.download_button("📥 Tải báo cáo", data=report, file_name="Bao_cao_tab1.docx")

# ---------- TAB 2 ---------- #
with tab2:
    st.subheader("Phân tích công tác cắt điện do chưa trả tiền")

    uploaded_cut = st.file_uploader("Tải lên file Excel công tác cắt điện", type=["xlsx"], key="cut")

    if uploaded_cut:
        df_cut = pd.read_excel(uploaded_cut)
        df_cut.columns = df_cut.columns.str.strip()

        if "Điện lực" not in df_cut.columns:
            st.error("❌ Không tìm thấy cột 'Điện lực' trong file Excel. Vui lòng kiểm tra lại.")
            st.stop()

        df_cut = df_cut.dropna(subset=["Điện lực"])
        df_cut = df_cut[df_cut["Điện lực"].str.upper() != "TỔNG"]

        df_cut["Khách hàng nợ quá hạn chưa cắt điện"] = df_cut["Khách hàng nợ quá hạn chưa cắt điện"].astype(int)
        df_cut["Số tiền"] = df_cut["Số tiền"].astype(float)

        st.write("### Dữ liệu công tác cắt điện")
        st.dataframe(df_cut)

        top_kh = df_cut.sort_values(by="Khách hàng nợ quá hạn chưa cắt điện", ascending=False).head(3)
        bot_kh = df_cut.sort_values(by="Khách hàng nợ quá hạn chưa cắt điện", ascending=True).head(3)

        top_tien = df_cut.sort_values(by="Số tiền", ascending=False).head(3)
        bot_tien = df_cut.sort_values(by="Số tiền", ascending=True).head(3)

        st.write("### 📌 Tổng quan")
        tong_kh = df_cut['Khách hàng nợ quá hạn chưa cắt điện'].sum()
        tong_tien = df_cut['Số tiền'].sum()
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Tổng khách hàng chưa cắt", f"{tong_kh:,}")
        with col2:
            st.metric("Tổng số tiền", f"{tong_tien:,.0f} đ")

        st.write("### 🔼 Top 3 KH chưa cắt cao nhất")
        st.dataframe(top_kh[["Điện lực", "Khách hàng nợ quá hạn chưa cắt điện"]])
        st.write("### 🔽 Top 3 KH chưa cắt thấp nhất")
        st.dataframe(bot_kh[["Điện lực", "Khách hàng nợ quá hạn chưa cắt điện"]])

        st.write("### 💰 Top 3 số tiền nợ cao nhất")
        st.dataframe(top_tien[["Điện lực", "Số tiền"]])
        st.write("### 💸 Top 3 số tiền nợ thấp nhất")
        st.dataframe(bot_tien[["Điện lực", "Số tiền"]])

        chart_kh = save_bar_chart(top_kh, "Điện lực", "Khách hàng nợ quá hạn chưa cắt điện", "Top 3 KH chưa cắt")
        chart_tien = save_bar_chart(top_tien, "Điện lực", "Số tiền", "Top 3 Số tiền nợ")
        chart_all_kh = save_overall_chart(df_cut, "Điện lực", "Khách hàng nợ quá hạn chưa cắt điện", "Tổng KH chưa cắt theo Điện lực")

        nhan_xet = f"Hiện tại còn {tong_kh:,} khách hàng chưa bị cắt điện với tổng số tiền lên đến {tong_tien:,.0f} đ. Cần rà soát các đơn vị có số lượng lớn hoặc số tiền cao để ưu tiên xử lý."

        st.image(chart_kh)
        st.image(chart_tien)
        st.image(chart_all_kh)

        if st.button("📄 Xuất báo cáo Word", key="xuat_tab2"):
            tong_quan = {
                "Tổng KH chưa cắt": f"{tong_kh:,}",
                "Tổng số tiền": f"{tong_tien:,.0f} đ"
            }
            report = export_word_report(tong_quan, df_cut[["Điện lực", "Khách hàng nợ quá hạn chưa cắt điện", "Số tiền"]], top_kh, bot_kh, [chart_kh, chart_tien, chart_all_kh], nhan_xet, filename="Bao_cao_tab2.docx")
            st.download_button("📥 Tải báo cáo", data=report, file_name="Bao_cao_tab2.docx")
