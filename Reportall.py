import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO

st.set_page_config(page_title="Báo cáo công tác kinh doanh", layout="wide")
st.title("📊 Báo cáo công tác kinh doanh")

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

tab1, tab2 = st.tabs(["📋 Kiểm tra hệ thống đo đếm", "🔌 Cắt điện do chưa trả tiền"])

# ---------- TAB 1: PHÂN TÍCH HỆ THỐNG ĐO ĐẾM ---------- # [giữ nguyên nội dung đã có]

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
            word_file = export_word_report(
                tong_quan,
                df_cut[["Điện lực", "Khách hàng nợ quá hạn chưa cắt điện", "Số tiền"]],
                top_kh,
                bot_kh,
                [chart_top_kh, chart_bot_kh, chart_all_kh],
                nhan_xet,
                filename="Bao_cao_CatDien.docx")
            st.download_button("📥 Tải báo cáo Word", data=word_file, file_name="Bao_cao_CatDien.docx")
