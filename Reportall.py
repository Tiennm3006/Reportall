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

# ---------- TAB 1: PHÂN TÍCH HỆ THỐNG ĐO ĐẾM ---------- #
with tab1:
    uploaded_file = st.file_uploader("📤 Tải lên file Excel chứa sheet 'Tong hop luy ke'", type=["xlsx"], key="kiemtra")

    if uploaded_file:
        df = pd.read_excel(uploaded_file, sheet_name="Tong hop luy ke", header=None)
        df_cleaned = df.iloc[4:].copy()
        expected_columns = ["STT", "Điện lực", "1P_GT", "1P_TT", "3P_GT", "3P_TT", "TU", "TI", "Tổng công tơ", "Kế hoạch", "Tỷ lệ"]

        if df_cleaned.shape[1] < len(expected_columns):
            st.error(f"❌ File thiếu cột. Cần {len(expected_columns)} cột, hiện có {df_cleaned.shape[1]}.")
            st.stop()
        df_cleaned = df_cleaned.iloc[:, :len(expected_columns)]
        df_cleaned.columns = expected_columns

        df_cleaned = df_cleaned[df_cleaned["Điện lực"].notna()]
        cols_to_num = ["1P_GT", "1P_TT", "3P_GT", "3P_TT", "TU", "TI", "Tổng công tơ", "Kế hoạch", "Tỷ lệ"]
        df_cleaned[cols_to_num] = df_cleaned[cols_to_num].apply(pd.to_numeric, errors='coerce')
        df_cleaned["Tỷ lệ"] = df_cleaned["Tỷ lệ"] * 100

        total_current = df_cleaned["Tổng công tơ"].sum()
        total_plan = df_cleaned["Kế hoạch"].sum()
        days_passed = (datetime.now() - datetime(2025, 1, 1)).days
        days_total = (datetime(2025, 9, 30) - datetime(2025, 1, 1)).days
        avg_per_day = total_current / days_passed
        forecast_total = avg_per_day * days_total
        forecast_ratio = forecast_total / total_plan

        danh_gia = "✅ ĐẠT kế hoạch" if forecast_ratio >= 1 else "❌ CHƯA ĐẠT kế hoạch"
        nhan_xet = f"Dự kiến đến 30/09/2025 sẽ hoàn thành khoảng {forecast_total:,.0f} thiết bị, tương đương {forecast_ratio*100:.2f}%. {danh_gia}"

        df_sorted = df_cleaned.sort_values(by="Tỷ lệ", ascending=False)
        top_3 = df_sorted.head(3)
        bot_3 = df_sorted.tail(3)

        chart_top = save_bar_chart(top_3, "Điện lực", "Tỷ lệ", "Top 3 Tỷ lệ cao")
        chart_bot = save_bar_chart(bot_3, "Điện lực", "Tỷ lệ", "Bottom 3 Tỷ lệ thấp")
        chart_all = save_overall_chart(df_sorted, "Điện lực", "Tỷ lệ", "Tổng hợp tỷ lệ hoàn thành")

        st.metric("Tổng đã thực hiện", f"{total_current:,}")
        st.metric("Kế hoạch", f"{total_plan:,}")
        st.metric("Dự báo đến 30/09/2025", f"{int(forecast_total):,}")
        st.metric("Tỷ lệ dự báo", f"{forecast_ratio*100:.2f}%")
        st.info(nhan_xet)

        st.subheader("📈 Biểu đồ")
        st.image(chart_top)
        st.image(chart_bot)
        st.image(chart_all)

        if st.button("📄 Xuất báo cáo Word", key="tab1_word"):
            tong_quan = {
                "Tổng công tơ đã thực hiện": f"{total_current:,}",
                "Kế hoạch giao": f"{total_plan:,}",
                "Tốc độ TB/ngày": f"{avg_per_day:.2f}",
                "Dự báo đến 30/09/2025": f"{int(forecast_total):,}",
                "Tỷ lệ dự báo": f"{forecast_ratio*100:.2f}%",
                "Đánh giá": danh_gia
            }
            word_file = export_word_report(tong_quan, df_cleaned[["Điện lực", "Tổng công tơ", "Kế hoạch", "Tỷ lệ"]], top_3, bot_3, [chart_top, chart_bot, chart_all], nhan_xet, filename="Bao_cao_HeThongDoDem.docx")
            st.download_button("📥 Tải báo cáo Word", data=word_file, file_name="Bao_cao_HeThongDoDem.docx")
