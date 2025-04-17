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
