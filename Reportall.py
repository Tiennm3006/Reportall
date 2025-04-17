# [Giữ nguyên phần import và định nghĩa hàm xử lý như hiện tại]

# ---------- TAB 1: PHÂN TÍCH HỆ THỐNG ĐO ĐẾM ---------- #
with tab1:
    uploaded_file = st.file_uploader("📤 Tải lên file Excel chứa sheet 'Tong hop luy ke'", type=["xlsx"], key="kiemtra")

    if uploaded_file:
        df = pd.read_excel(uploaded_file, sheet_name="Tong hop luy ke", header=None)
        df_cleaned = df.iloc[4:].copy()
        df_cleaned.columns = ["STT", "Điện lực", "1P_GT", "1P_TT", "3P_GT", "3P_TT", "TU", "TI", "Tổng công tơ", "Kế hoạch", "Tỷ lệ"]
        df_cleaned = df_cleaned[df_cleaned["Điện lực"].notna()]
        cols_to_num = ["1P_GT", "1P_TT", "3P_GT", "3P_TT", "TU", "TI", "Tổng công tơ", "Kế hoạch", "Tỷ lệ"]
        df_cleaned[cols_to_num] = df_cleaned[cols_to_num].apply(pd.to_numeric, errors='coerce')
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
