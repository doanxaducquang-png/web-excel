import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from utils import load_and_process_excel, find_excel_file, generate_exact_template_excel

st.set_page_config(
    page_title="Hệ Thống Quản Lý & Báo Cáo KH&CN",
    page_icon="📊",
    layout="wide"
)

excel_file_path = find_excel_file()

@st.cache_data(ttl=5)
def fetch_data(file_path):
    return load_and_process_excel(file_path)

df_74, df_25 = fetch_data(excel_file_path)

# Header
col_head1, col_head2 = st.columns([8, 2])
with col_head1:
    st.title("📌 HỆ THỐNG QUẢN LÝ & BÁO CÁO HỒ SƠ KH&CN")
with col_head2:
    if st.button("🔄 Cập nhật dữ liệu từ Excel"):
        st.cache_data.clear()
        st.rerun()

if df_74.empty and df_25.empty:
    st.error(f"❌ Không tìm thấy file Excel hoặc không đọc được dữ liệu từ file: **{excel_file_path}**")

st.markdown("---")

tab_dashboard, tab_74, tab_25, tab_export = st.tabs([
    "📊 Dashboard Trực Quan", 
    "📂 Phân Loại 74 Hồ Sơ", 
    "📑 Tổng Hợp 25 Nhiệm Vụ & Ý Kiến Phòng", 
    "🔍 Tra Cứu & Xuất Dữ Liệu"
])

# =============================================================================
# TAB 1: DASHBOARD TRỰC QUAN (TỔNG HỢP SONG SONG 74 VÀ 25 NHIỆM VỤ)
# =============================================================================
with tab_dashboard:
    st.header("📊 DASHBOARD TỔNG QUAN HỒ SƠ & NHIỆM VỤ KH&CN")
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Tổng hồ sơ đủ điều kiện xét", f"{len(df_74)} hồ sơ")
    col2.metric("Số nhiệm vụ tổng hợp trọng điểm", f"{len(df_25)} nhiệm vụ")
    
    total_kp_74 = df_74['Kinh phí đề xuất (trđ)'].sum() if not df_74.empty else 0
    total_kp_25 = pd.to_numeric(df_25['KINH PHÍ THỰC HIỆN (TRĐ)'], errors='coerce').sum() if not df_25.empty else 0
    
    col3.metric("Tổng kinh phí đề xuất (74 hồ sơ)", f"{total_kp_74:,.0f} tr.đ")
    col4.metric("Tổng kinh phí thực hiện (25 nhiệm vụ)", f"{total_kp_25:,.0f} tr.đ")
    
    st.markdown("---")
    
    # 1. BIỂU ĐỒ SHEET 74: PHÂN THEO LĨNH VỰC NGHIÊN CỨU (SẮP XẾP CHUẨN TỪ 01 ĐẾN 05)
    r1_c1, r1_c2 = st.columns(2)
    with r1_c1:
        st.subheader("1. Cơ cấu LĨNH VỰC NGHIÊN CỨU (Sheet 74)")
        if not df_74.empty:
            df_lv_count = df_74['Lĩnh vực'].value_counts().reset_index()
            df_lv_count.columns = ['Lĩnh vực', 'Số lượng']
            
            # SẮP XẾP THỨ TỰ CỐ ĐỊNH TỪ LĨNH VỰC 01 ĐẾN LĨNH VỰC 05
            df_lv_count = df_lv_count.sort_values(by='Lĩnh vực', ascending=True)
            
            fig_lv = px.pie(
                df_lv_count, 
                names='Lĩnh vực', 
                values='Số lượng', 
                hole=0.3, 
                color_discrete_sequence=px.colors.qualitative.Pastel,
                category_orders={"Lĩnh vực": sorted(df_lv_count['Lĩnh vực'].unique())}
            )
            fig_lv.update_layout(legend=dict(orientation="h", y=-0.2))
            st.plotly_chart(fig_lv, use_container_width=True)
        else:
            st.info("Chưa có dữ liệu Lĩnh vực")
            
    with r1_c2:
        st.subheader("2. Cơ cấu LOẠI HÌNH NHIỆM VỤ (Sheet 74)")
        if not df_74.empty:
            df_lh_count = df_74['Loại hình nhiệm vụ'].value_counts().reset_index()
            df_lh_count.columns = ['Loại hình', 'Số lượng']
            
            fig_lh = px.bar(df_lh_count, x='Loại hình', y='Số lượng', color='Loại hình', text='Số lượng')
            
            # ĐẢM BẢO SỐ HIỂN THỊ TRÊN CỘT ĐỨNG THẲNG 90° KHÔNG NẰM NGANG
            fig_lh.update_traces(textangle=0, textposition='outside')
            fig_lh.update_layout(showlegend=False, xaxis_title="", yaxis_title="Số lượng hồ sơ")
            st.plotly_chart(fig_lh, use_container_width=True)
        else:
            st.info("Chưa có dữ liệu Loại hình Sheet 74")

    st.markdown("---")
    
    # 2. PHÂN TÍCH CHO TỔNG HỢP 25 NHIỆM VỤ TRỌNG ĐIỂM
    r2_c1, r2_c2 = st.columns(2)
    with r2_c1:
        st.subheader("3. Cơ cấu LOẠI HÌNH NHIỆM VỤ (Sheet 25)")
        if not df_25.empty:
            df_lh25 = df_25['LOẠI HÌNH NHIỆM VỤ'].value_counts().reset_index()
            df_lh25.columns = ['Loại hình', 'Số lượng']
            
            fig_lh25 = px.bar(df_lh25, x='Loại hình', y='Số lượng', color='Loại hình', text='Số lượng')
            # CHỮ SỐ ĐỨNG THẲNG CÂN BẰNG
            fig_lh25.update_traces(textangle=0, textposition='outside')
            fig_lh25.update_layout(showlegend=False, xaxis_title="", yaxis_title="Số lượng nhiệm vụ")
            st.plotly_chart(fig_lh25, use_container_width=True)
        else:
            st.info("Chưa có dữ liệu Loại hình Sheet 25")
            
    with r2_c2:
        st.subheader("4. Top Đơn vị đề xuất chủ lực (Sheet 25 - Từ 2 NV trở lên)")
        if not df_25.empty:
            df_dv = df_25['ĐƠN VỊ ĐỀ XUẤT'].value_counts().reset_index()
            df_dv.columns = ['Đơn vị đề xuất', 'Số lượng']
            df_dv_filtered = df_dv[df_dv['Số lượng'] > 1].copy()
            
            if not df_dv_filtered.empty:
                fig_dv = px.bar(
                    df_dv_filtered, 
                    x='Đơn vị đề xuất', 
                    y='Số lượng', 
                    color='Đơn vị đề xuất', 
                    text='Số lượng'
                )
                # CHỮ SỐ ĐỨNG THẲNG VÀ NHÃN ĐỒNG BỘ
                fig_dv.update_traces(textangle=0, textposition='outside')
                fig_dv.update_layout(
                    showlegend=False, 
                    xaxis_title="", 
                    yaxis_title="Số lượng nhiệm vụ",
                    xaxis=dict(tickangle=0, automargin=True)
                )
                st.plotly_chart(fig_dv, use_container_width=True)
            else:
                st.info("Không có đơn vị nào đề xuất từ 02 nhiệm vụ trở lên.")

# =============================================================================
# TAB 2: SHEET PHÂN LOẠI 74
# =============================================================================
with tab_74:
    st.header("📂 DANH SÁCH PHÂN LOẠI 74 HỒ SƠ ĐỦ ĐIỀU KIỆN XẾT")
    if not df_74.empty:
        col_f1, col_f2, col_f3 = st.columns([2, 2, 3])
        with col_f1:
            list_lv = ["Tất cả"] + list(df_74['Lĩnh vực'].unique())
            selected_lv = st.selectbox("📂 Lọc Lĩnh vực:", list_lv)
        with col_f2:
            list_lh = ["Tất cả"] + list(df_74['Loại hình nhiệm vụ'].unique())
            selected_lh = st.selectbox("🎯 Lọc Loại hình:", list_lh)
        with col_f3:
            search_74 = st.text_input("🔑 Tìm kiếm (Tên / Đơn vị):")
            
        df_filtered_74 = df_74.copy()
        if selected_lv != "Tất cả":
            df_filtered_74 = df_filtered_74[df_filtered_74['Lĩnh vực'] == selected_lv]
        if selected_lh != "Tất cả":
            df_filtered_74 = df_filtered_74[df_filtered_74['Loại hình nhiệm vụ'] == selected_lh]
        if search_74:
            kw = search_74.lower()
            df_filtered_74 = df_filtered_74[
                df_filtered_74['Tên đề xuất/nhiệm vụ'].astype(str).str.lower().str.contains(kw) |
                df_filtered_74['Đơn vị đề xuất'].astype(str).str.lower().str.contains(kw)
            ]
            
        st.write(f"**Kết quả hiển thị:** {len(df_filtered_74)} / {len(df_74)} hồ sơ")
        st.dataframe(df_filtered_74, use_container_width=True, hide_index=True)

# =============================================================================
# TAB 3: SHEET TỔNG HỢP 25 & Ý KIẾN PHÒNG
# =============================================================================
with tab_25:
    st.header("📑 TỔNG HỢP 25 NHIỆM VỤ TRỌNG ĐIỂM & Ý KIẾN PHÒNG QLKH")
    st.info("💡 **Ghi chú:** Cột 'MỘT SỐ Ý KIẾN CỦA PHÒNG QLKH' chỉ tổng hợp duy nhất từ các cột D, F, H, J, M trong Sheet 'Ý kiến của phòng'. Ô không có ý kiến sẽ để trống.")
    
    if not df_25.empty:
        search_25 = st.text_input("🔑 Tìm kiếm nhiệm vụ / đơn vị (Sheet 25):")
        df_filtered_25 = df_25.copy()
        
        if search_25:
            kw = search_25.lower()
            df_filtered_25 = df_filtered_25[
                df_filtered_25['TÊN NHIỆM VỤ'].astype(str).str.lower().str.contains(kw) |
                df_filtered_25['ĐƠN VỊ ĐỀ XUẤT'].astype(str).str.lower().str.contains(kw)
            ]
            
        st.dataframe(
            df_filtered_25,
            use_container_width=True,
            hide_index=True,
            column_config={
                "MỘT SỐ Ý KIẾN CỦA PHÒNG QLKH": st.column_config.TextColumn("MỘT SỐ Ý KIẾN CỦA PHÒNG QLKH", width="large")
            }
        )
        
        st.markdown("---")
        st.subheader("🔍 Xem Chi Tiết Nội Dung Ý Kiến Theo Từng Nhiệm Vụ")
        selected_stt = st.selectbox("Chọn STT Nhiệm vụ để xem:", df_filtered_25['STT'].tolist())
        
        row_detail = df_filtered_25[df_filtered_25['STT'] == selected_stt].iloc[0]
        st.markdown(f"### 📌 STT {row_detail['STT']}: {row_detail['TÊN NHIỆM VỤ']}")
        st.markdown(f"**Đơn vị đề xuất:** {row_detail['ĐƠN VỊ ĐỀ XUẤT']} | **Kinh phí:** {row_detail['KINH PHÍ THỰC HIỆN (TRĐ)']} tr.đ | **Loại hình:** {row_detail['LOẠI HÌNH NHIỆM VỤ']}")
        
        c_dt1, c_dt2 = st.columns(2)
        with c_dt1:
            st.write("**Mục tiêu đề xuất:**", row_detail['MỤC TIÊU'])
            st.write("**Nội dung đề xuất:**", row_detail['NỘI DUNG'])
        with c_dt2:
            st.write("**Sản phẩm đề xuất:**", row_detail['SẢN PHẨM'])
            st.write("**Ghi chú:**", row_detail['GHI CHÚ'] or "Không có")
            
        st.markdown("#### 📝 Ý kiến gộp của Phòng QLKH (Tự động lọc từ Cột D, F, H, J, M):")
        if row_detail['MỘT SỐ Ý KIẾN CỦA PHÒNG QLKH']:
            st.warning(row_detail['MỘT SỐ Ý KIẾN CỦA PHÒNG QLKH'])
        else:
            st.success("Không có ý kiến điều chỉnh cho nhiệm vụ này.")

# =============================================================================
# TAB 4: XUẤT TÀI LIỆU CHUẨN MẪU EXCEL GỐC
# =============================================================================
with tab_export:
    st.header("📥 XUẤT BÁO CÁO EXCEL ĐÚNG CHUẨN NGUYÊN BẢN MẪU GỐC")
    st.info("📌 File xuất ra giữ nguyên toàn bộ cấu trúc các sheet, định dạng khung, màu sắc của file Excel gốc và đã cập nhật cột ý kiến tổng hợp.")
    
    excel_bytes = generate_exact_template_excel(excel_file_path)
    
    st.download_button(
        label="📥 Tải File Báo Cáo Excel Chuẩn Mẫu Gốc (Đã Cập Nhật Ý Kiến)",
        data=excel_bytes,
        file_name=f"3. PHÂN LOẠI HỒ SƠ ĐỦ ĐIỀU KIỆN XÉT_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )